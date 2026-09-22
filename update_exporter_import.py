import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Replace import_overtimes_from_excel in exporter.py
start_marker = "def import_overtimes_from_excel("
end_marker = "    return {\n        \"success\": True,\n        \"total_rows\": total_rows,\n        \"processed_count\": processed_count,\n        \"created_count\": created_count,\n        \"updated_count\": updated_count,\n        \"skipped_count\": skipped_count,\n        \"ignored_teams_count\": ignored_teams_count,\n        \"errors\": errors\n    }"

new_func = """def import_overtimes_from_excel(file_bytes: bytes, target_teams: list = None, admin_emp_id: str = None) -> dict:
    \"\"\"
    엑셀 파일(특근현황_일자별개인별정산의 1번째 시트 또는 특근목록)에서 데이터를 읽어와
    선택된 부서(target_teams)의 특근 정보를 DB(overtimes)에 갱신/추가합니다.
    
    - 선택된 부서만 처리하며 선택되지 않은 부서의 정보는 100% 보존합니다.
    - 동일 사원/동일 날짜/동일 구분의 기록은 중복 생성을 방지(기존 데이터 갱신)합니다.
    - 실패/오류 항목은 사유별 상세 로그를 반환합니다.
    \"\"\"
    import openpyxl
    import re
    from database import get_db_connection, log_audit

    if isinstance(file_bytes, (bytes, bytearray)):
        bio = io.BytesIO(file_bytes)
    else:
        bio = file_bytes

    try:
        wb = openpyxl.load_workbook(bio, data_only=True)
    except Exception as e:
        return {
            "success": False,
            "message": f"엑셀 파일을 열 수 없습니다: {str(e)}",
            "total_rows": 0, "processed_count": 0, "created_count": 0, "updated_count": 0,
            "skipped_count": 0, "ignored_teams_count": 0, "errors": []
        }

    ws = wb.worksheets[0]

    # 1. 헤더 행 찾기
    header_row_idx = None
    headers_map = {}

    for r in range(1, min(15, ws.max_row + 1)):
        row_vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        row_str = " ".join(row_vals)
        if any(kw in row_str for kw in ["신청자", "사번", "성명", "근무기간", "특근일자", "휴일날짜", "분류", "특근분류", "소속팀"]):
            header_row_idx = r
            for c_idx, val in enumerate(row_vals, 1):
                clean_v = val.replace(" ", "").replace("\n", "")
                headers_map[clean_v] = c_idx
            break

    if not header_row_idx:
        return {
            "success": False,
            "message": "엑셀 파일에서 올바른 헤더(사번/신청자, 소속팀, 휴일날짜/특근일자/근무기간 등)를 찾을 수 없습니다.",
            "total_rows": 0, "processed_count": 0, "created_count": 0, "updated_count": 0,
            "skipped_count": 0, "ignored_teams_count": 0, "errors": []
        }

    # 헤더 인덱스 매핑 찾기 유틸리티
    def find_col(possible_names):
        for k, col in headers_map.items():
            for name in possible_names:
                if name in k or k in name:
                    return col
        return None

    emp_col = find_col(["사번", "사원번호", "신청자", "emp_id"])
    name_col = find_col(["성명", "이름", "user_name"])
    team_col = find_col(["소속팀", "소속", "부서", "team"])
    cat_col = find_col(["특근분류", "분류", "특근구분", "구분", "category"])
    period_col = find_col(["휴일날짜", "휴일일자", "특근일자", "근무기간", "특근기간", "날짜", "일자"])
    start_col = find_col(["시작일", "시작일자", "start_date"])
    end_col = find_col(["종료일", "종료일자", "end_date"])
    sub_date_col = find_col(["대체휴일사용일", "대체휴일", "대체휴무사용일", "sub_holiday_date"])
    proj_col = find_col(["프로젝트번호", "프로젝트", "project_no"])
    loc_col = find_col(["근무장소", "장소", "location"])
    reason_col = find_col(["특근사유", "사유", "reason"])
    bonus_col = find_col(["보너스부여", "보너스", "bonus_granted"])
    confirm_col = find_col(["확인(승인)", "확인승인", "승인", "확인", "is_confirmed"])
    confirm_by_col = find_col(["확인자", "confirmed_by"])
    pre_deduct_col = find_col(["사전차감", "is_pre_deduct"])

    # 필터 타겟 부서 정제
    valid_target_teams = set()
    if target_teams:
        if isinstance(target_teams, str):
            valid_target_teams = set(t.strip() for t in target_teams.split(",") if t.strip())
        elif isinstance(target_teams, (list, set, tuple)):
            valid_target_teams = set(t.strip() for t in target_teams if t and str(t).strip())

    total_rows = 0
    processed_count = 0
    created_count = 0
    updated_count = 0
    skipped_count = 0
    ignored_teams_count = 0
    errors = []

    conn = get_db_connection()
    cursor = conn.cursor()

    # 사원 DB 정보 캐시 (이름/팀 보완용)
    cursor.execute("SELECT emp_id, name, team FROM users")
    db_users = {row["emp_id"]: dict(row) for row in cursor.fetchall()}

    def norm_date(d_str):
        if not d_str: return ""
        d_str = str(d_str).replace(".", "-").replace("/", "-").strip()
        m = re.search(r"(\d{4})[-_](\d{1,2})[-_](\d{1,2})", d_str)
        if m:
            return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        return d_str

    for r in range(header_row_idx + 1, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        str_vals = [str(v or "").strip() for v in row_vals]

        # 합계 / 공백 행 건너뛰기
        if not any(str_vals) or any(kw in str_vals[0] for kw in ["합계", "총계", "전체", "순번"]):
            continue

        # 1. 사번 / 성명 추출
        raw_emp = str_vals[emp_col - 1] if emp_col and emp_col <= len(str_vals) else ""
        raw_name = str_vals[name_col - 1] if name_col and name_col <= len(str_vals) else ""

        emp_id = ""
        name = ""

        if "(" in raw_emp and ")" in raw_emp:
            m = re.search(r"^(.*?)\((.*?)\)", raw_emp)
            if m:
                name = m.group(1).strip()
                emp_id = m.group(2).strip()
        else:
            emp_id = raw_emp.strip()
            name = raw_name.strip()

        if emp_id in db_users:
            if not name:
                name = db_users[emp_id]["name"]
        elif not emp_id and name:
            for uid, uinfo in db_users.items():
                if uinfo["name"] == name:
                    emp_id = uid
                    break

        if not emp_id or emp_id == "-" or emp_id.isdigit() == False and len(emp_id) < 2 and not any(c.isalnum() for c in emp_id):
            # 유효하지 않은 사번 행 패스
            if any(str_vals):
                skipped_count += 1
            continue

        total_rows += 1

        # 2. 소속팀 추출 및 부서 필터링
        team = str_vals[team_col - 1] if team_col and team_col <= len(str_vals) else ""
        if not team and emp_id in db_users:
            team = db_users[emp_id]["team"]

        db_user_team = db_users.get(emp_id, {}).get("team", "")

        if valid_target_teams:
            if team not in valid_target_teams and db_user_team not in valid_target_teams:
                ignored_teams_count += 1
                skipped_count += 1
                continue

        # 3. 날짜 추출
        start_date = ""
        end_date = ""

        if period_col and period_col <= len(str_vals):
            p_val = str_vals[period_col - 1]
            if "~" in p_val:
                parts = p_val.split("~")
                start_date = parts[0].strip()
                end_date = parts[1].strip()
            elif p_val:
                start_date = p_val.strip()
                end_date = p_val.strip()

        if not start_date and start_col and start_col <= len(str_vals):
            raw_s = ws.cell(r, start_col).value
            if isinstance(raw_s, datetime):
                start_date = raw_s.strftime("%Y-%m-%d")
            elif raw_s:
                start_date = str(raw_s).strip().split("T")[0].split(" ")[0]

        if not end_date and end_col and end_col <= len(str_vals):
            raw_e = ws.cell(r, end_col).value
            if isinstance(raw_e, datetime):
                end_date = raw_e.strftime("%Y-%m-%d")
            elif raw_e:
                end_date = str(raw_e).strip().split("T")[0].split(" ")[0]

        if not end_date and start_date:
            end_date = start_date

        start_date = norm_date(start_date)
        end_date = norm_date(end_date)

        if not start_date or not end_date:
            errors.append({
                "row": r,
                "name": name,
                "emp_id": emp_id,
                "reason": "근무기간 / 휴일날짜 형식을 읽을 수 없음"
            })
            skipped_count += 1
            continue

        # 4. 기타 속성 추출
        category = str_vals[cat_col - 1] if cat_col and cat_col <= len(str_vals) else "일반휴일"
        if not category or category == "-": category = "일반휴일"

        project = str_vals[proj_col - 1] if proj_col and proj_col <= len(str_vals) else ""
        if project == "-": project = ""
        location = str_vals[loc_col - 1] if loc_col and loc_col <= len(str_vals) else ""
        if location == "-": location = ""
        reason = str_vals[reason_col - 1] if reason_col and reason_col <= len(str_vals) else ""
        if reason == "-": reason = ""

        raw_b = str_vals[bonus_col - 1] if bonus_col and bonus_col <= len(str_vals) else "0"
        bonus_point = 1 if any(kw in raw_b for kw in ["1", "예", "Y", "True", "O", "부여"]) and "미부여" not in raw_b else 0

        raw_p = str_vals[pre_deduct_col - 1] if pre_deduct_col and pre_deduct_col <= len(str_vals) else "0"
        pre_deduct_point = 1 if any(kw in raw_p for kw in ["1", "예", "Y", "True", "O"]) else 0

        raw_confirm = str_vals[confirm_col - 1] if confirm_col and confirm_col <= len(str_vals) else ""
        is_confirmed = 1 if ("확인완료" in raw_confirm or "승인" in raw_confirm or raw_confirm == "1" or raw_confirm.lower() == "true") else 1

        confirmed_by = str_vals[confirm_by_col - 1] if confirm_by_col and confirm_by_col <= len(str_vals) else "관리자(엑셀업로드)"
        if not confirmed_by or confirmed_by == "-": confirmed_by = "관리자(엑셀업로드)"

        sub_holiday_date = ""
        sub_holiday_used = 0
        if sub_date_col and sub_date_col <= len(str_vals):
            raw_sub = str_vals[sub_date_col - 1]
            if raw_sub and raw_sub != "-":
                sub_holiday_date = norm_date(raw_sub)
                if sub_holiday_date:
                    sub_holiday_used = 1

        # 5. DB 중복 검사 및 갱신 / 신규 등록
        # (1) 정밀 일치
        cursor.execute(\"\"\"
            SELECT id FROM overtimes
            WHERE emp_id = ? AND start_date = ? AND end_date = ? AND category = ?
        \"\"\", (emp_id, start_date, end_date, category))
        exist_row = cursor.fetchone()

        # (2) 기간 포함 일치
        if not exist_row:
            cursor.execute(\"\"\"
                SELECT id FROM overtimes
                WHERE emp_id = ? AND start_date <= ? AND end_date >= ?
            \"\"\", (emp_id, start_date, start_date))
            exist_row = cursor.fetchone()

        # (3) 동일 시작일 일치
        if not exist_row:
            cursor.execute(\"\"\"
                SELECT id FROM overtimes
                WHERE emp_id = ? AND start_date = ?
            \"\"\", (emp_id, start_date))
            exist_row = cursor.fetchone()

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if exist_row:
            ot_id = exist_row["id"]
            cursor.execute(\"\"\"
                UPDATE overtimes
                SET user_name = ?, team = ?, category = ?, start_date = ?, end_date = ?,
                    project_no = ?, location = ?, reason = ?, bonus_granted = ?, is_pre_deduct = ?,
                    is_confirmed = ?, confirmed_by = ?,
                    sub_holiday_date = CASE WHEN ? != '' THEN ? ELSE sub_holiday_date END,
                    sub_holiday_used = CASE WHEN ? != '' THEN 1 ELSE sub_holiday_used END,
                    updated_at = ?
                WHERE id = ?
            \"\"\", (name, team or db_user_team, category, start_date, end_date,
                  project, location, reason, bonus_point, pre_deduct_point,
                  is_confirmed, confirmed_by,
                  sub_holiday_date, sub_holiday_date,
                  sub_holiday_date,
                  now_str, ot_id))
            updated_count += 1
            log_audit(ot_id, "EXCEL_IMPORT_UPDATE", admin_emp_id or "ADMIN", "관리자", None, {"emp_id": emp_id, "start_date": start_date})
        else:
            cursor.execute(\"\"\"
                INSERT INTO overtimes (
                    emp_id, user_name, team, category, start_date, end_date,
                    project_no, location, reason, bonus_granted, is_pre_deduct,
                    is_confirmed, confirmed_by, sub_holiday_date, sub_holiday_used,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            \"\"\", (emp_id, name, team or db_user_team, category, start_date, end_date,
                  project, location, reason, bonus_point, pre_deduct_point,
                  is_confirmed, confirmed_by, sub_holiday_date, sub_holiday_used,
                  now_str, now_str))
            ot_id = cursor.lastrowid
            created_count += 1
            log_audit(ot_id, "EXCEL_IMPORT_CREATE", admin_emp_id or "ADMIN", "관리자", None, {"emp_id": emp_id, "start_date": start_date})

        processed_count += 1

    conn.commit()
    conn.close()

    return {
        "success": True,
        "total_rows": total_rows,
        "processed_count": processed_count,
        "created_count": created_count,
        "updated_count": updated_count,
        "skipped_count": skipped_count,
        "ignored_teams_count": ignored_teams_count,
        "errors": errors
    }"""

idx_start = code.find(start_marker)
idx_end = code.find(end_marker) + len(end_marker)

if idx_start != -1 and idx_end != -1:
    new_code = code[:idx_start] + new_func + code[idx_end:]
    with open('exporter.py', 'w', encoding='utf-8') as f:
        f.write(new_code)
    print("✓ Successfully updated import_overtimes_from_excel in exporter.py")
else:
    print(f"❌ Marker not found: idx_start={idx_start}, idx_end={idx_end}")
