import io
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.table import Table, TableStyleInfo


# ===========================================================================
#  공통 스타일 팩토리 (openpyxl 엔진 전용)
# ===========================================================================

FONT_FAMILY = "맑은 고딕"

def _make_styles():
    """공통 스타일 dict 반환"""
    thin = Side(style='thin', color='CBD5E1')
    medium = Side(style='medium', color='94A3B8')
    thick = Side(style='thick', color='1E3A8A')

    return {
        "title_font":    Font(name=FONT_FAMILY, size=16, bold=True, color="0F172A"),
        "subtitle_font": Font(name=FONT_FAMILY, size=9,  italic=True, color="64748B"),
        "header_font":   Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF"),
        "data_font":     Font(name=FONT_FAMILY, size=9,  color="1E293B"),
        "total_font":    Font(name=FONT_FAMILY, size=10, bold=True, color="0F172A"),
        "hl_font":       Font(name=FONT_FAMILY, size=9,  bold=True, color="0D9488"),  # 실특근 강조
        "conf_font":     Font(name=FONT_FAMILY, size=9,  bold=True, color="15803D"),
        "pend_font":     Font(name=FONT_FAMILY, size=9,  color="B45309"),

        "fill_navy":    PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid"),
        "fill_teal":    PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid"),
        "fill_slate":   PatternFill(start_color="334155", end_color="334155", fill_type="solid"),
        "fill_alt":     PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"),
        "fill_total":   PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid"),
        "fill_conf":    PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
        "fill_pend":    PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid"),
        "fill_hl":      PatternFill(start_color="CCFBF1", end_color="CCFBF1", fill_type="solid"),  # 실특근 행 강조
        "fill_none":    PatternFill(fill_type=None),

        "border_thin":   Border(left=thin, right=thin, top=thin, bottom=thin),
        "border_total":  Border(left=medium, right=medium, top=medium,
                                bottom=Side(style='double', color='0F172A')),
        "border_header": Border(left=medium, right=medium, top=medium, bottom=medium),

        "align_center": Alignment(horizontal="center", vertical="center"),
        "align_left":   Alignment(horizontal="left",   vertical="center", indent=1),
        "align_right":  Alignment(horizontal="right",  vertical="center"),
        "align_wrap":   Alignment(horizontal="center", vertical="center", wrap_text=True),
    }


def _auto_col_width(ws, skip_rows=(1, 2)):
    """모든 시트 컬럼 너비 자동 조정 (타이틀 행 제외)"""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in skip_rows:
                continue
            val_str = str(cell.value or "")
            if val_str.startswith("="):
                val_str = "999999"
            val_len = sum(2 if ord(c) > 128 else 1 for c in val_str)
            if val_len > max_len:
                max_len = val_len
        ws.column_dimensions[col_letter].width = max(max_len + 4, 10)


def _write_title(ws, text: str, subtitle: str, merge_cols: int, s: dict):
    """타이틀 + 서브타이틀 2행 공통 기입"""
    ws.merge_cells(f"A1:{get_column_letter(merge_cols)}1")
    tc = ws["A1"]
    tc.value = text
    tc.font = s["title_font"]
    tc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 34

    ws.merge_cells(f"A2:{get_column_letter(merge_cols)}2")
    sc = ws["A2"]
    sc.value = subtitle
    sc.font = s["subtitle_font"]
    sc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18


def _write_header_row(ws, row: int, headers: list, fills: list, s: dict):
    """헤더 행 기입 (fills: 컬럼별 fill 또는 None = navy 기본)"""
    ws.row_dimensions[row].height = 28
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.font = s["header_font"]
        cell.fill = fills[col_idx - 1] if fills[col_idx - 1] else s["fill_navy"]
        cell.alignment = s["align_wrap"]
        cell.border = s["border_header"]


def _write_data_row(ws, row: int, values: list, s: dict, is_even: bool,
                    center_cols: set, right_cols: set,
                    num_fmt_map: dict = None, hl_cols: set = None):
    """데이터 행 기입"""
    fill = s["fill_alt"] if is_even else s["fill_none"]
    ws.row_dimensions[row].height = 22
    for col_idx, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col_idx, value=val)
        cell.font = s["data_font"]
        cell.fill = fill
        cell.border = s["border_thin"]
        if col_idx in center_cols:
            cell.alignment = s["align_center"]
        elif col_idx in right_cols:
            cell.alignment = s["align_right"]
        else:
            cell.alignment = s["align_left"]
        if num_fmt_map and col_idx in num_fmt_map:
            cell.number_format = num_fmt_map[col_idx]
        if hl_cols and col_idx in hl_cols:
            cell.font = s["hl_font"]


def _write_total_row(ws, row: int, label: str, label_merge_end: int,
                     sum_cols: list, s: dict, data_start_row: int = 4):
    """합계 행 기입 (SUM 수식 + 더블 밑줄 테두리)"""
    ws.row_dimensions[row].height = 26
    # 라벨 병합
    if label_merge_end > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=label_merge_end)
    lc = ws.cell(row=row, column=1, value=label)
    lc.font = s["total_font"]
    lc.alignment = s["align_center"]
    for c in range(1, label_merge_end + 1):
        ws.cell(row=row, column=c).fill = s["fill_total"]
        ws.cell(row=row, column=c).border = s["border_total"]

    # SUM 수식
    for col_i, num_fmt, is_hl in sum_cols:
        f_val = f"=SUM({get_column_letter(col_i)}{data_start_row}:{get_column_letter(col_i)}{row - 1})"
        c = ws.cell(row=row, column=col_i, value=f_val)
        c.font = s["hl_font"] if is_hl else s["total_font"]
        c.fill = s["fill_hl"] if is_hl else s["fill_total"]
        c.border = s["border_total"]
        c.alignment = s["align_right"]
        c.number_format = num_fmt


# ===========================================================================
#  보조 함수
# ===========================================================================

def expand_records_by_holiday_date(records: list) -> list:
    """다일 특근 신청 내역을 개별 휴일 날짜 단위로 전개 (요구사항 27-1)"""
    WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]
    expanded = []

    for r in records:
        s_date_str = r.get("start_date", "")
        e_date_str = r.get("end_date", "") or s_date_str

        try:
            d1 = datetime.strptime(s_date_str, "%Y-%m-%d")
            d2 = datetime.strptime(e_date_str, "%Y-%m-%d")
            if d2 < d1:
                d2 = d1
        except Exception:
            d1 = datetime.now()
            d2 = d1

        total_days = (d2 - d1).days + 1
        sub_used = float(r.get("sub_holiday_used", 0.0) or 0.0)
        sub_date = r.get("sub_holiday_date", "") or "-"
        is_conf = bool(r.get("is_confirmed", 0))
        status_text = "확인완료" if is_conf else "승인대기"

        curr = d1
        day_idx = 0
        while curr <= d2:
            date_str = curr.strftime("%Y-%m-%d")
            w_idx = curr.weekday()
            w_str = WEEKDAYS[w_idx]

            cur_sub_date = sub_date if (day_idx == 0 and sub_date != "-") else ("-" if total_days > 1 else sub_date)
            cur_sub_used = sub_used if (day_idx == 0) else 0.0
            cur_is_pre = int(r.get("is_pre_deduct", 0) or 0) if day_idx == 0 else 0

            expanded.append({
                "holiday_date": date_str,
                "weekday": w_str,
                "emp_id": r.get("emp_id", ""),
                "user_name": r.get("user_name", ""),
                "team": r.get("team", ""),
                "category": r.get("category", ""),
                "bonus_granted": r.get("bonus_granted", 0),
                "bonus_text": "부여(O)" if r.get("bonus_granted", 0) == 1 else "미부여(-)",
                "holiday_days": 1,
                "sub_holiday_date": cur_sub_date,
                "sub_holiday_used": cur_sub_used,
                "is_pre_deduct": cur_is_pre,
                "pre_deduct_text": "O" if cur_is_pre else "-",
                "trip_start_date": r.get("trip_start_date", "") or "-" if day_idx == 0 else "-",
                "trip_end_date": r.get("trip_end_date", "") or "-" if day_idx == 0 else "-",
                "project_no": r.get("project_no", "") or "-",
                "location": r.get("location", "") or "-",
                "reason": r.get("reason", "") or "-",
                "is_confirmed": is_conf,
                "status_text": status_text,
                "confirmed_by": r.get("confirmed_by", "") or "-",
                "created_at": r.get("created_at", "")
            })
            curr += timedelta(days=1)
            day_idx += 1

    expanded.sort(key=lambda x: (x["holiday_date"], x["team"], x["user_name"]))
    return expanded


def aggregate_user_holidays(records: list, user_positions: dict = None) -> list:
    """개인별 휴일수(대체근무, 법정휴일, 일반휴일, 대체휴무, 사전차감, 최종실특근) 집계"""
    user_positions = user_positions or {}
    user_map = {}
    user_trips = {}  # emp_id -> [(start, end)]

    # 1차: 대체휴무 출장기간 수집
    for r in records:
        emp_id = r.get("emp_id", "")
        cat = (r.get("category", "") or "").strip()
        t_s = (r.get("trip_start_date", "") or "").strip()
        t_e = (r.get("trip_end_date", "") or "").strip()
        if cat in ["대체휴무", "대체휴일"] and t_s and t_e:
            if emp_id not in user_trips:
                user_trips[emp_id] = []
            user_trips[emp_id].append((t_s, t_e))

    for r in records:
        emp_id = r.get("emp_id", "")
        name = r.get("user_name", "")
        team = r.get("team", "")
        cat = r.get("category", "")
        # 대체휴일 → 대체휴무 표기 정규화
        if cat == "대체휴일":
            cat = "대체휴무"
        sub_used = float(r.get("sub_holiday_used", 0.0) or 0.0)
        is_pre = int(r.get("is_pre_deduct", 0) or 0)
        s_date = r.get("start_date", "")
        e_date = r.get("end_date") or s_date

        try:
            d1 = datetime.strptime(s_date, "%Y-%m-%d")
            d2 = datetime.strptime(e_date, "%Y-%m-%d")
            days = max(1, (d2 - d1).days + 1)
        except Exception:
            days = 1

        if emp_id not in user_map:
            user_map[emp_id] = {
                "emp_id": emp_id,
                "name": name,
                "team": team,
                "position": user_positions.get(emp_id, "-"),
                "sub_work_days": 0,
                "legal_holiday_days": 0,
                "normal_holiday_days": 0,
                "sub_holiday_days": 0,       # 대체휴무 일수
                "pre_deduct_count": 0,       # 총 사전차감 횟수
                "trip_pre_deduct_count": 0,  # 출장기간내 사전차감 횟수
                "total_days": 0,
                "sub_holiday_used": 0.0,
                "actual_overtime_days": 0.0,
                "bonus_count": 0,            # 보너스 부여 건수
                "records_count": 0
            }

        u = user_map[emp_id]
        u["records_count"] += 1
        u["sub_holiday_used"] += sub_used

        if cat == "대체근무":
            u["sub_work_days"] += days
        elif cat == "법정휴일":
            u["legal_holiday_days"] += days
        elif cat == "대체휴무":
            u["sub_holiday_days"] += days
        else:  # 일반휴일 + 기타
            u["normal_holiday_days"] += days

        # 요구사항 1-1: 총 특근일수 = 대체근무 + 법정휴일 + 일반휴일 (대체휴무 제외)
        u["total_days"] = int(u["sub_work_days"] + u["legal_holiday_days"] + u["normal_holiday_days"])

        # 요구사항 1-3: 보너스 개수 집계
        if int(r.get("bonus_granted", 0) or 0) == 1:
            u["bonus_count"] += 1

        if is_pre:
            u["pre_deduct_count"] += 1
            # 출장기간 내 포함 여부 확인
            trips = user_trips.get(emp_id, [])
            for ts, te in trips:
                if not (e_date < ts or s_date > te):
                    u["trip_pre_deduct_count"] += 1
                    break

        # 최종 실특근일 = 일반특근 - 사전차감 - (대체휴무 - 대체휴무시 작성한 출장기간 이내의 사전차감)
        u["actual_overtime_days"] = max(0.0, round(float(u["normal_holiday_days"] - u["pre_deduct_count"] - (u["sub_holiday_days"] - u["trip_pre_deduct_count"])), 1))
        # 사전차감 잔여수 = 총사전차감수 - 출장기간내사전차감수 (요구사항 7)
        u["pre_deduct_remaining"] = max(0, u["pre_deduct_count"] - u["trip_pre_deduct_count"])
        # ★ 최종 실특근일+보너스 = 최종 실특근일 + 보너스 건수
        u["actual_overtime_with_bonus"] = max(0.0, round(float(u["actual_overtime_days"] + u["bonus_count"]), 1))

    user_list = list(user_map.values())
    user_list.sort(key=lambda x: (x["team"], x["name"]))
    return user_list


def aggregate_team_holidays(user_summaries: list) -> list:
    """부서(팀)별 합산 요약 집계"""
    team_map = {}
    for u in user_summaries:
        t = u["team"]
        if t not in team_map:
            team_map[t] = {
                "team": t,
                "member_count": 0,
                "sub_work_days": 0,
                "legal_holiday_days": 0,
                "normal_holiday_days": 0,
                "sub_holiday_days": 0,
                "total_days": 0,
                "actual_overtime_days": 0.0,
                "bonus_count": 0,
                "actual_overtime_with_bonus": 0.0,
                "records_count": 0
            }
        tm = team_map[t]
        tm["member_count"] += 1
        tm["sub_work_days"] += u["sub_work_days"]
        tm["legal_holiday_days"] += u["legal_holiday_days"]
        tm["normal_holiday_days"] += u["normal_holiday_days"]
        tm["sub_holiday_days"] += u.get("sub_holiday_days", 0)
        tm["total_days"] += u["total_days"]
        tm["actual_overtime_days"] = max(0.0, round(tm["actual_overtime_days"] + u["actual_overtime_days"], 1))
        tm["bonus_count"] += u.get("bonus_count", 0)
        tm["actual_overtime_with_bonus"] = max(0.0, round(tm.get("actual_overtime_with_bonus", 0.0) + u.get("actual_overtime_with_bonus", 0.0), 1))
        tm["records_count"] += u["records_count"]

    team_list = list(team_map.values())
    team_list.sort(key=lambda x: x["team"])
    return team_list


# ===========================================================================
#  메인 엑셀 생성 함수 1: 특근 신청 내역 전체 (4개 시트)
# ===========================================================================

def generate_overtime_excel(records: list, user_positions: dict = None, period_str: str = "") -> io.BytesIO:
    """
    특근 목록 데이터를 openpyxl로 4개 시트 워크북 생성 (요구사항 27-1, 27-2, 27-3)
    - 시트1: [휴일일자별_특근현황]  개별 날짜 전개 (일수=순수 int 1)
    - 시트2: [개인별_휴일합산_정산표]  인원별 분류별 집계 + SUM 공식 + 취합날짜/수식표기 (요구사항 8, 9, 10)
    - 시트3: [특근신청_전체원장]  원장 (일수=순수 int)
    - 시트4: [부서별_요약]  팀별 집계
    """
    s = _make_styles()
    wb = Workbook()
    now_str = datetime.now().strftime("%Y년 %m월 %d일 %H:%M")

    if not period_str and records:
        all_starts = [r.get("start_date") for r in records if r.get("start_date")]
        all_ends = [r.get("end_date") or r.get("start_date") for r in records if r.get("start_date")]
        if all_starts and all_ends:
            min_s = min(all_starts)
            max_e = max(all_ends)
            period_str = f"{min_s} ~ {max_e}"
    if not period_str:
        period_str = "전체 기간"

    # ─────────────────────────────────────────────
    # 시트 1: 휴일일자별_특근현황
    # ─────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "휴일일자별_특근현황"

    _write_title(ws1,
                 "일자별 휴일 특근(초과근무) 상세 현황",
                 f"취합 기간: {period_str}  |  출력일시: {now_str}  |  신청 내역을 개별 휴일 날짜 단위로 전개  |  일수 단위: 순수 숫자",
                 16, s)

    headers1 = [
        "순번", "휴일날짜", "요일", "사번", "성명", "소속팀", "특근분류", "보너스 부여",
        "휴일일수", "대체휴무 사용일", "대체휴가 사용일수", "사전차감",
        "출장기간(시작)", "출장기간(종료)",
        "프로젝트 번호", "근무 장소", "특근 사유", "확인(승인)", "확인자", "신청일시"
    ]
    fills1 = [None] * len(headers1)
    _write_header_row(ws1, 3, headers1, fills1, s)

    # 행 고정 (헤더까지)
    ws1.freeze_panes = "A4"

    expanded = expand_records_by_holiday_date(records)
    r_idx = 4
    total_hdays = 0
    total_sub_used_sum = 0.0

    for idx, r in enumerate(expanded, 1):
        h_val = int(r["holiday_days"])   # 순수 정수 1
        sub_v = float(r["sub_holiday_used"])  # 순수 float
        total_hdays += h_val
        total_sub_used_sum += sub_v
        is_conf = r["is_confirmed"]

        row_data = [
            idx,
            r["holiday_date"],
            r["weekday"],
            r["emp_id"],
            r["user_name"],
            r["team"],
            r["category"],
            r.get("bonus_text", "미부여(-)"),
            h_val,
            r["sub_holiday_date"],
            sub_v,
            r.get("pre_deduct_text", "-"),
            r.get("trip_start_date", "-"),
            r.get("trip_end_date", "-"),
            r["project_no"],
            r["location"],
            r["reason"],
            r["status_text"],
            r["confirmed_by"],
            r["created_at"]
        ]

        center_c = {1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 13, 14, 18, 19, 20}
        right_c  = {9, 11}
        num_fmt  = {9: "#,##0", 11: "0.0"}

        _write_data_row(ws1, r_idx, row_data, s, r_idx % 2 == 0, center_c, right_c, num_fmt)

        # 승인 상태 셀 개별 스타일
        conf_cell = ws1.cell(row=r_idx, column=18)
        conf_cell.fill = s["fill_conf"] if is_conf else s["fill_pend"]
        conf_cell.font = s["conf_font"] if is_conf else s["pend_font"]

        r_idx += 1

    # 합계 행
    _write_total_row(ws1, r_idx, f"총 합계 ({len(expanded)}건)", 8,
                     [(9, "#,##0", False), (11, "0.0", False)], s)
    ws1.cell(row=r_idx, column=10, value="-").alignment = s["align_center"]
    ws1.cell(row=r_idx, column=10).fill = s["fill_total"]
    ws1.cell(row=r_idx, column=10).border = s["border_total"]
    for c in range(12, 21):
        ws1.cell(row=r_idx, column=c, value="").fill = s["fill_total"]
        ws1.cell(row=r_idx, column=c).border = s["border_total"]

    # ─────────────────────────────────────────────
    # 시트 2: 개인별_휴일합산_정산표
    # ─────────────────────────────────────────────
    ws2 = wb.create_sheet(title="개인별_휴일합산_정산표")
    _write_title(ws2,
                 "개인별 특근 휴일수 세부 합산 및 최종 실특근일 정산표",
                 f"취합 기간: {period_str}  |  취합 일시: {now_str}  |  수식: [총특근일수] = [대체근무]+[법정휴일]+[일반휴일]  /  [★최종 실특근일] = [일반특근] - [사전차감] - ([대체휴무] - [출장내사전차감])  /  [★최종 실특근일+보너스] = [최종실특근일] + [보너스]",
                 16, s)

    headers2 = [
        "순번", "사원번호", "성명", "소속팀", "직급",
        "대체근무 (일)", "법정휴일 (일)", "일반휴일 (일)", "대체휴무 (일)",
        "총 특근일수\n(대체+법정+일반)", "총 사전차감 (회)", "사전차감 잔여수\n(총사전차감 - 출장내사전차감)",
        "★ 최종 실특근일\n(일반특근-사전차감-(대휴-출장내차감)) [일]", "보너스 부여 (건)",
        "★ 최종 실특근일+보너스\n(실특근 + 보너스) [일]", "신청건수"
    ]
    fills2 = [None, None, None, None, None,
              s["fill_slate"], s["fill_slate"], s["fill_navy"], s["fill_navy"],
              s["fill_navy"], s["fill_teal"], s["fill_teal"], s["fill_teal"], s["fill_gold"] if "fill_gold" in s else None,
              s["fill_teal"], None]
    _write_header_row(ws2, 3, headers2, fills2, s)
    ws2.freeze_panes = "A4"

    user_summaries = aggregate_user_holidays(records, user_positions)
    r2 = 4
    for idx, u in enumerate(user_summaries, 1):
        row_data = [
            idx,
            u["emp_id"],
            u["name"],
            u["team"],
            u["position"],
            int(u["sub_work_days"]),
            int(u["legal_holiday_days"]),
            int(u["normal_holiday_days"]),
            int(u.get("sub_holiday_days", 0)),
            int(u["total_days"]),
            int(u.get("pre_deduct_count", 0)),
            int(u.get("pre_deduct_remaining", 0)),
            float(u["actual_overtime_days"]),
            int(u.get("bonus_count", 0)),
            float(u.get("actual_overtime_with_bonus", u["actual_overtime_days"] + u.get("bonus_count", 0))),
            int(u["records_count"])
        ]
        center_c = {1, 2, 3, 4, 5}
        right_c  = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}
        num_fmt  = {6: "#,##0", 7: "#,##0", 8: "#,##0", 9: "#,##0", 10: "#,##0",
                    11: "#,##0", 12: "#,##0", 13: "0.0", 14: "#,##0", 15: "0.0", 16: "#,##0"}
        hl_c = {13, 15}

        _write_data_row(ws2, r2, row_data, s, r2 % 2 == 0, center_c, right_c, num_fmt, hl_c)
        r2 += 1

    # 합계 행
    _write_total_row(ws2, r2, f"전체 합계 ({len(user_summaries)}명)", 5,
                     [(6, "#,##0", False), (7, "#,##0", False), (8, "#,##0", False),
                      (9, "#,##0", False), (10, "#,##0", False),
                      (11, "#,##0", False), (12, "#,##0", False), (13, "0.0", True),
                      (14, "#,##0", False), (15, "0.0", True), (16, "#,##0", False)], s)

    # ─────────────────────────────────────────────
    # 시트 3: 특근신청_전체원장
    # ─────────────────────────────────────────────
    ws3 = wb.create_sheet(title="특근신청_전체원장")
    _write_title(ws3,
                 "특근(초과근무) 원 신청서 등록 원장 내역",
                 f"출력일시: {now_str}  |  총 {len(records)}건  |  일수 단위: 순수 숫자",
                 16, s)

    headers3 = [
        "순번", "사번", "성명", "소속팀", "특근분류", "보너스 부여",
        "시작일", "종료일", "일수", "대체휴무 사용일", "대체휴가 사용일수",
        "사전차감", "출장기간(시작)", "출장기간(종료)",
        "프로젝트 번호", "근무 장소", "특근 사유", "확인(승인)", "확인자", "신청일시"
    ]
    fills3 = [None] * len(headers3)
    _write_header_row(ws3, 3, headers3, fills3, s)
    ws3.freeze_panes = "A4"

    r3 = 4
    for idx, r in enumerate(records, 1):
        try:
            d1 = datetime.strptime(r["start_date"], "%Y-%m-%d")
            d2 = datetime.strptime(r.get("end_date") or r["start_date"], "%Y-%m-%d")
            days_count = max(1, (d2 - d1).days + 1)
        except Exception:
            days_count = 1

        is_conf = bool(r.get("is_confirmed", 0))
        sub_date = r.get("sub_holiday_date", "") or "-"
        sub_used = float(r.get("sub_holiday_used", 0.0) or 0.0)
        bonus_val = "부여(O)" if r.get("bonus_granted", 0) == 1 else "미부여(-)"
        pre_deduct_val = "O" if int(r.get("is_pre_deduct", 0) or 0) else "-"
        trip_s = r.get("trip_start_date", "") or "-"
        trip_e = r.get("trip_end_date", "") or "-"

        row_data = [
            idx,
            r.get("emp_id", ""),
            r.get("user_name", ""),
            r.get("team", ""),
            r.get("category", ""),
            bonus_val,
            r.get("start_date", ""),
            r.get("end_date", "") or r.get("start_date", ""),
            int(days_count),
            sub_date,
            sub_used,
            pre_deduct_val,
            trip_s,
            trip_e,
            r.get("project_no", "") or "-",
            r.get("location", "") or "-",
            r.get("reason", "") or "-",
            "확인완료" if is_conf else "승인대기",
            r.get("confirmed_by", "") or "-",
            r.get("created_at", "")
        ]
        center_c = {1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 13, 14, 18, 19, 20}
        right_c  = {9, 11}
        num_fmt  = {9: "#,##0", 11: "0.0"}

        _write_data_row(ws3, r3, row_data, s, r3 % 2 == 0, center_c, right_c, num_fmt)

        conf_cell = ws3.cell(row=r3, column=18)
        conf_cell.fill = s["fill_conf"] if is_conf else s["fill_pend"]
        conf_cell.font = s["conf_font"] if is_conf else s["pend_font"]

        r3 += 1

    _write_total_row(ws3, r3, f"합계 ({len(records)}건)", 8,
                     [(9, "#,##0", False), (11, "0.0", False)], s)
    ws3.cell(row=r3, column=10, value="-").alignment = s["align_center"]
    ws3.cell(row=r3, column=10).fill = s["fill_total"]
    ws3.cell(row=r3, column=10).border = s["border_total"]
    for c in range(12, 21):
        ws3.cell(row=r3, column=c, value="").fill = s["fill_total"]
        ws3.cell(row=r3, column=c).border = s["border_total"]

    # ─────────────────────────────────────────────
    # 시트 4: 부서별_요약
    # ─────────────────────────────────────────────
    ws4 = wb.create_sheet(title="부서별_요약")
    _write_title(ws4,
                 "부서(소속팀)별 특근 휴일 현황 총괄표",
                 f"취합 기간: {period_str}  |  출력 일시: {now_str}  |  수식: [총특근일수] = [대체근무]+[법정휴일]+[일반휴일]  /  [★ 팀 최종 실특근일] = 팀원 [★최종 실특근일] 합계",
                 11, s)

    headers4 = [
        "순번", "소속팀", "소속 인원수", "대체근무 (일)", "법정휴일 (일)",
        "일반휴일 (일)", "대체휴무 (일)", "총 특근일수\n(대체+법정+일반)", "★ 팀 최종 실특근일\n(일반특근-사전차감-(대휴-출장내차감)) [일]", "보너스 부여 (건)",
        "★ 팀 최종 실특근일+보너스 [일]", "총 신청건수"
    ]
    fills4 = [None, None, None, s["fill_slate"], s["fill_slate"],
              s["fill_navy"], s["fill_navy"], s["fill_navy"], s["fill_teal"], s["fill_gold"] if "fill_gold" in s else None,
              s["fill_teal"], None]
    _write_header_row(ws4, 3, headers4, fills4, s)
    ws4.freeze_panes = "A4"

    team_summaries = aggregate_team_holidays(user_summaries)
    r4 = 4
    for idx, tm in enumerate(team_summaries, 1):
        row_data = [
            idx,
            tm["team"],
            int(tm["member_count"]),
            int(tm["sub_work_days"]),
            int(tm["legal_holiday_days"]),
            int(tm["normal_holiday_days"]),
            int(tm.get("sub_holiday_days", 0)),
            int(tm["total_days"]),
            float(tm["actual_overtime_days"]),
            int(tm.get("bonus_count", 0)),
            float(tm.get("actual_overtime_with_bonus", tm["actual_overtime_days"] + tm.get("bonus_count", 0))),
            int(tm["records_count"])
        ]
        center_c = {1, 2}
        right_c  = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12}
        num_fmt  = {3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0",
                    7: "#,##0", 8: "#,##0", 9: "0.0", 10: "#,##0", 11: "0.0", 12: "#,##0"}
        hl_c = {9, 11}

        _write_data_row(ws4, r4, row_data, s, r4 % 2 == 0, center_c, right_c, num_fmt, hl_c)
        r4 += 1

    _write_total_row(ws4, r4, f"합계 ({len(team_summaries)}개 팀)", 2,
                     [(3, "#,##0", False), (4, "#,##0", False), (5, "#,##0", False),
                      (6, "#,##0", False), (7, "#,##0", False), (8, "#,##0", False),
                      (9, "0.0", True), (10, "#,##0", False), (11, "0.0", True), (12, "#,##0", False)], s)

    # 전 시트 열 너비 자동 조정
    for ws in [ws1, ws2, ws3, ws4]:
        _auto_col_width(ws)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ===========================================================================
#  메인 엑셀 생성 함수 2: 실특근 정산표 전용 (서버 API 사용)
# ===========================================================================

def generate_settlement_excel(
    summary_data: dict,
    start_date: str = "",
    end_date: str = "",
    team_filter: str = ""
) -> io.BytesIO:
    """
    실특근 정산 요약 데이터를 openpyxl로 스타일링된 2개 시트 워크북 생성
    - 시트1: [개인별_실특근_정산표]  인원별 대체근무/법정휴일/일반휴일/대체휴가/실특근 집계
    - 시트2: [부서별_정산_요약표]    팀별 현황
    """
    s = _make_styles()
    wb = Workbook()
    now_str = datetime.now().strftime("%Y년 %m월 %d일 %H:%M")

    period_str = ""
    if start_date and end_date:
        period_str = f"조회기간: {start_date} ~ {end_date}"
    elif start_date:
        period_str = f"조회기간: {start_date} 이후"
    elif end_date:
        period_str = f"조회기간: {end_date} 이전"
    else:
        period_str = "조회기간: 전체"

    if team_filter:
        period_str += f"  |  소속팀: {team_filter}"

    users = summary_data.get("user_summary", [])
    teams = summary_data.get("team_summary", [])

    # ─────────────────────────────────────────────
    # 시트 1: 개인별_실특근_정산표
    # ─────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "개인별_실특근_정산표"

    _write_title(ws1,
                 "★ 개인별 최종 실특근일 정산표",
                 f"{period_str}  |  산출식: [일반휴일] - [사전차감] - ([대체휴가 사용일수] - [출장내사전차감]) = [★최종 실특근일]  |  출력일시: {now_str}",
                 13, s)

    headers1 = [
        "순번", "사원번호", "성명", "소속팀",
        "대체근무 (일)", "법정휴일 (일)", "일반휴일 (일)", "대체휴무 (일)",
        "총 특근일수 (일)",
        "사전차감 (회)", "사전차감 잔여 (회)",
        "★ 최종 실특근일 (일)", "보너스 부여 (건)", "★ 최종 실특근일+보너스 [일]", "신청건수"
    ]
    fills1 = [None, None, None, None,
              s["fill_slate"], s["fill_slate"], s["fill_navy"], s["fill_navy"],
              s["fill_navy"], s["fill_teal"], s["fill_teal"],
              s["fill_teal"], s["fill_gold"] if "fill_gold" in s else None, s["fill_teal"], None]
    _write_header_row(ws1, 3, headers1, fills1, s)
    ws1.freeze_panes = "A4"

    r1 = 4
    for idx, u in enumerate(users, 1):
        # 총 특근일수 = 대체근무 + 법정휴일 + 일반휴일
        sub_d = int(u.get("excluded_sub_days", 0) or 0)
        leg_d = int(u.get("excluded_legal_days", 0) or 0)
        ot_d = int(u.get("overtime_days", 0) or 0)
        calculated_total = sub_d + leg_d + ot_d
        actual_ot = float(u.get("actual_overtime_days", 0.0) or 0.0)
        b_cnt = int(u.get("bonus_count", 0) or 0)
        actual_with_b = float(u.get("actual_overtime_with_bonus", actual_ot + b_cnt))

        row_data = [
            idx,
            u.get("emp_id", ""),
            u.get("name", ""),
            u.get("team", ""),
            sub_d,
            leg_d,
            ot_d,
            int(u.get("sub_holiday_days", 0) or 0),
            calculated_total,
            int(u.get("pre_deduct_count", 0) or 0),
            int(u.get("pre_deduct_remaining", 0) or 0),
            actual_ot,
            b_cnt,
            actual_with_b,
            int(u.get("records_count", 1) or 1)
        ]
        center_c = {1, 2, 3, 4}
        right_c  = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}
        num_fmt  = {5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "#,##0", 9: "#,##0",
                    10: "#,##0", 11: "#,##0", 12: "0.0", 13: "#,##0", 14: "0.0", 15: "#,##0"}
        hl_c = {12, 14}

        _write_data_row(ws1, r1, row_data, s, r1 % 2 == 0, center_c, right_c, num_fmt, hl_c)
        r1 += 1

    # 합계 행
    _write_total_row(ws1, r1, f"전체 합계 ({len(users)}명)", 4,
                     [(5, "#,##0", False), (6, "#,##0", False), (7, "#,##0", False),
                      (8, "#,##0", False), (9, "#,##0", False),
                      (10, "#,##0", False), (11, "#,##0", False), (12, "0.0", True),
                      (13, "#,##0", False), (14, "0.0", True), (15, "#,##0", False)], s)

    # ─────────────────────────────────────────────
    # 시트 2: 부서별_정산_요약표
    # ─────────────────────────────────────────────
    ws2 = wb.create_sheet(title="부서별_정산_요약표")
    _write_title(ws2,
                 "부서(소속팀)별 특근 정산 요약표",
                 f"{period_str}  |  팀별 인원수 및 분류별 일수 합산  |  출력일시: {now_str}",
                 8, s)

    headers2 = [
        "순번", "소속팀", "소속 인원수", "총 신청일수",
        "제외일수 (대체+법정)", "인정 특근일 (일반휴일)",
        "대체휴가 사용일수", "★ 팀 최종 실특근일"
    ]
    fills2 = [None, None, None, s["fill_slate"],
              s["fill_slate"], s["fill_navy"], s["fill_navy"], s["fill_teal"]]
    _write_header_row(ws2, 3, headers2, fills2, s)
    ws2.freeze_panes = "A4"

    r2 = 4
    for idx, t in enumerate(teams, 1):
        row_data = [
            idx,
            t.get("team", ""),
            int(t.get("member_count", 0) or 0),
            int(t.get("total_days", 0) or 0),
            int(t.get("excluded_days", 0) or 0),
            int(t.get("overtime_days", 0) or 0),
            float(t.get("sub_holiday_used", 0.0) or 0.0),
            float(t.get("actual_overtime_days", 0.0) or 0.0)
        ]
        center_c = {1, 2}
        right_c  = {3, 4, 5, 6, 7, 8}
        num_fmt  = {3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "0.0", 8: "0.0"}
        hl_c = {8}

        _write_data_row(ws2, r2, row_data, s, r2 % 2 == 0, center_c, right_c, num_fmt, hl_c)
        r2 += 1

    _write_total_row(ws2, r2, f"합계 ({len(teams)}개 팀)", 2,
                     [(3, "#,##0", False), (4, "#,##0", False), (5, "#,##0", False),
                      (6, "#,##0", False), (7, "0.0", False), (8, "0.0", True)], s)

    # 전 시트 열 너비 자동 조정
    for ws in [ws1, ws2]:
        _auto_col_width(ws)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
