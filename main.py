import os
import shutil
import json
import socket
import io
import base64
import sqlite3
import re
from datetime import datetime
from typing import Optional, List
from pathlib import Path

from fastapi import FastAPI, HTTPException, Query, Response, Request, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import qrcode

from database import (
    init_db, get_db_connection, log_audit, create_backup, DB_PATH,
    get_all_teams, create_team, delete_team, log_access_event,
    is_using_turso, get_db_mode, get_cached_user_role,
    invalidate_user_role_cache, invalidate_teams_cache
)
from schemas import (
    UserLoginRequest, UserRegisterRequest, UserUpdateRequest,
    OvertimeCreateRequest, OvertimeUpdateRequest, OvertimeDeleteRequest,
    OvertimeConfirmRequest, OvertimePreDeductRequest, ExportRequest, TeamCreateRequest,
    BackupSaveRequest, BackupLoadRequest,
    OvertimeFinalizeRequest, OvertimeBatchFinalizeRequest,
    OvertimeReviewRequest, OvertimeBatchReviewRequest,
    SuggestionCreateRequest, SuggestionReplyRequest
)
from urllib.parse import quote
from exporter import (
    generate_overtime_excel,
    generate_settlement_excel,
    aggregate_user_holidays,
    aggregate_team_holidays,
    import_overtimes_from_excel
)
from generate_manual import create_manual, PPTX_PATH, USER_PPTX_PATH, ADMIN_PPTX_PATH

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
WEB_BACKUP_DIR = BASE_DIR / "data" / "web_backups"
WEB_BACKUP_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_EXTERNAL_URL = "https://internal-overtime.company.com"
EXTERNAL_URL_FILE = BASE_DIR / "data" / "external_url.txt"

if not EXTERNAL_URL_FILE.exists():
    try:
        EXTERNAL_URL_FILE.parent.mkdir(parents=True, exist_ok=True)
        EXTERNAL_URL_FILE.write_text(DEFAULT_EXTERNAL_URL, encoding="utf-8")
    except Exception:
        pass

app = FastAPI(title="Team Overtime Manager", version="v1.52")

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    import traceback
    err_trace = traceback.format_exc()
    print(f"[Unhandled Exception on {request.url.path}] {err_trace}")
    return JSONResponse(
        status_code=500,
        content={
            "detail": f"서버 처리 중 오류가 발생했습니다: {str(exc)}",
            "error_type": type(exc).__name__,
            "path": request.url.path
        }
    )

def validate_emp_id(emp_id: str):
    r"""사원번호 유효성 검증:
    - 슈퍼관리자(ps37082) 제외
    - 무조건 숫자 6자리여야 하고 첫 자리가 1 또는 2로 시작 (^[12]\d{5}$)
    - 안내 문구 미노출 보안 원칙: 실패 시 '입력이 올바르지 않습니다.'만 표시
    """
    emp_id = (emp_id or "").strip()
    if not emp_id:
        raise HTTPException(status_code=400, detail="입력이 올바르지 않습니다.")
    if emp_id.lower() == "ps37082":
        return True
    if not re.match(r"^[12]\d{5}$", emp_id):
        raise HTTPException(status_code=400, detail="입력이 올바르지 않습니다.")
    return True

def get_client_ip(request: Request) -> str:
    """클라이언트 실제 IP 추출 (프록시 X-Forwarded-For 대응)"""
    forwarded = request.headers.get("X-Forwarded-For")
    if forwarded:
        return forwarded.split(",")[0].strip()
    if request.client and request.client.host:
        return request.client.host
    return "127.0.0.1"

# CORS 활성화 (모바일 및 외부 환경 접속 지원)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 캐시 방지 미들웨어: JS/HTML 파일은 항상 최신 버전을 받도록 강제
from starlette.middleware.base import BaseHTTPMiddleware

class NoCacheMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        response = await call_next(request)
        path = request.url.path
        if path.endswith(('.js', '.html', '.css')) or path == '/':
            response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'
        return response

app.add_middleware(NoCacheMiddleware)

@app.on_event("startup")
def startup_event():
    init_db()
    sync_static_files()
    # PPT 매뉴얼 생성 (비동기 백그라운드 스레드로 실행하여 서버 구동 지연 0초)
    import threading
    def _bg_create_manual():
        try:
            if not PPTX_PATH.exists():
                create_manual()
        except Exception as _e:
            print(f"[Manual Gen Background Notice] {_e}")
    threading.Thread(target=_bg_create_manual, daemon=True).start()

def get_local_ips():
    """서버의 로컬 네트워크 IP 목록 검색"""
    ips = []
    try:
        hostname = socket.gethostname()
        for ip in socket.gethostbyname_ex(hostname)[2]:
            if not ip.startswith("127.") and ":" not in ip:
                ips.append(ip)
    except Exception:
        pass
    if not ips:
        ips.append("127.0.0.1")
    return ips

# ----------------- 시스템 & 접속 API -----------------

def get_current_external_url(request: Optional[Request] = None) -> str:
    """외부 접속망 URL 획득 (우선순위: 환경변수 > 저장된 파일 > 요청 Host/헤더 > 기본 Cloudflare 도메인)"""
    env_url = os.environ.get("EXTERNAL_URL", "").strip()
    if env_url:
        return env_url

    if EXTERNAL_URL_FILE.exists():
        try:
            val = EXTERNAL_URL_FILE.read_text(encoding="utf-8").strip()
            if val:
                return val
        except Exception:
            pass

    if request:
        proto = request.headers.get("x-forwarded-proto", "https")
        host = request.headers.get("x-forwarded-host") or request.headers.get("host", "")
        if host and not any(host.startswith(h) for h in ["localhost", "127.0.0.1", "192.168.", "10.", "172."]):
            return f"{proto}://{host}"

    return DEFAULT_EXTERNAL_URL

@app.get("/api/system/info")
def get_system_info(request: Request, custom_url: Optional[str] = None):
    """서버 접속 정보 및 모바일 스캔용 QR 코드 제공 (외부망 접속 URL 기본 제공 및 영구 저장)"""
    ips = get_local_ips()
    primary_ip = ips[0]
    local_url = f"http://{primary_ip}:8000"

    # 외부 접속 URL 설정 및 영구 저장
    if custom_url and custom_url.strip():
        target_url = custom_url.strip()
        try:
            EXTERNAL_URL_FILE.write_text(target_url, encoding="utf-8")
        except Exception as e:
            print(f"[Warning] 외부 URL 저장 실패: {e}")
    else:
        target_url = get_current_external_url(request)

    # QR코드 생성 (외부망 접속 URL 기준)
    qr = qrcode.QRCode(version=1, box_size=6, border=2)
    qr.add_data(target_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#0f172a", back_color="#ffffff")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    qr_bytes = buf.getvalue()
    qr_b64 = base64.b64encode(qr_bytes).decode("utf-8")

    # external_qr.png 루트 및 static에 동기화 저장
    try:
        (BASE_DIR / "external_qr.png").write_bytes(qr_bytes)
        if STATIC_DIR.exists():
            (STATIC_DIR / "external_qr.png").write_bytes(qr_bytes)
    except Exception:
        pass

    from database import get_db_mode, is_using_turso
    return {
        "status": "success",
        "version": app.version,
        "db_mode": get_db_mode(),
        "is_turso": is_using_turso(),
        "storage": "Turso Cloud DB (영구 보존)" if is_using_turso() else "Local SQLite (로컬 저장소)",
        "local_ips": ips,
        "local_url": local_url,
        "external_url": target_url,
        "primary_url": target_url,
        "is_custom": bool(custom_url and custom_url.strip()),
        "qr_code_base64": f"data:image/png;base64,{qr_b64}",
        "server_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

# ----------------- 사용자 및 인증 API -----------------

@app.post("/api/users/login")
def login_user(req: UserLoginRequest, request: Request):
    """사번으로 회원 조회 및 로그인 (보안 접속 감사 로그 기록)"""
    emp_id = (req.emp_id or "").strip()
    client_ip = get_client_ip(request)
    ua = request.headers.get("User-Agent", "")[:250]

    if not emp_id:
        log_access_event(emp_id="", action_type="LOGIN", status="FAILURE", ip_address=client_ip, user_agent=ua, details="사원번호 미입력")
        raise HTTPException(status_code=400, detail="입력이 올바르지 않습니다.")

    # 사원번호 유효성 검사 (규칙: 1 또는 2로 시작하는 6자리 숫자, 슈퍼관리자 제외)
    validate_emp_id(emp_id)

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE emp_id = ?", (emp_id,))
    row = cursor.fetchone()
    conn.close()

    if row:
        user_dict = dict(row)
        log_access_event(
            emp_id=emp_id,
            user_name=user_dict.get("name"),
            action_type="LOGIN",
            status="SUCCESS",
            ip_address=client_ip,
            user_agent=ua,
            details=f"로그인 성공 ({user_dict.get('team')}, {user_dict.get('position')})"
        )
        return {"exists": True, "user": user_dict}
    else:
        log_access_event(
            emp_id=emp_id,
            user_name=None,
            action_type="LOGIN",
            status="FAILURE",
            ip_address=client_ip,
            user_agent=ua,
            details="미등록 사원번호 로그인 시도"
        )
        return {"exists": False, "emp_id": emp_id}

@app.post("/api/users/register")
def register_user(req: UserRegisterRequest, request: Request):
    """신규 팀원(사번) 등록 (감사 로그 기록)"""
    emp_id = (req.emp_id or "").strip()
    name = (req.name or "").strip()
    team = (req.team or "").strip()
    position = (req.position or "팀원").strip()
    is_admin = 1 if req.is_admin else 0
    client_ip = get_client_ip(request)
    ua = request.headers.get("User-Agent", "")[:250]

    if not emp_id or not name or not team:
        log_access_event(emp_id=emp_id, user_name=name, action_type="REGISTER", status="FAILURE", ip_address=client_ip, user_agent=ua, details="필수값(사번/이름/팀) 누락 등록 시도")
        raise HTTPException(status_code=400, detail="사원번호, 이름, 소속팀은 필수 입력값입니다.")

    # 사원번호 유효성 검사 (6자리 힌트 제거)
    validate_emp_id(emp_id)

    # 슈퍼관리자 사번인 경우 자동 슈퍼관리자 권한
    is_super = 1 if emp_id.lower() == "ps37082" else 0
    if is_super:
        is_admin = 1

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 중복 체크
    cursor.execute("SELECT * FROM users WHERE emp_id = ?", (emp_id,))
    if cursor.fetchone():
        conn.close()
        log_access_event(emp_id=emp_id, user_name=name, action_type="REGISTER", status="FAILURE", ip_address=client_ip, user_agent=ua, details="이미 등록된 사원번호 중복 등록 시도")
        raise HTTPException(status_code=400, detail="이미 등록된 사원번호입니다.")

    cursor.execute("""
    INSERT INTO users (emp_id, name, team, position, is_admin, is_super, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (emp_id, name, team, position, is_admin, is_super, now_str))
    conn.commit()

    cursor.execute("SELECT * FROM users WHERE emp_id = ?", (emp_id,))
    new_user = dict(cursor.fetchone())
    conn.close()

    log_access_event(
        emp_id=emp_id,
        user_name=name,
        action_type="REGISTER",
        status="SUCCESS",
        ip_address=client_ip,
        user_agent=ua,
        details=f"신규 사원 등록 완료 ({team}, {position})"
    )
    return {"message": "등록되었습니다.", "user": new_user}

@app.post("/api/users/logout")
def logout_user(req: UserLoginRequest, request: Request):
    """사용자 로그아웃 감사 로그 기록"""
    emp_id = (req.emp_id or "").strip()
    client_ip = get_client_ip(request)
    ua = request.headers.get("User-Agent", "")[:250]
    log_access_event(
        emp_id=emp_id,
        action_type="LOGOUT",
        status="SUCCESS",
        ip_address=client_ip,
        user_agent=ua,
        details="사용자 로그아웃"
    )
    return {"status": "ok"}

@app.get("/api/users/{emp_id}/preferences")
def get_user_preferences(emp_id: str):
    """사용자별 마지막 입력 선호 정보 조회 (프로젝트번호, 근무장소)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM user_preferences WHERE emp_id = ?", (emp_id.strip(),))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {"exists": True, "preferences": dict(row)}
    # fallback: 최근 특근 신청 내역에서 가져오기
    conn2 = get_db_connection()
    cursor2 = conn2.cursor()
    cursor2.execute("""
    SELECT project_no, location FROM overtimes
    WHERE emp_id = ? AND (project_no != '' OR location != '')
    ORDER BY created_at DESC, id DESC
    LIMIT 1
    """, (emp_id.strip(),))
    row2 = cursor2.fetchone()
    conn2.close()
    if row2:
        return {"exists": True, "preferences": {"emp_id": emp_id, "last_project_no": row2["project_no"] or "", "last_location": row2["location"] or ""}}
    return {"exists": False, "preferences": {"emp_id": emp_id, "last_project_no": "", "last_location": ""}}

@app.get("/api/users/{emp_id}/latest-overtime")
def get_latest_overtime(emp_id: str):
    """해당 사용자의 가장 최근 특근 신청 내역(프로젝트번호, 장소, 사유, 분류) 가이드 제공"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT project_no, location, reason, category
    FROM overtimes
    WHERE emp_id = ?
    ORDER BY created_at DESC, id DESC
    LIMIT 1
    """, (emp_id.strip(),))
    row = cursor.fetchone()
    conn.close()
    if row:
        ot_dict = dict(row)
        return {
            "exists": True,
            "has_previous": True,
            "overtime": ot_dict,
            "latest": ot_dict
        }
    return {
        "exists": False,
        "has_previous": False,
        "overtime": None,
        "latest": None
    }

@app.get("/api/users")
def list_users(search: Optional[str] = None, admin_emp_id: Optional[str] = None):
    """회원 목록 조회 (슈퍼관리자는 전체/팀별 구분, 팀관리자는 본인 소속팀만 조회)"""
    conn = get_db_connection()
    cursor = conn.cursor()

    req_team = None
    if admin_emp_id:
        caller = get_cached_user_role(admin_emp_id)
        if caller and caller.get("is_super") != 1 and caller.get("is_admin") == 1:
            req_team = caller.get("team")

    query = "SELECT * FROM users WHERE 1=1"
    params = []
    if req_team:
        query += " AND team = ?"
        params.append(req_team)
    if search:
        s = f"%{search.strip()}%"
        query += " AND (emp_id LIKE ? OR name LIKE ? OR team LIKE ?)"
        params.extend([s, s, s])

    query += " ORDER BY is_super DESC, is_admin DESC, name ASC"
    cursor.execute(query, tuple(params))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return {"users": rows}

@app.put("/api/users/{emp_id}")
def update_user(emp_id: str, req: UserUpdateRequest):
    """회원 정보 수정 (팀관리자는 본인 팀원 성명/직급만 수정 가능, 슈퍼관리자는 전원 및 소속팀/권한 수정 가능)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE emp_id = ?", (emp_id,))
    user = cursor.fetchone()
    if not user:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 회원을 찾을 수 없습니다.")

    user_dict = dict(user)
    
    # 수정 요청자(관리자) 권한 확인
    caller_is_super = 0
    caller_team = None
    if req.admin_emp_id:
        caller = get_cached_user_role(req.admin_emp_id)
        if caller:
            caller_is_super = caller.get("is_super", 0)
            caller_team = caller.get("team")

    # 팀관리자(is_super=0)의 제약 조건 검증
    if not caller_is_super:
        # 본인 소속팀의 팀원이 아니면 수정 불가
        if caller_team and user_dict["team"] != caller_team:
            conn.close()
            raise HTTPException(status_code=403, detail=f"팀관리자는 본인 소속팀({caller_team})의 팀원 정보만 수정할 수 있습니다.")
        # 소속팀 변경 시도 차단
        if req.team is not None and req.team != user_dict["team"]:
            conn.close()
            raise HTTPException(status_code=403, detail="팀원의 소속팀 변경은 슈퍼관리자만 가능합니다.")
        # 관리자 권한 변경 시도 차단
        if req.is_admin is not None and req.is_admin != user_dict["is_admin"]:
            conn.close()
            raise HTTPException(status_code=403, detail="팀관리자 지정 권한은 슈퍼관리자만 가능합니다.")

    # 슈퍼관리자 승격 / 하야 요청 검증
    is_super_val = user_dict.get("is_super", 0)
    if req.is_super is not None:
        if not caller_is_super:
            conn.close()
            raise HTTPException(status_code=403, detail="슈퍼관리자 승격 및 하야 권한은 슈퍼관리자만 가능합니다.")
        
        # 원조 총괄관리자 ps37082 보호
        if emp_id.lower() == "ps37082" and req.is_super == 0:
            conn.close()
            raise HTTPException(status_code=400, detail="원조 총괄 슈퍼관리자(ps37082) 계정은 하야할 수 없습니다.")
        
        # 본인 계정 직접 하야 방지 (권한 상실 사고 방지)
        if req.admin_emp_id and emp_id.strip() == req.admin_emp_id.strip() and req.is_super == 0:
            conn.close()
            raise HTTPException(status_code=400, detail="본인 계정을 직접 슈퍼관리자에서 하야할 수 없습니다.")
        
        is_super_val = int(req.is_super)

    name = req.name if req.name is not None else user_dict["name"]
    team = req.team if req.team is not None else user_dict["team"]
    position = req.position if req.position is not None else user_dict["position"]
    is_admin = req.is_admin if req.is_admin is not None else user_dict["is_admin"]

    # 슈퍼관리자는 관리자 권한 필수 부여
    if is_super_val == 1 or user_dict.get("is_super", 0) == 1:
        is_admin = 1

    cursor.execute("""
    UPDATE users SET name = ?, team = ?, position = ?, is_admin = ?, is_super = ?
    WHERE emp_id = ?
    """, (name, team, position, is_admin, is_super_val, emp_id))

    # 성명 또는 소속팀 변경 시 기존 특근 내역(overtimes) 일괄 동기화 (트랜잭션)
    if name != user_dict["name"] or team != user_dict["team"]:
        cursor.execute("""
        UPDATE overtimes SET user_name = ?, team = ?
        WHERE emp_id = ?
        """, (name, team, emp_id))

    conn.commit()

    cursor.execute("SELECT * FROM users WHERE emp_id = ?", (emp_id,))
    updated = dict(cursor.fetchone())
    conn.close()
    invalidate_user_role_cache(emp_id)
    return {"message": "회원 정보가 성공적으로 수정되었습니다.", "user": updated}

@app.get("/api/users/{emp_id}/overtime-stats")
def get_user_overtime_stats(emp_id: str):
    """관리자용: 특정 사원의 기간별 특근 약식 통계 (최근 3개월, 1~4분기, 상/하반기, 연도별)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name, team FROM users WHERE emp_id = ?", (emp_id.strip(),))
    user_row = cursor.fetchone()
    if not user_row:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 회원을 찾을 수 없습니다.")
    user_name = user_row["name"]
    user_team = user_row["team"]

    cursor.execute("""
    SELECT id, category, start_date, end_date, sub_holiday_used, is_confirmed, is_pre_deduct, trip_start_date, trip_end_date
    FROM overtimes WHERE emp_id = ? ORDER BY start_date ASC
    """, (emp_id.strip(),))
    rows = cursor.fetchall()
    conn.close()

    now = datetime.now()
    current_year = now.year
    current_month = now.month

    # 1. 대체휴무 등록 건들에서 출장기간 수집
    trip_ranges = []
    for r in rows:
        cat_raw = (r["category"] or "").strip()
        t_s = (r["trip_start_date"] or "").strip()
        t_e = (r["trip_end_date"] or "").strip()
        if cat_raw in ["대체휴무", "대체휴일"] and t_s and t_e:
            trip_ranges.append((t_s, t_e))

    def make_bucket():
        return {
            "대체근무": 0, "일반휴일": 0, "법정휴일": 0, "대체휴무": 0,
            "사전차감": 0, "출장내사전차감": 0, "사전차감잔여": 0,
            "최종실특근": 0.0, "total_records": 0
        }

    # 최근 3개월 (오늘이 9월이면 7월, 8월, 9월)
    recent_3_months = []
    for i in range(2, -1, -1):
        tm = (current_month - i - 1) % 12 + 1
        ty = current_year if current_month - i > 0 else current_year - 1
        recent_3_months.append((ty, tm, f"{tm}월"))

    month_buckets = {label: make_bucket() for _, _, label in recent_3_months}
    quarter_buckets = {
        "1분기": make_bucket(),
        "2분기": make_bucket(),
        "3분기": make_bucket(),
        "4분기": make_bucket(),
    }
    half_buckets = {
        "상반기": make_bucket(),
        "하반기": make_bucket(),
    }
    by_year = {}

    for row in rows:
        try:
            d1 = datetime.strptime(row["start_date"], "%Y-%m-%d")
            d2 = datetime.strptime(row["end_date"] or row["start_date"], "%Y-%m-%d")
            days = max(1, (d2 - d1).days + 1)
        except Exception:
            d1 = d2 = datetime.now()
            days = 1

        cat = row["category"] or "일반휴일"
        if cat == "대체휴일":
            cat = "대체휴무"

        is_pre = int(row["is_pre_deduct"] or 0)
        s_date = row["start_date"]
        e_date = row["end_date"] or row["start_date"]

        # 출장기간 내 포함된 사전차감인지 확인
        in_trip = False
        if is_pre:
            for ts, te in trip_ranges:
                if not (e_date < ts or s_date > te):
                    in_trip = True
                    break

        yr = d1.year
        mon = d1.month
        quarter = f"{(mon - 1) // 3 + 1}분기"
        half = "상반기" if mon <= 6 else "하반기"

        def accumulate(b):
            b[cat if cat in b else "일반휴일"] += days
            if is_pre:
                b["사전차감"] += 1
                if in_trip:
                    b["출장내사전차감"] += 1
            b["total_records"] += 1

        # 최근 3개월 누적
        for ty, tm, label in recent_3_months:
            if yr == ty and mon == tm:
                accumulate(month_buckets[label])

        # 분기별 누적 (올해 기준)
        if yr == current_year and quarter in quarter_buckets:
            accumulate(quarter_buckets[quarter])

        # 반기별 누적 (올해 기준)
        if yr == current_year and half in half_buckets:
            accumulate(half_buckets[half])

        # 연도별 누적
        if yr not in by_year:
            by_year[yr] = make_bucket()
        accumulate(by_year[yr])

    # 잔여수 및 최종 실특근일 계산:
    # 최종 실특근 = 일반특근 - 사전차감 - (대체휴무 - 대체휴무시 작성한 출장기간 이내의 사전차감)
    # 사전차감 잔여수 = 총사전차감 - 출장기간내사전차감
    def finalize_bucket(b):
        b["사전차감잔여"] = max(0, b["사전차감"] - b["출장내사전차감"])
        b["최종실특근"] = max(0.0, round(float(b["일반휴일"] - b["사전차감"] - (b["대체휴무"] - b["출장내사전차감"])), 1))

    for b in month_buckets.values(): finalize_bucket(b)
    for b in quarter_buckets.values(): finalize_bucket(b)
    for b in half_buckets.values(): finalize_bucket(b)
    for b in by_year.values(): finalize_bucket(b)

    return {
        "emp_id": emp_id,
        "name": user_name,
        "team": user_team,
        "recent_months": month_buckets,
        "quarters": quarter_buckets,
        "halves": half_buckets,
        "by_year": {str(k): v for k, v in sorted(by_year.items(), reverse=True)},
    }

@app.get("/api/overtimes/my-stats")
def get_my_overtime_stats(emp_id: str, year: Optional[int] = None):
    """사용자 모드: 연간, 반기별(상/하반기), 월간 특근 통계 제공"""
    now = datetime.now()
    target_year = year or now.year
    target_month = now.month

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT id, category, start_date, end_date, sub_holiday_used, is_confirmed, is_pre_deduct, trip_start_date, trip_end_date
    FROM overtimes
    WHERE emp_id = ? AND start_date LIKE ?
    ORDER BY start_date ASC
    """, (emp_id.strip(), f"{target_year}-%"))
    rows = cursor.fetchall()
    conn.close()

    # 0. 대체휴무 출장기간 수집
    trip_ranges = []
    for r in rows:
        cat_raw = (r["category"] or "").strip()
        t_s = (r["trip_start_date"] or "").strip()
        t_e = (r["trip_end_date"] or "").strip()
        if cat_raw in ["대체휴무", "대체휴일"] and t_s and t_e:
            trip_ranges.append((t_s, t_e))

    # 통계 초기화
    yearly = {
        "year": target_year,
        "total_count": 0,
        "confirmed_count": 0,
        "pending_count": 0,
        "total_days": 0.0,
        "normal_days": 0.0,
        "pre_deduct_count": 0,
        "trip_pre_deduct_count": 0,
        "sub_rest_days": 0.0,
        "actual_overtime_days": 0.0,
        "sub_holiday_used": 0.0,
        "approval_rate": 0
    }
    
    first_half = {"label": "상반기 (1~6월)", "total_count": 0, "confirmed_count": 0, "normal_days": 0.0, "pre_deduct_count": 0, "trip_pre_deduct_count": 0, "sub_rest_days": 0.0, "actual_days": 0.0}
    second_half = {"label": "하반기 (7~12월)", "total_count": 0, "confirmed_count": 0, "normal_days": 0.0, "pre_deduct_count": 0, "trip_pre_deduct_count": 0, "sub_rest_days": 0.0, "actual_days": 0.0}
    
    monthly = {
        "year": target_year,
        "month": target_month,
        "total_count": 0,
        "confirmed_count": 0,
        "pending_count": 0,
        "normal_days": 0.0,
        "pre_deduct_count": 0,
        "trip_pre_deduct_count": 0,
        "sub_rest_days": 0.0,
        "actual_days": 0.0,
        "sub_holiday_used": 0.0
    }

    for row in rows:
        cat = row["category"]
        start_d = row["start_date"]
        end_d = row["end_date"]
        sub_used = float(row["sub_holiday_used"] or 0.0)
        is_conf = int(row["is_confirmed"] or 0)
        is_pre = int(row["is_pre_deduct"] or 0)

        # 일수 계산
        try:
            d1 = datetime.strptime(start_d, "%Y-%m-%d")
            d2 = datetime.strptime(end_d, "%Y-%m-%d")
            days = (d2 - d1).days + 1
            if days < 1:
                days = 1
        except Exception:
            days = 1

        # 출장기간 내 포함된 사전차감인지 확인
        in_trip = False
        if is_pre == 1:
            for ts, te in trip_ranges:
                if not (end_d < ts or start_d > te):
                    in_trip = True
                    break

        ot_d = float(days) if cat == "일반휴일" else 0.0
        sub_d = float(days) if cat in ["대체휴무", "대체휴일"] else sub_used
        pre_c = 1 if is_pre == 1 else 0
        trip_pre_c = 1 if in_trip else 0

        # 월 파악
        try:
            m = int(start_d.split("-")[1])
        except Exception:
            m = 1

        # 1. 연간 누적
        yearly["total_count"] += 1
        yearly["total_days"] += days
        yearly["normal_days"] += ot_d
        yearly["pre_deduct_count"] += pre_c
        yearly["trip_pre_deduct_count"] += trip_pre_c
        yearly["sub_rest_days"] += sub_d
        yearly["sub_holiday_used"] += sub_used
        if is_conf == 1:
            yearly["confirmed_count"] += 1
        else:
            yearly["pending_count"] += 1

        # 2. 반기 누적
        target_half = first_half if (1 <= m <= 6) else second_half
        target_half["total_count"] += 1
        target_half["normal_days"] += ot_d
        target_half["pre_deduct_count"] += pre_c
        target_half["trip_pre_deduct_count"] += trip_pre_c
        target_half["sub_rest_days"] += sub_d
        if is_conf == 1:
            target_half["confirmed_count"] += 1

        # 3. 월간 누적
        if m == target_month:
            monthly["total_count"] += 1
            monthly["normal_days"] += ot_d
            monthly["pre_deduct_count"] += pre_c
            monthly["trip_pre_deduct_count"] += trip_pre_c
            monthly["sub_rest_days"] += sub_d
            monthly["sub_holiday_used"] += sub_used
            if is_conf == 1:
                monthly["confirmed_count"] += 1
            else:
                monthly["pending_count"] += 1

    # 최종 실특근 = 일반특근 - 사전차감 - (대체휴무 - 대체휴무시 작성한 출장기간 이내의 사전차감)
    def calc_act(norm, pre, trip_pre, sub):
        return max(0.0, round(float(norm - pre - (sub - trip_pre)), 1))

    yearly["total_days"] = round(yearly["total_days"], 1)
    yearly["actual_overtime_days"] = calc_act(yearly["normal_days"], yearly["pre_deduct_count"], yearly["trip_pre_deduct_count"], yearly["sub_rest_days"])
    yearly["sub_holiday_used"] = round(yearly["sub_holiday_used"], 1)
    if yearly["total_count"] > 0:
        yearly["approval_rate"] = round((yearly["confirmed_count"] / yearly["total_count"]) * 100, 1)

    first_half["actual_days"] = calc_act(first_half["normal_days"], first_half["pre_deduct_count"], first_half["trip_pre_deduct_count"], first_half["sub_rest_days"])
    second_half["actual_days"] = calc_act(second_half["normal_days"], second_half["pre_deduct_count"], second_half["trip_pre_deduct_count"], second_half["sub_rest_days"])
    monthly["actual_days"] = calc_act(monthly["normal_days"], monthly["pre_deduct_count"], monthly["trip_pre_deduct_count"], monthly["sub_rest_days"])
    monthly["sub_holiday_used"] = round(monthly["sub_holiday_used"], 1)

    return {
        "yearly": yearly,
        "half_yearly": {
            "first_half": first_half,
            "second_half": second_half,
            "current_is_first": (1 <= target_month <= 6)
        },
        "monthly": monthly
    }

@app.delete("/api/users/{emp_id}")
def delete_user(emp_id: str, admin_emp_id: Optional[str] = None):
    """회원 삭제 (슈퍼관리자는 전원 삭제 가능, 팀관리자는 본인 소속팀 팀원만 삭제 가능)"""
    if emp_id.lower() == "ps37082":
        raise HTTPException(status_code=400, detail="총괄 슈퍼관리자 계정은 삭제할 수 없습니다.")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE emp_id = ?", (emp_id,))
    user_row = cursor.fetchone()
    if not user_row:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 회원을 찾을 수 없습니다.")

    # 호출자 권한 확인 (요구사항 2: 팀관리자는 본인 팀만 관리 가능)
    if admin_emp_id:
        caller = get_cached_user_role(admin_emp_id)
        if not caller or (caller.get("is_super") != 1 and caller.get("is_admin") != 1):
            conn.close()
            raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다.")
        
        # 팀관리자(is_super=0)인 경우 본인 팀 검증
        if caller["is_super"] != 1:
            if user_row["team"] != caller["team"]:
                conn.close()
                raise HTTPException(status_code=403, detail=f"팀관리자는 본인 소속팀({caller['team']})의 팀원만 삭제할 수 있습니다.")
            if user_row["is_admin"] == 1 or user_row["is_super"] == 1:
                conn.close()
                raise HTTPException(status_code=403, detail="팀관리자는 다른 관리자 계정을 삭제할 수 없습니다.")

    user_name = user_row["name"]

    # 1. 외래키 제약 에러 방지: 해당 팀원의 특근 이력(overtime_history) 선행 삭제
    cursor.execute("""
        DELETE FROM overtime_history 
        WHERE overtime_id IN (SELECT id FROM overtimes WHERE emp_id = ?)
    """, (emp_id,))

    # 2. 해당 팀원의 특근 신청 데이터(overtimes) 삭제
    cursor.execute("DELETE FROM overtimes WHERE emp_id = ?", (emp_id,))

    # 3. 회원(users) 레코드 삭제
    cursor.execute("DELETE FROM users WHERE emp_id = ?", (emp_id,))
    conn.commit()
    conn.close()
    return {"message": f"'{user_name}({emp_id})' 회원이 성공적으로 삭제되었습니다."}

# ----------------- 소속팀(부서) 관리 API (요구사항 25) -----------------

@app.get("/api/teams")
def list_teams():
    """전체 소속팀 목록 조회 (리스트박스용)"""
    return {"teams": get_all_teams()}

@app.post("/api/teams")
def add_team(req: TeamCreateRequest):
    """신규 소속팀 생성 (관리자/슈퍼관리자)"""
    admin_row = get_cached_user_role(req.admin_emp_id)
    if not admin_row or (admin_row.get("is_super") != 1 and admin_row.get("is_admin") != 1):
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다.")

    conn = get_db_connection()
    cursor = conn.cursor()

    team_name = req.name.strip()
    if not team_name:
        conn.close()
        raise HTTPException(status_code=400, detail="소속팀 이름을 입력해주세요.")

    cursor.execute("SELECT COUNT(*) as cnt FROM teams WHERE name = ?", (team_name,))
    if cursor.fetchone()["cnt"] > 0:
        conn.close()
        raise HTTPException(status_code=400, detail=f"'{team_name}' 소속팀은 이미 존재합니다.")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("INSERT INTO teams (name, created_at) VALUES (?, ?)", (team_name, now_str))
    team_id = cursor.lastrowid or 0
    conn.commit()
    conn.close()
    return {"message": f"'{team_name}' 소속팀이 추가되었습니다.", "team": {"id": team_id, "name": team_name, "created_at": now_str}}

@app.delete("/api/teams/{team_name}")
def remove_team(team_name: str, admin_emp_id: str):
    """소속팀 삭제 (관리자/슈퍼관리자 전용, 소속 인원이 없는 경우만 가능)"""
    admin_row = get_cached_user_role(admin_emp_id)
    if not admin_row or (admin_row.get("is_super") != 1 and admin_row.get("is_admin") != 1):
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다.")

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) as cnt FROM users WHERE team = ?", (team_name,))
    user_cnt = cursor.fetchone()["cnt"]
    if user_cnt > 0:
        conn.close()
        raise HTTPException(status_code=400, detail=f"'{team_name}'에 등록된 팀원이 {user_cnt}명 있어 삭제할 수 없습니다. 팀원의 소속을 먼저 변경하거나 팀원을 삭제해주세요.")

    try:
        cursor.execute("DELETE FROM teams WHERE name = ?", (team_name,))
        conn.commit()
        conn.close()
        return {"message": f"'{team_name}' 소속팀이 삭제되었습니다."}
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=400, detail=f"소속팀 삭제 실패: {str(e)}")


# ----------------- 특근 신청 및 조회 API -----------------

@app.get("/api/overtimes")
def list_overtimes(
    emp_id: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    team: Optional[str] = None,
    category: Optional[str] = None,
    is_confirmed: Optional[int] = None,
    search: Optional[str] = None,
    admin_emp_id: Optional[str] = None
):
    """특근 목록 조회 (사용자 본인 또는 관리자 필터 조회 - 팀관리자는 슈퍼관리자 일정 열람 불가 및 본인 팀만 조회)"""
    conn = get_db_connection()
    cursor = conn.cursor()

    is_admin_viewer = False
    caller = get_cached_user_role(admin_emp_id) if admin_emp_id else None
    if caller and (caller.get("is_super") == 1 or caller.get("is_admin") == 1):
        is_admin_viewer = True
    if caller and caller.get("is_super") != 1 and caller.get("is_admin") == 1:
        team = caller.get("team")

    query = "SELECT * FROM overtimes WHERE 1=1"
    params = []

    # 요구사항 2: 팀관리자는 슈퍼관리자의 특근 일정을 열람할 수 없음
    if caller and caller.get("is_super") != 1 and caller.get("is_admin") == 1:
        query += " AND emp_id NOT IN (SELECT emp_id FROM users WHERE is_super = 1)"

    if emp_id:
        query += " AND emp_id = ?"
        params.append(emp_id.strip())
    if start_date:
        query += " AND end_date >= ?"
        params.append(start_date.strip())
    if end_date:
        query += " AND start_date <= ?"
        params.append(end_date.strip())
    if team:
        query += " AND team = ?"
        params.append(team.strip())
    if category:
        query += " AND category = ?"
        params.append(category.strip())
    if is_confirmed is not None:
        query += " AND is_confirmed = ?"
        params.append(is_confirmed)
    if search:
        s = f"%{search.strip()}%"
        query += " AND (user_name LIKE ? OR emp_id LIKE ? OR project_no LIKE ? OR reason LIKE ? OR location LIKE ?)"
        params.extend([s, s, s, s, s])

    query += " ORDER BY start_date DESC, id DESC"
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()

    # 요구사항 4: 일반 사용자는 보너스 부여 정보를 전혀 몰라야 함 (은닉/마스킹)
    if not is_admin_viewer:
        for r in rows:
            r["bonus_granted"] = 0

    return {"overtimes": rows}

@app.post("/api/overtimes")
def create_overtime(req: OvertimeCreateRequest):
    """신규 특근 신청 저장 (팀원 본인 신청 또는 관리자 대리 신청)"""
    conn = get_db_connection()
    cursor = conn.cursor()

    # 회원 정보 확인
    cursor.execute("SELECT * FROM users WHERE emp_id = ?", (req.emp_id.strip(),))
    user = cursor.fetchone()
    if not user:
        conn.close()
        raise HTTPException(status_code=400, detail="유효하지 않은 사번입니다.")

    user_name = user["name"]
    team = user["team"]
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 요구사항 5 & 4: 동일 사원의 동일/겹치는 일자 중복 신청 원천 차단
    cursor.execute("""
    SELECT id, category, start_date, end_date FROM overtimes
    WHERE emp_id = ? AND start_date <= ? AND end_date >= ?
    """, (req.emp_id.strip(), req.end_date.strip(), req.start_date.strip()))
    conflict = cursor.fetchone()
    if conflict:
        conn.close()
        raise HTTPException(
            status_code=400,
            detail=f"이미 해당 기간({conflict['start_date']} ~ {conflict['end_date']}, [{conflict['category']}])에 신청된 특근 내역이 존재합니다. 중복 신청은 불가능합니다."
        )

    bonus_granted_val = int(req.bonus_granted or 0)
    is_pre_deduct_val = int(req.is_pre_deduct or 0)
    trip_start_val = (req.trip_start_date or "").strip()
    trip_end_val = (req.trip_end_date or "").strip()

    cursor.execute("""
    INSERT INTO overtimes (
        emp_id, user_name, team, category, start_date, end_date,
        project_no, location, reason, sub_holiday_used, sub_holiday_date, is_confirmed,
        is_pre_deduct, trip_start_date, trip_end_date,
        created_at, updated_at, bonus_granted
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?, ?, ?)
    """, (
        req.emp_id.strip(),
        user_name,
        team,
        req.category.strip(),
        req.start_date.strip(),
        req.end_date.strip(),
        (req.project_no or "").strip(),
        (req.location or "").strip(),
        (req.reason or "").strip(),
        float(req.sub_holiday_used or 0),
        (req.sub_holiday_date or "").strip(),
        is_pre_deduct_val,
        trip_start_val,
        trip_end_val,
        now_str,
        now_str,
        bonus_granted_val
    ))
    new_id = cursor.lastrowid
    conn.commit()

    # Turso 등 HTTP 드라이버에서 lastrowid가 None인 경우 직전 삽입된 ID 조회
    if not new_id:
        try:
            cursor.execute("SELECT id FROM overtimes WHERE emp_id = ? ORDER BY id DESC LIMIT 1", (req.emp_id.strip(),))
            last_row = cursor.fetchone()
            if last_row:
                new_id = last_row["id"]
        except Exception as e:
            print(f"[create_overtime fallback warning] {e}")

    new_record = None
    if new_id:
        try:
            cursor.execute("SELECT * FROM overtimes WHERE id = ?", (new_id,))
            fetched = cursor.fetchone()
            if fetched:
                new_record = dict(fetched)
        except Exception as e:
            print(f"[create_overtime fetch warning] {e}")

    conn.close()

    # user_preferences 자동 업데이트 (프로젝트번호, 장소)
    proj_val = (req.project_no or "").strip()
    loc_val = (req.location or "").strip()
    if proj_val or loc_val:
        try:
            pref_conn = get_db_connection()
            pref_cursor = pref_conn.cursor()
            pref_cursor.execute("""
            INSERT INTO user_preferences (emp_id, last_project_no, last_location, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(emp_id) DO UPDATE SET
                last_project_no = excluded.last_project_no,
                last_location = excluded.last_location,
                updated_at = excluded.updated_at
            """, (req.emp_id.strip(), proj_val, loc_val, now_str))
            pref_conn.commit()
            pref_conn.close()
        except Exception as e:
            print(f"[user_preferences update warning] {e}")

    if not new_record:
        # DB 복제 지연 또는 fetch 실패 시에도 정상 응답을 보장하는 스마트 폴백
        new_record = {
            "id": new_id or 0,
            "emp_id": req.emp_id.strip(),
            "user_name": user_name,
            "team": team,
            "category": req.category.strip(),
            "start_date": req.start_date.strip(),
            "end_date": req.end_date.strip(),
            "project_no": (req.project_no or "").strip(),
            "location": (req.location or "").strip(),
            "reason": (req.reason or "").strip(),
            "sub_holiday_used": float(req.sub_holiday_used or 0),
            "sub_holiday_date": (req.sub_holiday_date or "").strip(),
            "is_confirmed": 0,
            "created_at": now_str,
            "updated_at": now_str,
            "bonus_granted": bonus_granted_val
        }

    # 감사 로그 비동기 기록
    log_audit(
        overtime_id=new_id or 0,
        action="신청",
        changed_by=req.emp_id.strip(),
        changed_by_name=user_name,
        previous_data=None,
        new_data=new_record
    )

    return {"message": "특근 신청이 완료되었습니다.", "overtime": new_record}

@app.get("/api/overtimes/summary")
def get_overtime_summary(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    team: Optional[str] = None,
    admin_emp_id: Optional[str] = None
):
    """특정 기간 동안 인원별/팀별 최종 특근일 정산 요약 (팀관리자는 본인 팀만 요약, 슈퍼관리자 제외)"""
    conn = get_db_connection()
    cursor = conn.cursor()

    query = "SELECT * FROM overtimes WHERE 1=1"
    params = []

    # 팀관리자인 경우 본인 소속팀 강제 고정 및 슈퍼관리자 일정 차단
    if admin_emp_id:
        caller = get_cached_user_role(admin_emp_id)
        if caller and caller.get("is_super") != 1 and caller.get("is_admin") == 1:
            team = caller.get("team")
            query += " AND emp_id NOT IN (SELECT emp_id FROM users WHERE is_super = 1)"

    if start_date:
        query += " AND end_date >= ?"
        params.append(start_date.strip())
    if end_date:
        query += " AND start_date <= ?"
        params.append(end_date.strip())
    if team:
        query += " AND team = ?"
        params.append(team.strip())

    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()

    # 직급 정보 조회
    user_positions = {}
    if rows:
        p_conn = get_db_connection()
        p_cur = p_conn.cursor()
        p_cur.execute("SELECT emp_id, position FROM users")
        for p in p_cur.fetchall():
            user_positions[p["emp_id"]] = p["position"]
        p_conn.close()

    user_list = aggregate_user_holidays(rows, user_positions)
    team_list = aggregate_team_holidays(user_list)

    # 이전 버전 호환성을 위한 키 매핑
    for u in user_list:
        u["overtime_days"] = u["normal_holiday_days"]
        u["excluded_sub_days"] = u["sub_work_days"]
        u["excluded_legal_days"] = u["legal_holiday_days"]
        u["pre_deduct_days"] = u["pre_deduct_count"]

    for tm in team_list:
        tm["excluded_days"] = tm["sub_work_days"] + tm["legal_holiday_days"]
        tm["overtime_days"] = tm["normal_holiday_days"]

    return {
        "user_summary": user_list,
        "team_summary": team_list,
        "total_records": len(rows)
    }

@app.get("/api/overtimes/{item_id}")
def get_overtime_detail(item_id: int, viewer_emp_id: Optional[str] = None):
    """특근 단건 상세 및 수정 이력 조회 (일반 사용자 조회 시 보너스 은닉)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    item = cursor.fetchone()
    if not item:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 특근 내역이 없습니다.")

    cursor.execute("""
    SELECT * FROM overtime_history WHERE overtime_id = ? ORDER BY id DESC
    """, (item_id,))
    history = [dict(h) for h in cursor.fetchall()]
    
    is_admin_viewer = False
    if viewer_emp_id:
        u_chk = get_cached_user_role(viewer_emp_id)
        if u_chk and (u_chk.get("is_super") == 1 or u_chk.get("is_admin") == 1):
            is_admin_viewer = True

    item_dict = dict(item)
    if not is_admin_viewer:
        item_dict["bonus_granted"] = 0

    conn.close()
    return {"overtime": item_dict, "history": history}

@app.put("/api/overtimes/{item_id}")
def update_overtime(item_id: int, req: OvertimeUpdateRequest):
    """특근 내역 수정 (본인, 슈퍼관리자, 또는 소속팀 팀관리자만 가능)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 특근 내역이 없습니다.")

    prev_data = dict(row)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 권한 검증 (요구사항 2)
    cursor.execute("SELECT name, is_super, is_admin, team FROM users WHERE emp_id = ?", (req.changed_by.strip(),))
    caller = cursor.fetchone()
    is_owner = (req.changed_by.strip() == prev_data["emp_id"])
    is_super = caller and (caller["is_super"] == 1 or req.changed_by.lower() == "ps37082")
    is_team_admin = caller and (caller["is_admin"] == 1 and caller["team"] == prev_data["team"])

    if not (is_owner or is_super or is_team_admin):
        conn.close()
        raise HTTPException(status_code=403, detail="본인 또는 소속팀 관리자만 수정할 수 있습니다.")

    # 요구사항 1: 승인(is_confirmed == 1)된 특근은 관리자만 수정 가능, 일반 사용자는 수정 불가 차단
    if prev_data.get("is_confirmed") == 1 and not (is_super or is_team_admin):
        conn.close()
        raise HTTPException(status_code=403, detail="승인 완료된 특근은 관리자만 수정할 수 있습니다.")

    changed_name = caller["name"] if caller else req.changed_by

    # 신규 필드 처리
    is_pre_deduct_val = int(req.is_pre_deduct) if req.is_pre_deduct is not None else int(prev_data.get("is_pre_deduct") or 0)
    trip_start_val = (req.trip_start_date or "").strip() if req.trip_start_date is not None else (prev_data.get("trip_start_date") or "")
    trip_end_val = (req.trip_end_date or "").strip() if req.trip_end_date is not None else (prev_data.get("trip_end_date") or "")

    bonus_clause = ""
    bonus_params = []
    if (is_super or is_team_admin) and req.bonus_granted is not None:
        bonus_clause = ", bonus_granted = ?"
        bonus_params.append(int(req.bonus_granted))

    extra_status_clauses = ""
    extra_status_params = []
    if (is_super or is_team_admin):
        if req.is_confirmed is not None:
            extra_status_clauses += ", is_confirmed = ?"
            extra_status_params.append(int(req.is_confirmed))
            if int(req.is_confirmed) == 1 and not prev_data.get("confirmed_by"):
                extra_status_clauses += ", confirmed_by = ?, confirmed_at = ?"
                extra_status_params.extend([f"{changed_name}({req.changed_by.strip()})", now_str])
            elif int(req.is_confirmed) == 0:
                extra_status_clauses += ", confirmed_by = NULL, confirmed_at = NULL"
        if req.is_finalized is not None:
            extra_status_clauses += ", is_finalized = ?"
            extra_status_params.append(int(req.is_finalized))
            if int(req.is_finalized) == 1 and not prev_data.get("finalized_by"):
                extra_status_clauses += ", finalized_by = ?, finalized_at = ?"
                extra_status_params.extend([f"{changed_name}({req.changed_by.strip()})", now_str])
            elif int(req.is_finalized) == 0:
                extra_status_clauses += ", finalized_by = NULL, finalized_at = NULL"
        if req.is_reviewed is not None:
            extra_status_clauses += ", is_reviewed = ?"
            extra_status_params.append(int(req.is_reviewed))
            if int(req.is_reviewed) == 1 and not prev_data.get("reviewed_by"):
                extra_status_clauses += ", reviewed_by = ?, reviewed_at = ?"
                extra_status_params.extend([f"{changed_name}({req.changed_by.strip()})", now_str])
            elif int(req.is_reviewed) == 0:
                extra_status_clauses += ", reviewed_by = NULL, reviewed_at = NULL"

    cursor.execute(f"""
    UPDATE overtimes SET
        category = ?, start_date = ?, end_date = ?,
        project_no = ?, location = ?, reason = ?,
        sub_holiday_used = ?, sub_holiday_date = ?,
        is_pre_deduct = ?, trip_start_date = ?, trip_end_date = ?,
        updated_at = ?
        {bonus_clause}
        {extra_status_clauses}
    WHERE id = ?
    """, (
        req.category.strip(),
        req.start_date.strip(),
        req.end_date.strip(),
        (req.project_no or "").strip(),
        (req.location or "").strip(),
        (req.reason or "").strip(),
        float(req.sub_holiday_used or 0),
        (req.sub_holiday_date or "").strip(),
        is_pre_deduct_val,
        trip_start_val,
        trip_end_val,
        now_str,
        *bonus_params,
        *extra_status_params,
        item_id
    ))
    conn.commit()

    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    fetched = cursor.fetchone()
    new_data = dict(fetched) if fetched else prev_data
    conn.close()

    # 감사 로그 기록
    log_audit(
        overtime_id=item_id,
        action="수정",
        changed_by=req.changed_by.strip(),
        changed_by_name=changed_name,
        previous_data=prev_data,
        new_data=new_data
    )

    return {"message": "수정되었습니다.", "overtime": new_data}

@app.delete("/api/overtimes/{item_id}")
def delete_overtime(item_id: int, changed_by: str = Query(...)):
    """특근 내역 삭제 (본인, 슈퍼관리자, 또는 소속팀 팀관리자만 가능)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 특근 내역이 없습니다.")

    prev_data = dict(row)

    # 권한 검증 (요구사항 2)
    cursor.execute("SELECT name, is_super, is_admin, team FROM users WHERE emp_id = ?", (changed_by.strip(),))
    caller = cursor.fetchone()
    is_owner = (changed_by.strip() == prev_data["emp_id"])
    is_super = caller and (caller["is_super"] == 1 or changed_by.lower() == "ps37082")
    is_team_admin = caller and (caller["is_admin"] == 1 and caller["team"] == prev_data["team"])

    if not (is_owner or is_super or is_team_admin):
        conn.close()
        raise HTTPException(status_code=403, detail="본인 또는 소속팀 관리자만 삭제할 수 있습니다.")

    # 요구사항 1: 승인(is_confirmed == 1)된 특근은 관리자만 삭제 가능, 일반 사용자는 삭제 불가 차단
    if prev_data.get("is_confirmed") == 1 and not (is_super or is_team_admin):
        conn.close()
        raise HTTPException(status_code=403, detail="승인 완료된 특근은 관리자만 삭제할 수 있습니다.")

    changed_name = caller["name"] if caller else changed_by

    cursor.execute("DELETE FROM overtimes WHERE id = ?", (item_id,))
    conn.commit()
    conn.close()

    # 감사 로그 기록
    log_audit(
        overtime_id=item_id,
        action="삭제",
        changed_by=changed_by.strip(),
        changed_by_name=changed_name,
        previous_data=prev_data,
        new_data=None
    )

    return {"message": "삭제되었습니다."}

@app.post("/api/overtimes/{item_id}/confirm")
def confirm_overtime(item_id: int, req: OvertimeConfirmRequest):
    """관리자 확인(승인) 상태 토글 (팀관리자는 본인 소속팀만 승인 가능)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 내역을 찾을 수 없습니다.")

    prev_data = dict(row)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 관리자 정보 및 2등급 권한 검증
    cursor.execute("SELECT name, is_super, is_admin, team FROM users WHERE emp_id = ?", (req.admin_emp_id.strip(),))
    u_row = cursor.fetchone()
    if not u_row or (u_row["is_super"] != 1 and u_row["is_admin"] != 1):
        conn.close()
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다.")

    # 팀관리자(is_super=0)는 본인 소속팀의 특근만 승인 가능
    if u_row["is_super"] != 1:
        if prev_data["team"] != u_row["team"]:
            conn.close()
            raise HTTPException(status_code=403, detail=f"팀관리자는 본인 소속팀({u_row['team']})의 특근만 승인할 수 있습니다.")

    admin_name = u_row["name"] if u_row else req.admin_emp_id

    if req.is_confirmed == 1:
        cursor.execute("""
        UPDATE overtimes SET is_confirmed = 1, confirmed_by = ?, confirmed_at = ?, updated_at = ?
        WHERE id = ?
        """, (f"{admin_name}({req.admin_emp_id})", now_str, now_str, item_id))
        action_name = "확인완료"
    else:
        cursor.execute("""
        UPDATE overtimes SET is_confirmed = 0, confirmed_by = NULL, confirmed_at = NULL, updated_at = ?
        WHERE id = ?
        """, (now_str, item_id))
        action_name = "확인취소"

    conn.commit()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    new_data = dict(cursor.fetchone())
    conn.close()

    # 감사 로그 기록
    log_audit(
        overtime_id=item_id,
        action=action_name,
        changed_by=req.admin_emp_id.strip(),
        changed_by_name=admin_name,
        previous_data=prev_data,
        new_data=new_data
    )

    return {"message": f"상태가 '{action_name}'(으)로 변경되었습니다.", "overtime": new_data}

@app.post("/api/overtimes/{item_id}/pre-deduct")
def toggle_overtime_pre_deduct(item_id: int, req: OvertimePreDeductRequest):
    """리스트에서 사전차감 상태 원클릭 토글 (본인 또는 관리자)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 내역을 찾을 수 없습니다.")

    prev_data = dict(row)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 권한 확인
    caller = get_cached_user_role(req.admin_emp_id)
    is_super = caller and caller.get("is_super") == 1
    is_admin = caller and (caller.get("is_super") == 1 or caller.get("is_admin") == 1)
    is_owner = (prev_data["emp_id"] == req.admin_emp_id.strip())

    if not (is_admin or is_owner):
        conn.close()
        raise HTTPException(status_code=403, detail="본인 또는 관리자만 사전차감을 변경할 수 있습니다.")

    if is_admin and not is_super and not is_owner:
        if prev_data["team"] != caller.get("team"):
            conn.close()
            raise HTTPException(status_code=403, detail="팀관리자는 본인 소속팀의 특근만 수정할 수 있습니다.")

    caller_name = caller["name"] if caller else req.admin_emp_id
    new_val = 1 if req.is_pre_deduct == 1 else 0

    cursor.execute("""
    UPDATE overtimes SET is_pre_deduct = ?, updated_at = ?
    WHERE id = ?
    """, (new_val, now_str, item_id))
    conn.commit()

    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    new_data = dict(cursor.fetchone())
    conn.close()

    log_audit(
        overtime_id=item_id,
        action="사전차감" if new_val == 1 else "사전차감취소",
        changed_by=req.admin_emp_id.strip(),
        changed_by_name=caller_name,
        previous_data=prev_data,
        new_data=new_data
    )

    return {"message": "사전차감 상태가 변경되었습니다.", "is_pre_deduct": new_val, "overtime": new_data}

@app.post("/api/overtimes/batch-confirm")
def batch_confirm_overtimes(payload: dict):
    """선택된 다중 특근 일괄 확인/취소 (팀관리자는 본인 소속팀만 승인 가능)"""
    ids = payload.get("ids", [])
    admin_emp_id = payload.get("admin_emp_id", "")
    is_confirmed = payload.get("is_confirmed", 1)

    if not ids:
        raise HTTPException(status_code=400, detail="선택된 항목이 없습니다.")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name, is_super, is_admin, team FROM users WHERE emp_id = ?", (admin_emp_id,))
    u_row = cursor.fetchone()
    if not u_row or (u_row["is_super"] != 1 and u_row["is_admin"] != 1):
        conn.close()
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다.")

    is_team_admin = (u_row["is_super"] != 1)
    admin_team = u_row["team"]
    admin_name = u_row["name"] if u_row else admin_emp_id
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for itm_id in ids:
        cursor.execute("SELECT team FROM overtimes WHERE id = ?", (itm_id,))
        ot_chk = cursor.fetchone()
        if not ot_chk:
            continue
        if is_team_admin and ot_chk["team"] != admin_team:
            continue  # 타 팀 특근은 팀관리자가 일괄 승인 시 자동 건너뜀

        if is_confirmed == 1:
            cursor.execute("""
            UPDATE overtimes SET is_confirmed = 1, confirmed_by = ?, confirmed_at = ?, updated_at = ?
            WHERE id = ?
            """, (f"{admin_name}({admin_emp_id})", now_str, now_str, itm_id))
        else:
            cursor.execute("""
            UPDATE overtimes SET is_confirmed = 0, confirmed_by = NULL, confirmed_at = NULL, updated_at = ?
            WHERE id = ?
            """, (now_str, itm_id))
    conn.commit()
    conn.close()

    return {"message": f"{len(ids)}건이 일괄 처리되었습니다."}

# ----------------- v1.44: 특근 완료 확정 및 검토완료 API -----------------

@app.post("/api/overtimes/{item_id}/finalize")
def finalize_overtime(item_id: int, req: OvertimeFinalizeRequest):
    """실제 특근 완료 시 사원 또는 관리자가 확정 상태 피드백 (1: 특근완료 확정, 0: 확정 취소)
    요구사항: 특근 신청 후 관리자 승인이 없어도 사원/관리자가 확정 가능하며 승인과 확정은 별개 독립 관리됨.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 특근 내역을 찾을 수 없습니다.")

    prev_data = dict(row)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    actor_id = (req.emp_id or req.actor_emp_id or "").strip()
    if not actor_id:
        conn.close()
        raise HTTPException(status_code=400, detail="사원번호가 전달되지 않았습니다.")

    # 권한 검증: 본인 또는 관리자만 확정 가능
    caller = get_cached_user_role(actor_id)
    is_admin = caller and (caller.get("is_super") == 1 or caller.get("is_admin") == 1)
    is_owner = (prev_data["emp_id"] == actor_id)

    if not (is_admin or is_owner):
        conn.close()
        raise HTTPException(status_code=403, detail="본인 또는 관리자만 특근 확정을 처리할 수 있습니다.")

    if is_admin and caller.get("is_super") != 1 and not is_owner:
        if prev_data["team"] != caller.get("team"):
            conn.close()
            raise HTTPException(status_code=403, detail="팀관리자는 본인 소속팀의 특근만 확정할 수 있습니다.")

    # 검토완료(is_reviewed=1)된 건은 누구도 확정 상태를 변경할 수 없음 (먼저 검토취소 필요)
    if prev_data.get("is_reviewed") == 1:
        conn.close()
        raise HTTPException(status_code=400, detail="검토완료된 특근은 확정 상태를 변경할 수 없습니다. 먼저 '검토취소'를 진행하세요.")

    caller_name = caller["name"] if caller else actor_id

    if req.is_finalized == 1:
        cursor.execute("""
        UPDATE overtimes SET is_finalized = 1, finalized_by = ?, finalized_at = ?, updated_at = ?
        WHERE id = ?
        """, (f"{caller_name}({actor_id})", now_str, now_str, item_id))
        action_name = "특근확정"
    else:
        cursor.execute("""
        UPDATE overtimes SET is_finalized = 0, finalized_by = NULL, finalized_at = NULL, updated_at = ?
        WHERE id = ?
        """, (now_str, item_id))
        action_name = "확정취소"

    conn.commit()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    new_data = dict(cursor.fetchone())
    conn.close()

    log_audit(
        overtime_id=item_id,
        action=action_name,
        changed_by=actor_id,
        changed_by_name=caller_name,
        previous_data=prev_data,
        new_data=new_data
    )

    return {"message": f"특근 상태가 '{action_name}'(으)로 변경되었습니다.", "overtime": new_data}


@app.post("/api/overtimes/batch-finalize")
def batch_finalize_overtimes(payload: OvertimeBatchFinalizeRequest):
    """선택된 다중 특근 일괄 확정/취소"""
    ids = payload.ids
    actor_id = (payload.emp_id or payload.actor_emp_id or "").strip()
    is_finalized = payload.is_finalized

    if not ids:
        raise HTTPException(status_code=400, detail="선택된 항목이 없습니다.")
    if not actor_id:
        raise HTTPException(status_code=400, detail="사원번호가 전달되지 않았습니다.")

    conn = get_db_connection()
    cursor = conn.cursor()
    caller = get_cached_user_role(actor_id)
    is_admin = caller and (caller.get("is_super") == 1 or caller.get("is_admin") == 1)
    caller_name = caller["name"] if caller else actor_id
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    processed = 0
    for itm_id in ids:
        cursor.execute("SELECT * FROM overtimes WHERE id = ?", (itm_id,))
        ot = cursor.fetchone()
        if not ot:
            continue
        ot_dict = dict(ot)
        is_owner = (ot_dict["emp_id"] == actor_id)
        if not (is_admin or is_owner):
            continue
        if is_admin and caller.get("is_super") != 1 and not is_owner:
            if ot_dict["team"] != caller.get("team"):
                continue

        # 검토완료된 건은 확정 상태 변경 불가 (일괄 처리 시 skip)
        if ot_dict.get("is_reviewed") == 1:
            continue

        if is_finalized == 1:
            cursor.execute("""
            UPDATE overtimes SET is_finalized = 1, finalized_by = ?, finalized_at = ?, updated_at = ?
            WHERE id = ?
            """, (f"{caller_name}({actor_id})", now_str, now_str, itm_id))
        else:
            cursor.execute("""
            UPDATE overtimes SET is_finalized = 0, finalized_by = NULL, finalized_at = NULL, updated_at = ?
            WHERE id = ?
            """, (now_str, itm_id))
        processed += 1

    conn.commit()
    conn.close()
    return {"message": f"{processed}건의 특근이 성공적으로 일괄 확정/취소 처리되었습니다."}


@app.post("/api/overtimes/{item_id}/review")
def review_overtime(item_id: int, req: OvertimeReviewRequest):
    """관리자가 특근에 대해 최종 검토완료 처리 (1: 검토완료, 0: 검토 취소)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="해당 특근 내역을 찾을 수 없습니다.")

    prev_data = dict(row)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    admin_emp_id = (req.admin_emp_id or req.actor_emp_id or req.emp_id or "").strip()
    if not admin_emp_id:
        conn.close()
        raise HTTPException(status_code=400, detail="관리자 사원번호가 전달되지 않았습니다.")

    # 관리자 권한 필수
    cursor.execute("SELECT name, is_super, is_admin, team FROM users WHERE emp_id = ?", (admin_emp_id,))
    u_row = cursor.fetchone()
    if not u_row or (u_row["is_super"] != 1 and u_row["is_admin"] != 1):
        conn.close()
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다.")

    if u_row["is_super"] != 1:
        if prev_data["team"] != u_row["team"]:
            conn.close()
            raise HTTPException(status_code=403, detail=f"팀관리자는 본인 소속팀({u_row['team']})의 특근만 검토할 수 있습니다.")

    admin_name = u_row["name"] if u_row else admin_emp_id

    if req.is_reviewed == 1:
        # 검토완료 시 승인 + 확정 동시 자동 처리
        actor_label = f"{admin_name}({admin_emp_id})"
        # 이미 승인/확정된 것은 덮어 쓰지 않게: 기존값 유지
        confirmed_by_val = prev_data.get("confirmed_by") or actor_label
        confirmed_at_val = prev_data.get("confirmed_at") or now_str
        finalized_by_val = prev_data.get("finalized_by") or actor_label
        finalized_at_val = prev_data.get("finalized_at") or now_str
        cursor.execute("""
        UPDATE overtimes
        SET is_reviewed = 1, reviewed_by = ?, reviewed_at = ?,
            is_confirmed = 1, confirmed_by = ?, confirmed_at = ?,
            is_finalized = 1, finalized_by = ?, finalized_at = ?,
            updated_at = ?
        WHERE id = ?
        """, (actor_label, now_str,
              confirmed_by_val, confirmed_at_val,
              finalized_by_val, finalized_at_val,
              now_str, item_id))
        action_name = "검토완료(승인+확정 자동)"
    else:
        cursor.execute("""
        UPDATE overtimes SET is_reviewed = 0, reviewed_by = NULL, reviewed_at = NULL, updated_at = ?
        WHERE id = ?
        """, (now_str, item_id))
        action_name = "검토취소"

    conn.commit()
    cursor.execute("SELECT * FROM overtimes WHERE id = ?", (item_id,))
    new_data = dict(cursor.fetchone())
    conn.close()

    log_audit(
        overtime_id=item_id,
        action=action_name,
        changed_by=admin_emp_id,
        changed_by_name=admin_name,
        previous_data=prev_data,
        new_data=new_data
    )

    return {"message": f"특근 상태가 '{action_name}'(으)로 변경되었습니다.", "overtime": new_data}


@app.post("/api/overtimes/batch-review")
def batch_review_overtimes(payload: OvertimeBatchReviewRequest):
    """선택된 다중 특근 일괄 검토완료/취소"""
    ids = payload.ids
    admin_emp_id = (payload.admin_emp_id or payload.actor_emp_id or payload.emp_id or "").strip()
    is_reviewed = payload.is_reviewed

    if not ids:
        raise HTTPException(status_code=400, detail="선택된 항목이 없습니다.")
    if not admin_emp_id:
        raise HTTPException(status_code=400, detail="관리자 사원번호가 전달되지 않았습니다.")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name, is_super, is_admin, team FROM users WHERE emp_id = ?", (admin_emp_id,))
    u_row = cursor.fetchone()
    if not u_row or (u_row["is_super"] != 1 and u_row["is_admin"] != 1):
        conn.close()
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다.")

    is_team_admin = (u_row["is_super"] != 1)
    admin_team = u_row["team"]
    admin_name = u_row["name"] if u_row else admin_emp_id
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    processed = 0
    for itm_id in ids:
        cursor.execute("SELECT * FROM overtimes WHERE id = ?", (itm_id,))
        ot = cursor.fetchone()
        if not ot:
            continue
        ot_dict = dict(ot)
        if is_team_admin and ot_dict["team"] != admin_team:
            continue

        if is_reviewed == 1:
            # 검토완료 시 승인 + 확정 동시 자동 처리
            actor_label = f"{admin_name}({admin_emp_id})"
            confirmed_by_val = ot_dict.get("confirmed_by") or actor_label
            confirmed_at_val = ot_dict.get("confirmed_at") or now_str
            finalized_by_val = ot_dict.get("finalized_by") or actor_label
            finalized_at_val = ot_dict.get("finalized_at") or now_str
            cursor.execute("""
            UPDATE overtimes
            SET is_reviewed = 1, reviewed_by = ?, reviewed_at = ?,
                is_confirmed = 1, confirmed_by = ?, confirmed_at = ?,
                is_finalized = 1, finalized_by = ?, finalized_at = ?,
                updated_at = ?
            WHERE id = ?
            """, (actor_label, now_str,
                  confirmed_by_val, confirmed_at_val,
                  finalized_by_val, finalized_at_val,
                  now_str, itm_id))
        else:
            cursor.execute("""
            UPDATE overtimes SET is_reviewed = 0, reviewed_by = NULL, reviewed_at = NULL, updated_at = ?
            WHERE id = ?
            """, (now_str, itm_id))
        processed += 1

    conn.commit()
    conn.close()
    return {"message": f"{processed}건의 특근이 성공적으로 일괄 검토완료/취소 처리되었습니다."}


# ----------------- v1.44: 무기명 건의사항란 (소통 게시판) API -----------------

@app.get("/api/suggestions")
def list_suggestions(category: Optional[str] = None, status: Optional[str] = None):
    """무기명 건의사항 목록 조회 (사번/이름 완전 배제)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    query = "SELECT id, category, title, content, status, admin_reply, reply_at, created_at FROM suggestions WHERE 1=1"
    params = []
    if category and category.strip():
        query += " AND category = ?"
        params.append(category.strip())
    if status and status.strip():
        query += " AND status = ?"
        params.append(status.strip())
    query += " ORDER BY id DESC"
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return {"status": "success", "suggestions": rows}


@app.post("/api/suggestions")
def create_suggestion(req: SuggestionCreateRequest):
    """무기명 건의사항 등록 (100% 익명 보장, 사번/이름 저장 없음)"""
    title = req.title.strip()
    content = req.content.strip()
    category = req.category.strip() or "불편사항"

    if not title or not content:
        raise HTTPException(status_code=400, detail="제목과 내용을 모두 입력해주세요.")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
    INSERT INTO suggestions (category, title, content, status, admin_reply, reply_at, created_at)
    VALUES (?, ?, ?, '접수됨', '', '', ?)
    """, (category, title, content, now_str))
    conn.commit()
    new_id = cursor.lastrowid
    conn.close()
    return {
        "status": "success",
        "message": "소중한 의견이 무기명으로 안전하게 등록되었습니다.",
        "id": new_id
    }


@app.post("/api/suggestions/{s_id}/reply")
def reply_suggestion(s_id: int, req: SuggestionReplyRequest):
    """관리자 건의사항 답변 및 처리 상태 변경 (관리자 전용)"""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT name, is_super, is_admin FROM users WHERE emp_id = ?", (req.admin_emp_id.strip(),))
    u_row = cursor.fetchone()
    if not u_row or (u_row["is_super"] != 1 and u_row["is_admin"] != 1):
        conn.close()
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다.")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
    UPDATE suggestions SET status = ?, admin_reply = ?, reply_at = ?
    WHERE id = ?
    """, (req.status.strip() or "조치완료", req.admin_reply.strip(), now_str, s_id))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "건의사항 답변 및 상태가 저장되었습니다."}


@app.delete("/api/suggestions/{s_id}")
def delete_suggestion(s_id: int, admin_emp_id: str = Query(..., description="슈퍼관리자 사번")):
    """건의사항 삭제 (슈퍼관리자 전용)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT is_super FROM users WHERE emp_id = ?", (admin_emp_id.strip(),))
    u_row = cursor.fetchone()
    if not u_row or u_row["is_super"] != 1:
        conn.close()
        raise HTTPException(status_code=403, detail="슈퍼관리자만 삭제할 수 있습니다.")

    cursor.execute("DELETE FROM suggestions WHERE id = ?", (s_id,))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "건의사항이 삭제되었습니다."}

@app.post("/api/overtimes/batch-delete")
def batch_delete_overtimes(payload: dict):
    """선택된 다중 특근 일괄 삭제 (팀관리자는 본인 소속팀만 삭제 가능)"""
    ids = payload.get("ids", [])
    admin_emp_id = payload.get("admin_emp_id", "")

    if not ids:
        raise HTTPException(status_code=400, detail="선택된 항목이 없습니다.")

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name, is_super, is_admin, team FROM users WHERE emp_id = ?", (admin_emp_id,))
    u_row = cursor.fetchone()
    if not u_row or (u_row["is_super"] != 1 and u_row["is_admin"] != 1):
        conn.close()
        raise HTTPException(status_code=403, detail="관리자 권한이 필요합니다.")

    is_team_admin = (u_row["is_super"] != 1)
    admin_team = u_row["team"]
    admin_name = u_row["name"] if u_row else admin_emp_id

    deleted_count = 0
    for itm_id in ids:
        cursor.execute("SELECT * FROM overtimes WHERE id = ?", (itm_id,))
        ot_chk = cursor.fetchone()
        if not ot_chk:
            continue
        if is_team_admin and ot_chk["team"] != admin_team:
            continue  # 타 팀 특근은 팀관리자가 일괄 삭제 시 건너뜀

        prev_data = dict(ot_chk)
        cursor.execute("DELETE FROM overtimes WHERE id = ?", (itm_id,))
        log_audit(
            overtime_id=itm_id,
            action="일괄삭제",
            changed_by=admin_emp_id,
            changed_by_name=admin_name,
            previous_data=prev_data,
            new_data=None
        )
        deleted_count += 1

    conn.commit()
    conn.close()

# ----------------- 엑셀 가져오기 & 내보내기 -----------------

@app.post("/api/overtimes/import-excel")
async def import_overtimes_excel_endpoint(request: Request):
    """
    특근 엑셀 파일(휴일일자별_특근현황 시트)을 업로드받아 선택된 부서의 특근 데이터를 DB에 갱신/추가합니다.
    """
    file_bytes = None
    target_teams = None
    admin_emp_id = None

    try:
        payload = await request.json()
        file_b64 = payload.get("file_b64") or payload.get("file") or ""
        if file_b64:
            if "," in file_b64:
                file_b64 = file_b64.split(",")[1]
            file_bytes = base64.b64decode(file_b64)
        target_teams = payload.get("target_teams") or payload.get("team")
        admin_emp_id = payload.get("admin_emp_id")
    except Exception:
        try:
            file_bytes = await request.body()
            target_teams = request.query_params.get("target_teams") or request.query_params.get("team")
            admin_emp_id = request.query_params.get("admin_emp_id")
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"요청 데이터를 읽을 수 없습니다: {str(e)}")

    if not file_bytes:
        raise HTTPException(status_code=400, detail="업로드할 엑셀 파일 데이터가 없습니다.")

    res = import_overtimes_from_excel(file_bytes, target_teams=target_teams, admin_emp_id=admin_emp_id)
    return res


@app.post("/api/overtimes/export-settlement")
async def export_settlement(req: Request):
    """실특근 정산표 openpyxl 고품질 엑셀 내보내기 (정산 화면 전용)"""
    try:
        payload = await req.json()
    except Exception:
        payload = {}
    start_date = payload.get("start_date", "") or ""
    end_date   = payload.get("end_date", "")   or ""
    team       = payload.get("team", "")        or ""
    admin_emp_id = payload.get("admin_emp_id", "") or ""

    conn = get_db_connection()
    cursor = conn.cursor()
    query = "SELECT * FROM overtimes WHERE 1=1"
    params = []

    if admin_emp_id:
        cursor.execute("SELECT is_super, is_admin, team FROM users WHERE emp_id = ?", (admin_emp_id.strip(),))
        caller = cursor.fetchone()
        if caller and caller["is_super"] != 1 and caller["is_admin"] == 1:
            team = caller["team"]
            query += " AND emp_id NOT IN (SELECT emp_id FROM users WHERE is_super = 1)"

    if start_date:
        query += " AND end_date >= ?"
        params.append(start_date.strip())
    if end_date:
        query += " AND start_date <= ?"
        params.append(end_date.strip())
    if team:
        query += " AND team = ?"
        params.append(team.strip())
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()

    if not rows:
        raise HTTPException(status_code=400, detail="내보낼 정산 내역 데이터가 없습니다.")

    # 직급 정보 조회
    user_positions = {}
    if rows:
        p_conn = get_db_connection()
        p_cur = p_conn.cursor()
        p_cur.execute("SELECT emp_id, position FROM users")
        for p in p_cur.fetchall():
            user_positions[p["emp_id"]] = p["position"]
        p_conn.close()

    user_list = aggregate_user_holidays(rows, user_positions)
    team_list = aggregate_team_holidays(user_list)

    # 이전 버전 호환성을 위한 키 매핑
    for u in user_list:
        u["overtime_days"] = u["normal_holiday_days"]
        u["excluded_sub_days"] = u["sub_work_days"]
        u["excluded_legal_days"] = u["legal_holiday_days"]
        u["pre_deduct_days"] = u["pre_deduct_count"]

    for tm in team_list:
        tm["excluded_days"] = tm["sub_work_days"] + tm["legal_holiday_days"]
        tm["overtime_days"] = tm["normal_holiday_days"]

    summary_data = {
        "user_summary": user_list,
        "team_summary": team_list,
    }

    excel_buffer = generate_settlement_excel(
        summary_data=summary_data,
        start_date=start_date,
        end_date=end_date,
        team_filter=team
    )

    period = f"{start_date or '전체'}_{end_date or '전체'}"
    filename = f"실특근정산표_{period}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    from urllib.parse import quote
    encoded_filename = quote(filename)

    return Response(
        content=excel_buffer.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }
    )

@app.post("/api/overtimes/import-settlement")
async def import_settlement(req: Request):
    """
    관리자모드 특근현황_일자별개인별정산 엑셀 가져오기 API (JSON base64 페이로드 지원)
    - 선택된 부서(target_team)가 있는 경우 해당 부서 데이터만 처리
    - 타 부서들의 데이터는 절대 변경되지 않고 안전하게 처리 제외 (skipped_other_dept_count)
    """
    try:
        payload = await req.json()
    except Exception:
        raise HTTPException(status_code=400, detail="올바른 JSON 요청 데이터가 아닙니다.")

    file_base64 = payload.get("file_base64", "")
    target_team = payload.get("target_team", "") or ""
    admin_emp_id = payload.get("admin_emp_id", "") or ""
    filename = payload.get("filename", "import.xlsx")

    if not file_base64:
        raise HTTPException(status_code=400, detail="업로드된 엑셀 파일 데이터가 없습니다.")

    try:
        contents = base64.b64decode(file_base64)
        excel_bytes = io.BytesIO(contents)
        import openpyxl
        wb = openpyxl.load_workbook(excel_bytes, data_only=True)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 해석할 수 없습니다: {str(e)}")

    sheet_names = wb.sheetnames
    target_sheet = None
    for name in ["휴일일자별_특근현황", "특근신청_전체원장", "개인별_실특근_정산표"]:
        if name in sheet_names:
            target_sheet = wb[name]
            break
    if not target_sheet:
        target_sheet = wb.active

    rows_data = list(target_sheet.iter_rows(values_only=True))
    if not rows_data or len(rows_data) < 2:
        raise HTTPException(status_code=400, detail="엑셀 파일에 처리할 데이터 행이 없습니다.")

    header_idx = -1
    col_map = {}
    for idx, r in enumerate(rows_data[:10]):
        if not r:
            continue
        row_str = [str(cell or '').strip() for cell in r]
        for c_i, val in enumerate(row_str):
            if val in ["사번", "사원번호", "emp_id"]:
                header_idx = idx
                break
        if header_idx != -1:
            break

    if header_idx == -1:
        header_idx = 0

    header_row = [str(cell or '').strip() for cell in rows_data[header_idx]]
    for c_i, name in enumerate(header_row):
        col_map[name] = c_i

    def get_val(row, aliases, default=""):
        for a in aliases:
            if a in col_map and col_map[a] < len(row):
                v = row[col_map[a]]
                if v is not None and str(v).strip() != "":
                    return str(v).strip()
        return default

    data_rows = rows_data[header_idx + 1:]
    conn = get_db_connection()
    cursor = conn.cursor()

    total_rows = 0
    processed_count = 0
    skipped_other_dept = 0
    error_count = 0
    errors = []

    cursor.execute("SELECT emp_id, name, team FROM users")
    db_users = {r["emp_id"]: dict(r) for r in cursor.fetchall()}

    target_team_clean = (target_team or "").strip()

    try:
        with conn:
            for r_num, row in enumerate(data_rows, start=header_idx + 2):
                if not any(cell is not None and str(cell).strip() != '' for cell in row):
                    continue

                total_rows += 1
                emp_id = get_val(row, ["사번", "사원번호", "emp_id"])
                user_name = get_val(row, ["성명", "이름", "name"])
                team = get_val(row, ["소속팀", "부서", "team"])

                if not emp_id:
                    errors.append(f"[{r_num}행] 사원번호(사번) 정보가 누락되어 스킵되었습니다.")
                    error_count += 1
                    continue

                user_info = db_users.get(emp_id)
                actual_team = team or (user_info["team"] if user_info else "")

                # ★ 부서 현황에서 선택된 부서만 처리하고, 타 부서 데이터는 스킵
                if target_team_clean and actual_team and actual_team != target_team_clean:
                    skipped_other_dept += 1
                    continue

                # 신규 사용자인 경우 users 테이블에 자동 등록 (FK 제약조건 보호)
                if not user_info:
                    now_created = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    new_user_team = actual_team or target_team_clean or "미지정"
                    new_user_name = user_name or emp_id
                    cursor.execute("""
                        INSERT OR IGNORE INTO users (emp_id, name, team, position, is_admin, is_super, created_at)
                        VALUES (?, ?, ?, '팀원', 0, 0, ?)
                    """, (emp_id, new_user_name, new_user_team, now_created))
                    user_info = {"emp_id": emp_id, "name": new_user_name, "team": new_user_team}
                    db_users[emp_id] = user_info
                    actual_team = new_user_team

                start_date = get_val(row, ["휴일날짜", "시작일", "start_date"])
                end_date = get_val(row, ["종료일", "end_date"]) or start_date
                category = get_val(row, ["특근분류", "분류", "category"], "일반휴일")
                project_no = get_val(row, ["프로젝트 번호", "프로젝트번호", "project_no"])
                location = get_val(row, ["근무 장소", "근무장소", "location"])
                reason = get_val(row, ["특근 사유", "사유", "reason"])
                status_str = get_val(row, ["진행단계", "상태", "status"])
                bonus_str = get_val(row, ["보너스 부여", "보너스", "bonus_granted"])

                bonus_val = 1 if bonus_str in ["1", "Y", "예", "부여", "True"] else 0
                is_confirmed = 1 if status_str in ["승인", "확정", "검토완료", "검토", "1"] else 0
                now_iso = datetime.now().isoformat()

                if start_date:
                    cursor.execute("""
                        SELECT id FROM overtimes 
                        WHERE emp_id = ? AND start_date = ? AND category = ?
                    """, (emp_id, start_date, category))
                    existing = cursor.fetchone()

                    if existing:
                        cursor.execute("""
                            UPDATE overtimes 
                            SET user_name = COALESCE(NULLIF(?, ''), user_name),
                                team = COALESCE(NULLIF(?, ''), team),
                                end_date = ?,
                                project_no = ?,
                                location = ?,
                                reason = ?,
                                is_confirmed = ?,
                                bonus_granted = ?,
                                updated_at = ?
                            WHERE id = ?
                        """, (user_name, actual_team, end_date, project_no, location, reason, is_confirmed, bonus_val, now_iso, existing["id"]))
                    else:
                        cursor.execute("""
                            INSERT INTO overtimes (
                                emp_id, user_name, team, category, start_date, end_date,
                                project_no, location, reason, is_confirmed, bonus_granted,
                                created_at, updated_at
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            emp_id, user_name or (user_info["name"] if user_info else emp_id),
                            actual_team or "미지정", category, start_date, end_date,
                            project_no, location, reason, is_confirmed, bonus_val,
                            now_iso, now_iso
                        ))
                    processed_count += 1
                else:
                    processed_count += 1

    except Exception as ex:
        conn.rollback()
        raise HTTPException(status_code=500, detail=f"가져오기 DB 처리 중 오류 발생: {str(ex)}")

    conn.close()

    if total_rows > 0 and processed_count == 0 and skipped_other_dept > 0:
        return JSONResponse(status_code=400, content={
            "success": False,
            "message": f"선택된 부서('{target_team_clean}')에 해당하는 데이터가 없습니다. (타 부서 {skipped_other_dept}건 제외됨)",
            "total_rows": total_rows,
            "processed_count": 0,
            "skipped_other_dept_count": skipped_other_dept,
            "error_count": error_count,
            "errors": errors or [f"파일 내 모든 {skipped_other_dept}건이 선택된 부서('{target_team_clean}')와 달라 제외되었습니다."],
            "target_team": target_team_clean
        })

    return {
        "success": True,
        "message": f"정산표 엑셀 가져오기가 완료되었습니다.",
        "total_rows": total_rows,
        "processed_count": processed_count,
        "skipped_other_dept_count": skipped_other_dept,
        "error_count": error_count,
        "errors": errors,
        "target_team": target_team_clean
    }

@app.post("/api/overtimes/export")
def export_overtimes(req: ExportRequest):
    """다중 선택 또는 필터된 특근 내역 엑셀(.xlsx) 파일 스트리밍"""
    conn = get_db_connection()
    cursor = conn.cursor()

    if req.ids and len(req.ids) > 0:
        placeholders = ",".join("?" for _ in req.ids)
        query = f"SELECT * FROM overtimes WHERE id IN ({placeholders}) ORDER BY start_date DESC, id DESC"
        cursor.execute(query, req.ids)
    else:
        query = "SELECT * FROM overtimes WHERE 1=1"
        params = []
        if req.admin_emp_id:
            cursor.execute("SELECT is_super, is_admin, team FROM users WHERE emp_id = ?", (req.admin_emp_id.strip(),))
            caller = cursor.fetchone()
            if caller and caller["is_super"] != 1 and caller["is_admin"] == 1:
                req.team = caller["team"]
                query += " AND emp_id NOT IN (SELECT emp_id FROM users WHERE is_super = 1)"

        if req.start_date:
            query += " AND end_date >= ?"
            params.append(req.start_date.strip())
        if req.end_date:
            query += " AND start_date <= ?"
            params.append(req.end_date.strip())
        if req.team:
            query += " AND team = ?"
            params.append(req.team.strip())
        if req.category:
            query += " AND category = ?"
            params.append(req.category.strip())
        if req.is_confirmed is not None:
            query += " AND is_confirmed = ?"
            params.append(req.is_confirmed)
        if req.search:
            s = f"%{req.search.strip()}%"
            query += " AND (user_name LIKE ? OR emp_id LIKE ? OR project_no LIKE ? OR reason LIKE ?)"
            params.extend([s, s, s, s])
        query += " ORDER BY start_date DESC, id DESC"
        cursor.execute(query, params)

    records = [dict(r) for r in cursor.fetchall()]
    
    cursor.execute("SELECT emp_id, position FROM users")
    user_positions = {u["emp_id"]: u["position"] or "-" for u in cursor.fetchall()}
    conn.close()

    if not records:
        raise HTTPException(status_code=400, detail="내보낼 특근 내역 데이터가 없습니다.")

    excel_buffer = generate_overtime_excel(records, user_positions=user_positions)
    filename = f"특근현황_일자별및개인별정산_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    from urllib.parse import quote
    encoded_filename = quote(filename)

    return Response(
        content=excel_buffer.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }
    )

@app.get("/api/manual/download")
def download_manual():
    """PPT 매뉴얼 파일 다운로드"""
    if not PPTX_PATH.exists():
        create_manual()
    from urllib.parse import quote
    filename = "특근관리시스템_사용자_및_관리자_매뉴얼.pptx"
    encoded_filename = quote(filename)
    return FileResponse(
        str(PPTX_PATH),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }
    )

# ----------------- 웹 저장 및 열기 (스냅샷 백업/복원) -----------------
@app.post("/api/backup/save")
def save_web_backup(req: BackupSaveRequest):
    """웹상으로 현재 전체 DB 데이터를 JSON 스냅샷으로 저장"""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM teams")
    teams = [dict(r) for r in cursor.fetchall()]

    cursor.execute("SELECT * FROM users")
    users = [dict(r) for r in cursor.fetchall()]

    cursor.execute("SELECT * FROM overtimes")
    overtimes = [dict(r) for r in cursor.fetchall()]

    cursor.execute("SELECT * FROM overtime_history")
    history = [dict(r) for r in cursor.fetchall()]

    conn.close()

    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    custom_name = (req.name or f"백업_{now_str}").strip()
    safe_name = "".join(c for c in custom_name if c.isalnum() or c in ("-", "_", " ")).strip()
    if not safe_name:
        safe_name = f"backup_{now_str}"
    filename = f"{now_str}_{safe_name}.json"
    file_path = WEB_BACKUP_DIR / filename

    payload = {
        "version": "1.3.0",
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "name": custom_name,
        "description": req.description or "",
        "counts": {
            "teams": len(teams),
            "users": len(users),
            "overtimes": len(overtimes),
            "history": len(history)
        },
        "data": {
            "teams": teams,
            "users": users,
            "overtimes": overtimes,
            "overtime_history": history
        }
    }

    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return {
        "status": "success",
        "filename": filename,
        "name": custom_name,
        "created_at": payload["created_at"],
        "counts": payload["counts"]
    }

@app.get("/api/backup/list")
def list_web_backups():
    """저장된 웹 백업 스냅샷 목록 조회"""
    backups = []
    if not WEB_BACKUP_DIR.exists():
        return {"backups": []}

    for p in sorted(WEB_BACKUP_DIR.glob("*.json"), key=os.path.getmtime, reverse=True):
        try:
            with open(p, "r", encoding="utf-8") as f:
                content = json.load(f)
            backups.append({
                "filename": p.name,
                "name": content.get("name", p.name),
                "created_at": content.get("created_at", ""),
                "description": content.get("description", ""),
                "counts": content.get("counts", {}),
                "file_size": p.stat().st_size
            })
        except Exception:
            continue

    return {"backups": backups}

@app.post("/api/backup/load")
def load_web_backup(req: BackupLoadRequest):
    """지정한 웹 백업 스냅샷을 현재 DB에 복원(열기)"""
    filename = os.path.basename(req.filename)
    file_path = WEB_BACKUP_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="지정한 백업 파일을 찾을 수 없습니다.")

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            snapshot = json.load(f)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"백업 파일 파싱 실패: {str(e)}")

    data = snapshot.get("data", {})
    teams = data.get("teams", [])
    users = data.get("users", [])
    overtimes = data.get("overtimes", [])
    history = data.get("overtime_history", [])

    # 복원 전 안전을 위해 자동 백업 수행
    create_backup()

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("PRAGMA foreign_keys = OFF;")

        cursor.execute("DELETE FROM overtime_history")
        cursor.execute("DELETE FROM overtimes")
        cursor.execute("DELETE FROM users")
        cursor.execute("DELETE FROM teams")

        # 1. teams 복원
        for t in teams:
            cursor.execute("""
            INSERT OR REPLACE INTO teams (id, name, created_at)
            VALUES (?, ?, ?)
            """, (t.get("id"), t["name"], t.get("created_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))))

        # 2. users 복원
        has_super = False
        for u in users:
            if u.get("is_super") == 1 or u["emp_id"].lower() == "ps37082":
                has_super = True
            cursor.execute("""
            INSERT OR REPLACE INTO users (emp_id, name, team, position, is_admin, is_super, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                u["emp_id"], u["name"], u.get("team", ""), u.get("position", ""),
                u.get("is_admin", 0), u.get("is_super", 0),
                u.get("created_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            ))

        if not has_super:
            cursor.execute("""
            INSERT OR REPLACE INTO users (emp_id, name, team, position, is_admin, is_super)
            VALUES ('ps37082', '관리자', 'IT운영팀', '관리자', 1, 1)
            """)

        # 3. overtimes 복원
        for o in overtimes:
            cursor.execute("""
            INSERT OR REPLACE INTO overtimes (
                id, emp_id, user_name, team, category, start_date, end_date,
                project_no, location, reason, is_confirmed,
                confirmed_by, confirmed_at, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                o.get("id"), o["emp_id"], o["user_name"], o["team"], o["category"],
                o["start_date"], o["end_date"], o.get("project_no", ""), o.get("location", ""),
                o.get("reason", ""), o.get("is_confirmed", 0),
                o.get("confirmed_by"), o.get("confirmed_at"),
                o.get("created_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                o.get("updated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            ))

        # 4. overtime_history 복원
        for h in history:
            cursor.execute("""
            INSERT OR REPLACE INTO overtime_history (
                id, overtime_id, action, changed_by, changed_by_name, previous_data, new_data, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                h.get("id"), h.get("overtime_id"), h.get("action", "UPDATE"),
                h.get("changed_by", "system"), h.get("changed_by_name", ""),
                h.get("previous_data", ""), h.get("new_data", ""),
                h.get("created_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            ))

        cursor.execute("PRAGMA foreign_keys = ON;")
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        raise HTTPException(status_code=500, detail=f"복원 중 오류 발생: {str(e)}")

    conn.close()

    return {
        "status": "success",
        "message": f"'{snapshot.get('name', filename)}' 스냅샷이 성공적으로 복원되었습니다.",
        "restored_counts": {
            "teams": len(teams),
            "users": len(users),
            "overtimes": len(overtimes),
            "history": len(history)
        }
    }


# ----------------- 슈퍼관리자 전용 보안 감사 및 접속 로그 API (신규 요구사항 5) -----------------

@app.get("/api/admin/access-logs")
def get_access_logs(
    admin_emp_id: str = Query(..., description="조회 요청자의 사원번호"),
    start_date: Optional[str] = Query(None, description="시작일 (YYYY-MM-DD)"),
    end_date: Optional[str] = Query(None, description="종료일 (YYYY-MM-DD)"),
    action_type: Optional[str] = Query(None, description="액션 타입 (LOGIN, REGISTER, LOGOUT 등)"),
    status: Optional[str] = Query(None, description="상태 (SUCCESS, FAILURE)"),
    search: Optional[str] = Query(None, description="사번, 이름, IP 검색어"),
    limit: int = Query(300, description="최대 조회 건수")
):
    """슈퍼관리자 전용 보안 감사 및 접속 로그 조회"""
    conn = get_db_connection()
    cursor = conn.cursor()

    # 슈퍼관리자 권한 확인
    cursor.execute("SELECT is_super FROM users WHERE emp_id = ?", (admin_emp_id.strip(),))
    admin_row = cursor.fetchone()
    if not admin_row or int(admin_row["is_super"] or 0) != 1:
        conn.close()
        raise HTTPException(status_code=403, detail="슈퍼관리자만 접근 가능한 보안 로그입니다.")

    # 1. 오늘 날짜 기준 통계 요약 산출
    today_prefix = datetime.now().strftime("%Y-%m-%d")
    cursor.execute("SELECT COUNT(*) as total_cnt FROM access_logs")
    total_cnt = cursor.fetchone()["total_cnt"]

    cursor.execute("SELECT COUNT(*) as today_cnt FROM access_logs WHERE created_at LIKE ?", (f"{today_prefix}%",))
    today_cnt = cursor.fetchone()["today_cnt"]

    cursor.execute("SELECT COUNT(*) as today_fail FROM access_logs WHERE status = 'FAILURE' AND created_at LIKE ?", (f"{today_prefix}%",))
    today_fail = cursor.fetchone()["today_fail"]

    cursor.execute("SELECT COUNT(*) as today_reg FROM access_logs WHERE action_type = 'REGISTER' AND status = 'SUCCESS' AND created_at LIKE ?", (f"{today_prefix}%",))
    today_reg = cursor.fetchone()["today_reg"]

    # 2. 동적 쿼리 구성
    where_clauses = ["1=1"]
    params = []

    if start_date:
        where_clauses.append("created_at >= ?")
        params.append(f"{start_date.strip()} 00:00:00")
    if end_date:
        where_clauses.append("created_at <= ?")
        params.append(f"{end_date.strip()} 23:59:59")
    if action_type and action_type.strip():
        where_clauses.append("action_type = ?")
        params.append(action_type.strip().upper())
    if status and status.strip():
        where_clauses.append("status = ?")
        params.append(status.strip().upper())
    if search and search.strip():
        s = f"%{search.strip()}%"
        where_clauses.append("(emp_id LIKE ? OR user_name LIKE ? OR ip_address LIKE ? OR details LIKE ?)")
        params.extend([s, s, s, s])

    query = f"""
        SELECT id, emp_id, user_name, action_type, status, ip_address, user_agent, details, created_at
        FROM access_logs
        WHERE {' AND '.join(where_clauses)}
        ORDER BY id DESC
        LIMIT ?
    """
    params.append(limit)

    cursor.execute(query, tuple(params))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()

    return {
        "status": "success",
        "stats": {
            "total_logs": total_cnt,
            "today_total": today_cnt,
            "today_failed": today_fail,
            "today_registered": today_reg
        },
        "logs": rows
    }


@app.get("/api/admin/access-logs/export")
def export_access_logs(
    admin_emp_id: str = Query(..., description="조회 요청자의 사원번호"),
    start_date: Optional[str] = Query(None),
    end_date: Optional[str] = Query(None),
    action_type: Optional[str] = Query(None),
    status: Optional[str] = Query(None),
    search: Optional[str] = Query(None)
):
    """슈퍼관리자 전용 보안 감사 로그 엑셀 다운로드 (.xlsx)"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT is_super FROM users WHERE emp_id = ?", (admin_emp_id.strip(),))
    admin_row = cursor.fetchone()
    if not admin_row or int(admin_row["is_super"] or 0) != 1:
        conn.close()
        raise HTTPException(status_code=403, detail="슈퍼관리자만 내보내기가 가능합니다.")

    where_clauses = ["1=1"]
    params = []

    if start_date:
        where_clauses.append("created_at >= ?")
        params.append(f"{start_date.strip()} 00:00:00")
    if end_date:
        where_clauses.append("created_at <= ?")
        params.append(f"{end_date.strip()} 23:59:59")
    if action_type and action_type.strip():
        where_clauses.append("action_type = ?")
        params.append(action_type.strip().upper())
    if status and status.strip():
        where_clauses.append("status = ?")
        params.append(status.strip().upper())
    if search and search.strip():
        s = f"%{search.strip()}%"
        where_clauses.append("(emp_id LIKE ? OR user_name LIKE ? OR ip_address LIKE ? OR details LIKE ?)")
        params.extend([s, s, s, s])

    query = f"""
        SELECT id, emp_id, user_name, action_type, status, ip_address, details, user_agent, created_at
        FROM access_logs
        WHERE {' AND '.join(where_clauses)}
        ORDER BY id DESC
        LIMIT 5000
    """
    cursor.execute(query, tuple(params))
    logs = [dict(r) for r in cursor.fetchall()]
    conn.close()

    # 워크북 생성 및 스타일 적용
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "접속및보안감사로그"
    ws.views.sheetView[0].showGridLines = True

    # 1. 대제목
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = "🛡️ 시스템 보안 및 사용자 접속 감사 로그 (Team Overtime Manager)"
    title_cell.font = Font(name="맑은 고딕", size=15, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # 2. 메타 정보
    ws.merge_cells("A2:I2")
    meta_cell = ws["A2"]
    meta_cell.value = f"출력일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  |  조회자: 슈퍼관리자({admin_emp_id})  |  총 {len(logs)}건"
    meta_cell.font = Font(name="맑은 고딕", size=10, italic=True, color="64748B")
    meta_cell.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[2].height = 20

    # 3. 헤더
    headers = ["순번", "기록 일시", "사원번호", "성명", "구분(액션)", "상태", "접속 IP", "상세 내용", "기기 환경(User-Agent)"]
    ws.row_dimensions[3].height = 26

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.value = h
        cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # 4. 데이터 채우기
    success_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # 연녹색
    fail_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")    # 연빨강
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    for i, log in enumerate(logs, 1):
        row_num = i + 3
        ws.row_dimensions[row_num].height = 22
        is_even = (i % 2 == 0)

        c1 = ws.cell(row=row_num, column=1, value=log["id"])
        c2 = ws.cell(row=row_num, column=2, value=log["created_at"])
        c3 = ws.cell(row=row_num, column=3, value=log["emp_id"])
        c4 = ws.cell(row=row_num, column=4, value=log["user_name"] or "-")
        c5 = ws.cell(row=row_num, column=5, value=log["action_type"])
        c6 = ws.cell(row=row_num, column=6, value=log["status"])
        c7 = ws.cell(row=row_num, column=7, value=log["ip_address"])
        c8 = ws.cell(row=row_num, column=8, value=log["details"])
        c9 = ws.cell(row=row_num, column=9, value=log["user_agent"])

        # 정렬
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c5.alignment = Alignment(horizontal="center", vertical="center")
        c6.alignment = Alignment(horizontal="center", vertical="center")
        c7.alignment = Alignment(horizontal="center", vertical="center")
        c8.alignment = Alignment(horizontal="left", vertical="center")
        c9.alignment = Alignment(horizontal="left", vertical="center")

        # 폰트
        for c in (c1, c2, c3, c4, c5, c6, c7, c8, c9):
            c.font = Font(name="맑은 고딕", size=9)
            c.border = thin_border
            if is_even:
                c.fill = zebra_fill

        # 상태별 강조 배지
        if log["status"] == "SUCCESS":
            c6.fill = success_fill
            c6.font = Font(name="맑은 고딕", size=9, bold=True, color="166534")
        else:
            c6.fill = fail_fill
            c6.font = Font(name="맑은 고딕", size=9, bold=True, color="991B1B")

    # 컬럼 너비 설정
    col_widths = {1: 8, 2: 20, 3: 13, 4: 12, 5: 14, 6: 12, 7: 17, 8: 35, 9: 45}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"접속및보안감사로그_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    encoded_filename = quote(filename)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename=\"{encoded_filename}\"; filename*=UTF-8''{encoded_filename}"
        }
    )


# ----------------- PPT 매뉴얼 다운로드 API (사용자용 / 관리자용 분리) -----------------



@app.get("/api/manual/user")
def download_user_manual():
    """사용자 모드 전용 매뉴얼 다운로드 (.pptx)"""
    if not USER_PPTX_PATH.exists():
        create_manual()
    filename = "특근관리시스템_사용자_매뉴얼(v1.45).pptx"
    encoded_filename = quote(filename)
    return FileResponse(
        str(USER_PPTX_PATH),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename=\"{encoded_filename}\"; filename*=UTF-8''{encoded_filename}"}
    )

@app.get("/api/manual/admin")
def download_admin_manual():
    """관리자 모드 전용 운영 매뉴얼 다운로드 (.pptx)"""
    if not ADMIN_PPTX_PATH.exists():
        create_manual()
    filename = "특근관리시스템_관리자_운영매뉴얼(v1.45).pptx"
    encoded_filename = quote(filename)
    return FileResponse(
        str(ADMIN_PPTX_PATH),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename=\"{encoded_filename}\"; filename*=UTF-8''{encoded_filename}"}
    )

@app.get("/api/manual/download")
def download_manual(type: str = Query("user", description="user 또는 admin")):
    """통합 매뉴얼 다운로드 엔드포인트 (기본값: user)"""
    if type.lower() == "admin":
        return download_admin_manual()
    return download_user_manual()



# ----------------- 정적 파일 호스팅 (루트/하위 폴더 자동 호환) -----------------
def sync_static_files():
    try:
        (BASE_DIR / "css").mkdir(exist_ok=True)
        (BASE_DIR / "js").mkdir(exist_ok=True)
        (BASE_DIR / "downloads").mkdir(exist_ok=True)

        if (BASE_DIR / "style.css").exists():
            shutil.copy(str(BASE_DIR / "style.css"), str(BASE_DIR / "css" / "style.css"))
            if STATIC_DIR.exists():
                shutil.copy(str(BASE_DIR / "style.css"), str(STATIC_DIR / "style.css"))

        if (BASE_DIR / "app.js").exists():
            shutil.copy(str(BASE_DIR / "app.js"), str(BASE_DIR / "js" / "app.js"))
            if STATIC_DIR.exists():
                (STATIC_DIR / "js").mkdir(parents=True, exist_ok=True)
                shutil.copy(str(BASE_DIR / "app.js"), str(STATIC_DIR / "js" / "app.js"))
                shutil.copy(str(BASE_DIR / "app.js"), str(STATIC_DIR / "app.js"))

        if (BASE_DIR / "index.html").exists():
            if STATIC_DIR.exists():
                shutil.copy(str(BASE_DIR / "index.html"), str(STATIC_DIR / "index.html"))
    except Exception as sync_e:
        print(f"[File Sync Notice] {sync_e}")

sync_static_files()

@app.get("/")
def serve_index():
    return FileResponse(str(BASE_DIR / "index.html"), headers={"Cache-Control": "no-cache, no-store, must-revalidate"})

@app.get("/favicon.ico")
def serve_favicon():
    return Response(status_code=204)

@app.get("/style.css")
@app.get("/css/style.css")
def serve_css():
    target = BASE_DIR / "style.css"
    if not target.exists():
        target = BASE_DIR / "css" / "style.css"
    return FileResponse(str(target), media_type="text/css", headers={"Cache-Control": "no-cache, no-store, must-revalidate"})

@app.get("/app.js")
@app.get("/js/app.js")
def serve_js():
    target = BASE_DIR / "app.js"
    if not target.exists():
        target = BASE_DIR / "js" / "app.js"
    return FileResponse(str(target), media_type="application/javascript", headers={"Cache-Control": "no-cache, no-store, must-revalidate"})

app.mount("/css", StaticFiles(directory=str(BASE_DIR / "css")), name="css")
app.mount("/js", StaticFiles(directory=str(BASE_DIR / "js")), name="js")
app.mount("/downloads", StaticFiles(directory=str(BASE_DIR / "downloads")), name="downloads")
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
