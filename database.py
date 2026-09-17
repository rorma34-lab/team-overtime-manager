import os
import sqlite3
import shutil
import json
import time
import threading
from queue import Queue
from datetime import datetime
from pathlib import Path
import urllib.parse
import http.client
import ssl

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
BACKUP_DIR = DATA_DIR / "backups"
DB_PATH = DATA_DIR / "overtime.db"

# Turso 클라우드 데이터베이스 설정 (Render 환경변수 연동)
TURSO_DATABASE_URL = os.environ.get("TURSO_DATABASE_URL") or os.environ.get("TURSO_URL")
TURSO_AUTH_TOKEN = os.environ.get("TURSO_AUTH_TOKEN") or os.environ.get("TURSO_TOKEN")

# 디렉토리 생성
DATA_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

def is_using_turso() -> bool:
    """Turso 클라우드 데이터베이스 사용 여부 확인"""
    return bool(TURSO_DATABASE_URL and TURSO_AUTH_TOKEN)

def get_db_mode() -> str:
    """현재 연결된 데이터베이스 모드 명칭 반환"""
    return "TURSO_CLOUD" if is_using_turso() else "LOCAL_SQLITE"

# -------------------------------------------------------------
# ⚡ 1. Turso HTTP Persistent Keep-Alive 세션 최적화 엔진
# -------------------------------------------------------------
_TURSO_CONNECTION_POOL = {}
_TURSO_POOL_LOCK = threading.Lock()

def _get_persistent_turso_client(base_url: str):
    """Turso 호스트에 대해 TLS 핸드셰이크를 매번 반복하지 않는 영구 Keep-Alive HTTPS 연결 유지"""
    parsed = urllib.parse.urlparse(base_url)
    host_key = f"{parsed.hostname}:{parsed.port or 443}"
    with _TURSO_POOL_LOCK:
        conn = _TURSO_CONNECTION_POOL.get(host_key)
        if conn is None:
            ctx = ssl.create_default_context()
            conn = http.client.HTTPSConnection(parsed.hostname, parsed.port or 443, context=ctx, timeout=12.0)
            _TURSO_CONNECTION_POOL[host_key] = conn
        return conn

def _patch_turso_keepalive():
    """turso_serverless 라이브러리의 _post 메서드를 Keep-Alive 영구 연결로 고속화"""
    try:
        import turso_serverless.session
        orig_post = turso_serverless.session.Session._post

        def fast_post(self, path: str, body: dict) -> bytes:
            payload = json.dumps(body, allow_nan=False).encode("utf-8")
            headers = self._headers()
            headers["Connection"] = "keep-alive"
            headers["Content-Length"] = str(len(payload))

            parsed = urllib.parse.urlparse(self._base_url)
            full_path = f"{parsed.path.rstrip('/')}{path}"
            if not full_path:
                full_path = "/"

            conn = _get_persistent_turso_client(self._base_url)
            try:
                conn.request("POST", full_path, body=payload, headers=headers)
                resp = conn.getresponse()
                if resp.status == 200:
                    return resp.read()
                raw = resp.read().decode("utf-8", errors="replace")
                self._reset_stream()
                raise RuntimeError(f"HTTP status {resp.status}: {raw}")
            except Exception as e:
                with _TURSO_POOL_LOCK:
                    _TURSO_CONNECTION_POOL.pop(f"{parsed.hostname}:{parsed.port or 443}", None)
                try:
                    conn.close()
                except Exception:
                    pass
                return orig_post(self, path, body)

        turso_serverless.session.Session._post = fast_post
        print("[Turso Accelerator] HTTP Keep-Alive persistent connection engine successfully attached.")
    except Exception as e:
        print(f"[Turso Accelerator Warning] Could not attach Keep-Alive patch: {e}")

if is_using_turso():
    _patch_turso_keepalive()

def get_db_connection():
    """DB 연결 객체 반환 (Turso 클라우드 DB 또는 로컬 고성능 SQLite 자동 선택)"""
    if is_using_turso():
        try:
            import turso_serverless
            conn = turso_serverless.connect(
                TURSO_DATABASE_URL.strip(),
                auth_token=TURSO_AUTH_TOKEN.strip()
            )
            conn.row_factory = turso_serverless.Row
            return conn
        except Exception as e:
            print(f"[Turso Connection Error] Fallback to local SQLite: {e}")

    # 로컬 SQLite 연결 (Check same thread 해제, busy timeout 15초)
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False, timeout=15.0)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA busy_timeout=15000;")
        conn.execute("PRAGMA foreign_keys=ON;")
    except Exception:
        pass
    return conn


def create_backup():
    """데이터 무소실을 위한 스냅샷 백업 생성 (로컬 SQLite 모드 전용)"""
    if is_using_turso() or not DB_PATH.exists():
        return None
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = BACKUP_DIR / f"overtime_backup_{timestamp}.db"
        src = sqlite3.connect(str(DB_PATH), timeout=15.0)
        dst = sqlite3.connect(str(backup_file))
        with dst:
            src.backup(dst)
        src.close()
        dst.close()
        
        backups = sorted(BACKUP_DIR.glob("overtime_backup_*.db"))
        if len(backups) > 30:
            for old_backup in backups[:-30]:
                try:
                    old_backup.unlink()
                except Exception:
                    pass
        return str(backup_file)
    except Exception as e:
        print(f"[Backup Error] {e}")
        return None


def init_db():
    """데이터베이스 및 테이블 초기화, 슈퍼관리자 시딩"""
    if not is_using_turso():
        create_backup()
    conn = get_db_connection()
    if not is_using_turso():
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA synchronous=NORMAL;")
        except Exception:
            pass
    cursor = conn.cursor()

    # 1. 회원(사용자/관리자) 테이블
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        emp_id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        team TEXT NOT NULL,
        position TEXT DEFAULT '팀원',
        is_admin INTEGER DEFAULT 0,
        is_super INTEGER DEFAULT 0,
        created_at TEXT NOT NULL
    )
    """)

    # 2. 특근 신청 내역 테이블
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS overtimes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_id TEXT NOT NULL,
        user_name TEXT NOT NULL,
        team TEXT NOT NULL,
        category TEXT NOT NULL,
        start_date TEXT NOT NULL,
        end_date TEXT NOT NULL,
        project_no TEXT,
        location TEXT,
        reason TEXT,
        is_confirmed INTEGER DEFAULT 0,
        confirmed_by TEXT,
        confirmed_at TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        bonus_granted INTEGER DEFAULT 0,
        FOREIGN KEY (emp_id) REFERENCES users (emp_id) ON UPDATE CASCADE
    )
    """)

    # 3. 수정/삭제/승인 감사 이력 테이블
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS overtime_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        overtime_id INTEGER,
        action TEXT NOT NULL,
        changed_by TEXT NOT NULL,
        changed_by_name TEXT,
        previous_data TEXT,
        new_data TEXT,
        created_at TEXT NOT NULL
    )
    """)

    # 4. 보안 감사 및 접속 로그 테이블
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS access_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_id TEXT NOT NULL,
        user_name TEXT,
        action_type TEXT NOT NULL,
        status TEXT NOT NULL,
        ip_address TEXT,
        user_agent TEXT,
        details TEXT,
        created_at TEXT NOT NULL
    )
    """)

    # 인덱스 생성
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_overtimes_emp ON overtimes(emp_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_overtimes_date ON overtimes(start_date, end_date);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_history_overtime ON overtime_history(overtime_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_access_logs_emp ON access_logs(emp_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_access_logs_created ON access_logs(created_at);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_access_logs_status ON access_logs(status);")

    # 대체휴일 사용 및 보너스 부여 컬럼 마이그레이션
    cursor.execute("PRAGMA table_info(overtimes);")
    columns = [row["name"] for row in cursor.fetchall()]
    if "sub_holiday_used" not in columns:
        cursor.execute("ALTER TABLE overtimes ADD COLUMN sub_holiday_used REAL DEFAULT 0;")
    if "sub_holiday_date" not in columns:
        cursor.execute("ALTER TABLE overtimes ADD COLUMN sub_holiday_date TEXT DEFAULT '';")
    if "bonus_granted" not in columns:
        cursor.execute("ALTER TABLE overtimes ADD COLUMN bonus_granted INTEGER DEFAULT 0;")

    # 4. 소속팀(부서) 관리 테이블
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS teams (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        created_at TEXT NOT NULL
    )
    """)

    # 기본 소속팀 시딩: DB에 소속팀이 하나도 없는 최초 1회 생성 시에만 시딩
    cursor.execute("SELECT COUNT(*) as cnt FROM teams")
    teams_cnt = cursor.fetchone()["cnt"]
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if teams_cnt == 0:
        default_teams = ['제어실', '전장배선팀', '전장설계팀', 'PLC제어팀']
        cursor.execute("SELECT DISTINCT team FROM users WHERE team IS NOT NULL AND team != ''")
        for r in cursor.fetchall():
            if r["team"] not in default_teams:
                default_teams.append(r["team"])
        for t_name in default_teams:
            cursor.execute("INSERT OR IGNORE INTO teams (name, created_at) VALUES (?, ?)", (t_name, now_str))

    # 5. 슈퍼관리자 (ps37082) 등록 보장
    cursor.execute("SELECT * FROM users WHERE emp_id = 'ps37082'")
    super_admin = cursor.fetchone()
    if not super_admin:
        cursor.execute("""
        INSERT INTO users (emp_id, name, team, position, is_admin, is_super, created_at)
        VALUES ('ps37082', '슈퍼관리자', '제어실', '총괄관리자', 1, 1, ?)
        """, (now_str,))
    else:
        cursor.execute("UPDATE users SET is_admin = 1, is_super = 1 WHERE emp_id = 'ps37082'")

    conn.commit()
    conn.close()

# -------------------------------------------------------------
# ⚡ 2. 소속팀 및 권한 인메모리 스마트 캐시 엔진 (반복 쿼리 제거)
# -------------------------------------------------------------
_TEAMS_CACHE = {"data": None, "expires": 0}
_USER_ROLE_CACHE = {} # emp_id -> (role_dict, expire_timestamp)
_CACHE_LOCK = threading.Lock()

def invalidate_teams_cache():
    with _CACHE_LOCK:
        _TEAMS_CACHE["data"] = None
        _TEAMS_CACHE["expires"] = 0

def invalidate_user_role_cache(emp_id: str = None):
    with _CACHE_LOCK:
        if emp_id:
            _USER_ROLE_CACHE.pop(emp_id.strip(), None)
        else:
            _USER_ROLE_CACHE.clear()

def get_cached_user_role(emp_id: str):
    if not emp_id:
        return None
    now = time.time()
    with _CACHE_LOCK:
        cached = _USER_ROLE_CACHE.get(emp_id.strip())
        if cached:
            role, exp = cached
            if now < exp:
                return role
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT is_super, is_admin, team, name, position FROM users WHERE emp_id = ?", (emp_id.strip(),))
        row = cursor.fetchone()
        if row:
            role_dict = dict(row)
            with _CACHE_LOCK:
                _USER_ROLE_CACHE[emp_id.strip()] = (role_dict, now + 45.0)
            return role_dict
        return None
    finally:
        conn.close()

def get_all_teams() -> list:
    """등록된 소속팀 목록 조회 (45초 인메모리 스마트 캐시 적용으로 초고속 반환)"""
    now = time.time()
    with _CACHE_LOCK:
        if _TEAMS_CACHE["data"] is not None and now < _TEAMS_CACHE["expires"]:
            return _TEAMS_CACHE["data"]

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, created_at FROM teams ORDER BY name ASC")
        data = [dict(row) for row in cursor.fetchall()]
        with _CACHE_LOCK:
            _TEAMS_CACHE["data"] = data
            _TEAMS_CACHE["expires"] = now + 45.0
        return data
    finally:
        conn.close()

def create_team(name: str) -> dict:
    """새로운 소속팀 추가 (슈퍼관리자 전용)"""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("INSERT INTO teams (name, created_at) VALUES (?, ?)", (name, now_str))
        team_id = cursor.lastrowid
        conn.commit()
        invalidate_teams_cache()
        return {"id": team_id, "name": name, "created_at": now_str}
    finally:
        conn.close()

def delete_team(name: str) -> bool:
    """소속팀 삭제 (소속 팀원이 없는 경우만)"""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) as cnt FROM users WHERE team = ?", (name,))
        if cursor.fetchone()["cnt"] > 0:
            raise ValueError("해당 팀에 등록된 팀원이 존재하여 삭제할 수 없습니다.")
        cursor.execute("DELETE FROM teams WHERE name = ?", (name,))
        conn.commit()
        invalidate_teams_cache()
        return True
    finally:
        conn.close()

# -------------------------------------------------------------
# ⚡ 3. 비동기 백그라운드 감사 로깅 엔진 (응답 지연 0초화)
# -------------------------------------------------------------
_LOG_QUEUE = Queue()

def _async_log_worker():
    """백그라운드에서 접속 로그 및 감사 로그를 순차적으로 처리하여 사용자 응답을 가로막지 않음"""
    while True:
        try:
            task_type, payload = _LOG_QUEUE.get()
            if task_type == "AUDIT":
                _sync_log_audit(*payload)
            elif task_type == "ACCESS":
                _sync_log_access(*payload)
        except Exception as e:
            print(f"[Async Log Worker Error] {e}")
        finally:
            _LOG_QUEUE.task_done()

_log_thread = threading.Thread(target=_async_log_worker, daemon=True, name="AuditLogWorker")
_log_thread.start()

def _sync_log_audit(overtime_id, action, changed_by, changed_by_name, prev_json, new_json, now_str):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO overtime_history (overtime_id, action, changed_by, changed_by_name, previous_data, new_data, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (overtime_id, action, changed_by, changed_by_name, prev_json, new_json, now_str))
        conn.commit()
    finally:
        conn.close()

def log_audit(overtime_id: int, action: str, changed_by: str, changed_by_name: str, previous_data: dict = None, new_data: dict = None):
    """특근 생성/수정/삭제/확인 감사 이력 비동기 기록"""
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prev_json = json.dumps(previous_data, ensure_ascii=False) if previous_data else None
    new_json = json.dumps(new_data, ensure_ascii=False) if new_data else None
    _LOG_QUEUE.put(("AUDIT", (overtime_id, action, changed_by, changed_by_name, prev_json, new_json, now_str)))

def _sync_log_access(emp_id, user_name, action_type, status, ip_address, user_agent, details, now_str):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO access_logs (emp_id, user_name, action_type, status, ip_address, user_agent, details, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (emp_id, user_name, action_type, status, ip_address, user_agent, details, now_str))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[Access Log Error] {e}")

def log_access_event(emp_id: str, user_name: str = None, action_type: str = "LOGIN", status: str = "SUCCESS", ip_address: str = "", user_agent: str = "", details: str = ""):
    """사용자 로그인 시도, 등록, 접속 감사 로그 비동기 기록"""
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _LOG_QUEUE.put(("ACCESS", (emp_id or "", user_name or "", action_type, status, ip_address or "", user_agent or "", details or "", now_str)))
