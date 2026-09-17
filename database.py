import os
import sqlite3
import shutil
import json
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
BACKUP_DIR = DATA_DIR / "backups"
DB_PATH = DATA_DIR / "overtime.db"

# 디렉토리 생성
DATA_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

def get_db_connection():
    """SQLite DB 연결 객체 반환 (Row 팩토리 적용, 타임아웃 및 외래키 설정)"""
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False, timeout=15.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=15000;")
    conn.execute("PRAGMA foreign_keys=ON;")
    return conn


def create_backup():
    """데이터 무소실을 위한 스냅샷 백업 생성"""
    if not DB_PATH.exists():
        return None
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = BACKUP_DIR / f"overtime_backup_{timestamp}.db"
        # SQLite 온라인 백업 API 사용 (가장 안전한 방법)
        src = sqlite3.connect(str(DB_PATH), timeout=15.0)
        dst = sqlite3.connect(str(backup_file))
        with dst:
            src.backup(dst)
        src.close()
        dst.close()
        
        # 최근 30개 백업 유지, 오래된 백업 정리
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
    create_backup()
    conn = get_db_connection()
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
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

    # 4. 보안 감사 및 접속 로그 테이블 (신규 요구사항 5)
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

    # 대체휴일 사용 및 보너스 부여 컬럼 마이그레이션 (기존 DB 안전 업그레이드)
    cursor.execute("PRAGMA table_info(overtimes);")
    columns = [row["name"] for row in cursor.fetchall()]
    if "sub_holiday_used" not in columns:
        cursor.execute("ALTER TABLE overtimes ADD COLUMN sub_holiday_used REAL DEFAULT 0;")
    if "sub_holiday_date" not in columns:
        cursor.execute("ALTER TABLE overtimes ADD COLUMN sub_holiday_date TEXT DEFAULT '';")
    if "bonus_granted" not in columns:
        cursor.execute("ALTER TABLE overtimes ADD COLUMN bonus_granted INTEGER DEFAULT 0;")

    # 4. 소속팀(부서) 관리 테이블 (요구사항 25)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS teams (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        created_at TEXT NOT NULL
    )
    """)

    # 기본 소속팀 시딩: DB에 소속팀이 하나도 없는 최초 1회 생성 시에만 시딩 (기 삭제된 팀 재부활 완전 방지)
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

def get_all_teams() -> list:
    """등록된 소속팀 목록 조회"""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, created_at FROM teams ORDER BY name ASC")
        return [dict(row) for row in cursor.fetchall()]
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
        return True
    finally:
        conn.close()

def log_audit(overtime_id: int, action: str, changed_by: str, changed_by_name: str, previous_data: dict = None, new_data: dict = None):
    """특근 생성/수정/삭제/확인 감사 이력 기록"""
    conn = get_db_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
    INSERT INTO overtime_history (overtime_id, action, changed_by, changed_by_name, previous_data, new_data, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        overtime_id,
        action,
        changed_by,
        changed_by_name,
        json.dumps(previous_data, ensure_ascii=False) if previous_data else None,
        json.dumps(new_data, ensure_ascii=False) if new_data else None,
        now_str
    ))
    conn.commit()
    conn.close()

def log_access_event(emp_id: str, user_name: str = None, action_type: str = "LOGIN", status: str = "SUCCESS", ip_address: str = "", user_agent: str = "", details: str = ""):
    """사용자 로그인 시도, 등록, 접속 감사 로그 기록 (신규 요구사항 5)"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("""
        INSERT INTO access_logs (emp_id, user_name, action_type, status, ip_address, user_agent, details, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            emp_id or "",
            user_name or "",
            action_type,
            status,
            ip_address or "",
            user_agent or "",
            details or "",
            now_str
        ))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[Access Log Error] {e}")
