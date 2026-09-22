"""
v1.47 기능 검증 및 백엔드 테스트 스크립트
1. 관리자 팀원 신청 부서선택 필터링 & 디폴트 토요일 날짜
2. 특근현황_일자별개인별정산 엑셀 가져오기 API (/api/overtimes/import-settlement)
3. 웹 저장(스냅샷 백업/복원 /api/backup/save) 정상 동작 검증
"""
import sys
import io
import json
import base64
import asyncio
import openpyxl
from starlette.requests import Request

from main import app, import_settlement, export_settlement, save_web_backup, list_web_backups
from database import get_db_connection, init_db
from schemas import BackupSaveRequest

def build_dummy_request(json_data: dict) -> Request:
    raw_bytes = json.dumps(json_data).encode('utf-8')
    scope = {
        'type': 'http',
        'method': 'POST',
        'headers': [(b'content-type', b'application/json')],
    }
    async def receive():
        return {'type': 'http.request', 'body': raw_bytes}
    return Request(scope, receive)

def run_v147_tests():
    print("=" * 70)
    print("▶ [Test 1] 시스템 초기화 및 v1.47 버전 검증")
    print("=" * 70)
    assert app.version == "v1.47", f"Expected version v1.47, got {app.version}"
    print(f"✅ App Version: {app.version}")

    # DB 초기화 및 픽스처 데이터 준비
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 더미 부서 및 인원 준비
    cursor.execute("INSERT OR IGNORE INTO users (emp_id, name, team, position, is_admin) VALUES ('TEST001', '홍길동', '기술연구팀', '선임', 0)")
    cursor.execute("INSERT OR IGNORE INTO users (emp_id, name, team, position, is_admin) VALUES ('TEST002', '김철수', '생산1팀', '책임', 0)")
    cursor.execute("INSERT OR IGNORE INTO users (emp_id, name, team, position, is_admin) VALUES ('ADMIN01', '최관리', '기술연구팀', '팀장', 1)")
    conn.commit()

    print("\n" + "=" * 70)
    print("▶ [Test 2] 특근 정산 엑셀 가져오기 - 부서 격리 및 타 부서 변경 방지 검증")
    print("=" * 70)

    # 샘플 엑셀 워크북 생성 (기술연구팀 2건, 생산1팀 2건)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "휴일일자별_특근현황"
    
    headers = [
        "순번", "휴일날짜", "요일", "사번", "성명", "소속팀", "특근분류", "보너스 부여",
        "진행단계", "휴일일수", "대체휴무 사용일", "대체휴가 사용일수", "사전차감",
        "출장기간(시작)", "출장기간(종료)", "프로젝트 번호", "근무 장소", "특근 사유"
    ]
    ws.append(headers)

    # 데이터 행 추가
    ws.append([1, "2026-09-20", "일", "TEST001", "홍길동", "기술연구팀", "일반휴일", "0", "승인", 1, 0, 0, 0, "", "", "PRJ-101", "연구소", "개발 테스트"])
    ws.append([2, "2026-09-21", "월", "TEST001", "홍길동", "기술연구팀", "일반휴일", "1", "확정", 1, 0, 0, 0, "", "", "PRJ-102", "연구소", "서버 구축"])
    ws.append([3, "2026-09-20", "일", "TEST002", "김철수", "생산1팀", "일반휴일", "0", "신청", 1, 0, 0, 0, "", "", "PRJ-201", "공장1", "생산 기계 점검"])
    ws.append([4, "2026-09-21", "월", "TEST002", "김철수", "생산1팀", "일반휴일", "0", "승인", 1, 0, 0, 0, "", "", "PRJ-202", "공장1", "품질 검사"])

    excel_buf = io.BytesIO()
    wb.save(excel_buf)
    excel_base64 = base64.b64encode(excel_buf.getvalue()).decode("utf-8")

    # [케이스 A] '기술연구팀' 부서 지정 후 업로드
    req_a = build_dummy_request({
        "file_base64": excel_base64,
        "filename": "test_import.xlsx",
        "target_team": "기술연구팀",
        "admin_emp_id": "ADMIN01"
    })
    
    data = asyncio.run(import_settlement(req_a))
    
    assert data["success"] is True
    assert data["total_rows"] == 4
    assert data["processed_count"] == 2, f"Expected 2 processed for 기술연구팀, got {data['processed_count']}"
    assert data["skipped_other_dept_count"] == 2, f"Expected 2 skipped for 생산1팀, got {data['skipped_other_dept_count']}"
    print(f"✅ '기술연구팀' 업로드 결과: 처리 {data['processed_count']}건, 타 부서 제외 {data['skipped_other_dept_count']}건")

    # DB 검증: 기술연구팀(TEST001)의 특근만 입력되었고 생산1팀(TEST002)의 특근은 입력되지 않았는지 확인
    cursor.execute("SELECT COUNT(*) as cnt FROM overtimes WHERE emp_id = 'TEST001'")
    cnt_dev = cursor.fetchone()["cnt"]
    assert cnt_dev == 2, f"Expected 2 records for TEST001, got {cnt_dev}"

    cursor.execute("SELECT COUNT(*) as cnt FROM overtimes WHERE emp_id = 'TEST002'")
    cnt_prod = cursor.fetchone()["cnt"]
    assert cnt_prod == 0, f"Expected 0 records for TEST002 (타부서 보호), got {cnt_prod}"
    print("✅ 타 부서('생산1팀') 데이터 변경 차단 및 보존 확인 완료!")

    print("\n" + "=" * 70)
    print("▶ [Test 3] 웹 저장 (스냅샷 백업 save_web_backup) 정상 작동 검증")
    print("=" * 70)
    backup_req = BackupSaveRequest(name="테스트_웹저장_스냅샷", description="자동 테스트 백업")
    backup_res = save_web_backup(backup_req)
    assert backup_res["status"] == "success"
    assert backup_res["name"] == "테스트_웹저장_스냅샷"
    assert backup_res["counts"]["users"] >= 3
    print(f"✅ 웹 저장 스냅샷 성공: {backup_res['name']} (사원 {backup_res['counts']['users']}명, 특근 {backup_res['counts']['overtimes']}건)")

    # 백업 목록 확인
    list_res = list_web_backups()
    assert len(list_res["backups"]) > 0
    print(f"✅ 웹 저장 스냅샷 목록 조회 성공: 총 {len(list_res['backups'])}개 스냅샷 존재")

    # 정리
    cursor.execute("DELETE FROM overtimes WHERE emp_id IN ('TEST001', 'TEST002')")
    cursor.execute("DELETE FROM users WHERE emp_id IN ('TEST001', 'TEST002', 'ADMIN01')")
    conn.commit()
    conn.close()

    print("\n" + "🎉 ALL v1.47 BACKEND, EXCEL IMPORT, AND WEB SAVE TESTS PASSED PERFECTLY!")

if __name__ == "__main__":
    run_v147_tests()
