import io
import sys
import os
import openpyxl
import database
import main
import exporter

def run_tests():
    print("=" * 70)
    print("▶ [Test 1] 시스템 초기화 및 v1.48 버전 검증")
    print("=" * 70)
    database.init_db()
    version = main.app.version
    print(f"✅ App Version: {version}")
    assert version == "v1.48", f"Expected v1.48, got {version}"

    print("\n" + "=" * 70)
    print("▶ [Test 2] 코드 백업 스냅샷 폴더 존재 검증")
    print("=" * 70)
    backup_dir = os.path.join("data", "code_backups")
    assert os.path.exists(backup_dir), "data/code_backups directory should exist"
    snapshots = [d for d in os.listdir(backup_dir) if os.path.isdir(os.path.join(backup_dir, d))]
    print(f"✅ 발견된 소스 백업 스냅샷 수: {len(snapshots)}개 ({snapshots[-1] if snapshots else 'None'})")
    assert len(snapshots) >= 1, "At least one code backup snapshot should exist"

    print("\n" + "=" * 70)
    print("▶ [Test 3] 특근 엑셀 가져오기 (첫 번째 시트) & 복수 부서 격리 & 중복 필터 검증")
    print("=" * 70)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "휴일일자별_특근현황"
    ws.append(["순번", "신청자", "소속팀", "분류", "근무기간", "프로젝트", "장소", "사유", "보너스", "사전차감"])
    
    # 1행: PLC제어팀 (테스트맨 - 신규)
    ws.append([1, "테스트맨(123456)", "PLC제어팀", "일반휴일", "2026-10-01 ~ 2026-10-01", "P-101", "공장", "설비점검", "예", "아니오"])
    # 2행: 제어실 (정진규 - 갱신)
    ws.append([2, "정진규(113019)", "제어실", "일반휴일", "2026-09-01 ~ 2026-09-01", "P-102", "연구소", "업무 대응", "아니오", "아니오"])
    # 3행: 타부서 (생산1팀 - 대상 부서 제외)
    ws.append([3, "생산원(999999)", "생산1팀", "일반휴일", "2026-10-05 ~ 2026-10-05", "P-103", "라인", "생산지원", "아니오", "아니오"])
    # 4행: 오류 데이터 (사번 없음)
    ws.append([4, "", "PLC제어팀", "일반휴일", "2026-10-10 ~ 2026-10-10", "P-104", "라인", "오류건", "아니오", "아니오"])

    bio = io.BytesIO()
    wb.save(bio)
    file_bytes = bio.getvalue()

    # 복수 부서 target_teams = ["PLC제어팀", "제어실"] 로 가져오기 테스트
    res = exporter.import_overtimes_from_excel(file_bytes, target_teams=["PLC제어팀", "제어실"], admin_emp_id="ps37082")
    print(f"✅ 가져오기 결과: 총 {res['total_rows']}행 읽음 | 처리 {res['processed_count']}건 | 신규 {res['created_count']}건 | 갱신 {res['updated_count']}건 | 타부서 제외 {res['ignored_teams_count']}건 | 오류 {len(res['errors'])}건")

    assert res['success'] == True
    assert res['processed_count'] == 2, f"Expected 2 processed, got {res['processed_count']}"
    assert res['ignored_teams_count'] == 1, f"Expected 1 ignored team, got {res['ignored_teams_count']}"
    assert len(res['errors']) == 1, f"Expected 1 error row, got {len(res['errors'])}"
    print(f"✅ 오류 상세 항목: {res['errors'][0]}")

    print("\n" + "=" * 70)
    print("🎉 ALL v1.48 BACKEND & EXCEL IMPORT TESTS PASSED PERFECTLY!")
    print("=" * 70)

if __name__ == "__main__":
    run_tests()
