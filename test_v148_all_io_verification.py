import sys
import os
import filecmp

sys.stdout.reconfigure(encoding='utf-8')

print("=== Running v1.48 All I/O & Alert Verification ===")

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

io_targets = [
    ("특근 엑셀 내보내기", "exportExcelBtn", "⏳ 특근 엑셀 생성 중..."),
    ("특근 엑셀 가져오기", "handleOvertimeExcelFileSelected", "⏳ 엑셀 읽는 중..."),
    ("팀원 명부 내보내기", "exportUsersBtn", "⏳ 팀원명부 엑셀 생성 중..."),
    ("팀원 명부 가져오기", "handleUserExcelFileSelected", "⏳ 엑셀 읽는 중..."),
    ("실특근 정산표 내보내기", "exportSummaryExcelBtn", "⏳ 정산표 엑셀 생성 중..."),
    ("실특근 정산표 가져오기", "importSummaryExcelBtn", "⏳ 정산표 읽는 중..."),
    ("웹 스냅샷 백업 저장", "saveWebBackup", "⏳ 스냅샷 저장 중..."),
    ("웹 스냅샷 복원(열기)", "btn-restore-backup", "⏳ 스냅샷 복원(열기) 중..."),
    ("보안 감사 로그 내보내기", "exportAccessLogsBtn", "⏳ 감사로그 엑셀 생성 중...")
]

passed = 0
for title, key_sym, loading_text in io_targets:
    if key_sym in app_code and loading_text in app_code:
        print(f"  [PASS] {title} ({key_sym}): Loading state & Alert configured")
        passed += 1
    else:
        print(f"  [FAIL] {title} ({key_sym}): Missing key elements")

print(f"\nTotal I/O handlers verified: {passed}/{len(io_targets)}")

# Check mirror synchronization
mirrors = ["js/app.js", "static/app.js", "static/js/app.js"]
sync_ok = True
for m in mirrors:
    if filecmp.cmp("app.js", m, shallow=False):
        print(f"  [PASS] Mirror file {m} is synchronized")
    else:
        print(f"  [FAIL] Mirror file {m} is OUT OF SYNC")
        sync_ok = False

if passed == len(io_targets) and sync_ok:
    print("\n✅ ALL VERIFICATION CHECKS PASSED SUCCESSFULLY!")
    sys.exit(0)
else:
    print("\n❌ VERIFICATION FAILED!")
    sys.exit(1)
