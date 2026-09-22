import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

operations = [
    ("1. 特勤 Excel Export", "exportExcelBtn"),
    ("2. 特勤 Excel Import", "handleOvertimeExcelFileSelected"),
    ("3. 팀원 Excel Export", "exportUsersBtn"),
    ("4. 팀원 Excel Import", "handleUserExcelFileSelected"),
    ("5. 실특근정산 Excel Export", "exportSummaryExcelBtn"),
    ("6. 실특근정산 Excel Import", "importSummaryExcelBtn"),
    ("7. 웹 스냅샷 백업 저장", "saveWebBackup"),
    ("8. 웹 스냅샷 백업 목록/복원(열기)", "loadBackupList"),
    ("9. 보안 감사 로그 Export", "exportAccessLogsBtn")
]

for title, kw in operations:
    print(f"\n==================== {title} (Keyword: {kw}) ====================")
    lines = app_code.split('\n')
    found = False
    for idx, line in enumerate(lines, 1):
        if kw in line:
            found = True
            start = max(1, idx - 2)
            end = min(len(lines), idx + 25)
            snippet = "\n".join([f"{i}: {lines[i-1]}" for i in range(start, end)])
            print(snippet)
            print("--------------------------------------------------")
            break
    if not found:
        print("NOT FOUND")
