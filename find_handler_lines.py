import re

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

keywords = [
    'exportExcelBtn',
    'exportUsersBtn',
    'importUsersBtn',
    'handleUserExcelFileSelected',
    'exportSummaryExcelBtn',
    'importSummaryExcelBtn',
    'saveWebBackup',
    'btnExecuteWebBackupSave',
    'loadBackupList',
    'restore',
    'exportAccessLogsBtn'
]

lines = app_code.split('\n')
for idx, line in enumerate(lines, 1):
    for kw in keywords:
        if kw in line:
            print(f"Line {idx}: [{kw}] -> {line[:120]}")
