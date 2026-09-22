import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

lines = app_code.split('\n')
for idx, line in enumerate(lines, 1):
    if 'importSummaryExcelBtn' in line:
        print(f"Line {idx}: {line[:120]}")
