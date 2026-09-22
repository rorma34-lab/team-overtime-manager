import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('app.js', 'r', encoding='utf-8') as f:
    code = f.read()

lines = code.split('\n')
for idx, line in enumerate(lines, 1):
    if 'import-excel' in line or 'importovertime' in line.lower() or 'excelimport' in line.lower():
        print(f"Line {idx}: {line[:120]}")
