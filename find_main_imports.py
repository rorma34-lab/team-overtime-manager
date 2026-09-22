import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('main.py', 'r', encoding='utf-8') as f:
    code = f.read()

lines = code.split('\n')
for idx, line in enumerate(lines, 1):
    if 'import' in line.lower() or 'excel' in line.lower() or 'overtime' in line.lower():
        if 'api' in line.lower() or 'def ' in line.lower() or '@app' in line.lower():
            print(f"Line {idx}: {line[:120]}")
