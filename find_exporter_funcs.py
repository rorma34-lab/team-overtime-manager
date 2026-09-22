import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

lines = code.split('\n')
for idx, line in enumerate(lines, 1):
    if 'def ' in line:
        print(f"Line {idx}: {line[:120]}")
