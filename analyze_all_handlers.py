import re

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

with open('index.html', 'r', encoding='utf-8') as f:
    html_code = f.read()

print("=== Functions in app.js related to import/export/save/load/backup/restore/excel ===")
funcs = re.findall(r'(?:async\s+)?function\s+([a-zA-Z0-9_]+)\s*\(', app_code)
for fn in sorted(set(funcs)):
    if any(k in fn.lower() for k in ['export', 'import', 'save', 'backup', 'restore', 'load', 'excel', 'snapshot', 'upload', 'download']):
        print(f" - {fn}")

print("\n=== Click / Change event listeners in app.js ===")
listeners = re.findall(r'document\.getElementById\([\'"]([^\'"]+)[\'"]\)\s*\.\s*addEventListener\([\'"]([^\'"]+)[\'"]\s*,\s*([^;\)]+)', app_code)
for el_id, evt, fn in listeners:
    if any(k in el_id.lower() or k in fn.lower() for k in ['export', 'import', 'save', 'backup', 'restore', 'load', 'excel', 'snapshot', 'upload', 'download', 'btn']):
        print(f" - #{el_id} ({evt}) => {fn.strip()}")

print("\n=== onclick / onchange handlers in index.html ===")
html_handlers = re.findall(r'(onclick|onchange|id)=["\']([^"\']+)["\']', html_code)
for attr, val in html_handlers:
    if any(k in val.lower() for k in ['export', 'import', 'save', 'backup', 'restore', 'load', 'excel', 'snapshot', 'upload', 'download']):
        print(f" - {attr}=\"{val}\"")
