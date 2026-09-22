import re

with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

buttons = re.findall(r'<button[^>]*id=["\']([^"\']+)["\'][^>]*>', html)
print("=== All Buttons with IDs in index.html ===")
for b in sorted(set(buttons)):
    print(f" - #{b}")
