import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

print("App.js total lines:", len(app_code.split('\n')))

# Check if triggerUserExcelImport is defined
if "triggerUserExcelImport" in app_code:
    print("triggerUserExcelImport found in app.js")
else:
    print("triggerUserExcelImport NOT found in app.js")

if "triggerOvertimeExcelImport" in app_code:
    print("triggerOvertimeExcelImport found in app.js")
else:
    print("triggerOvertimeExcelImport NOT found in app.js")
