import io
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

from database import get_db_connection, init_db
from exporter import generate_overtime_excel, import_overtimes_from_excel

init_db()

conn = get_db_connection()
cursor = conn.cursor()

# Get sample overtimes from DB
cursor.execute("SELECT * FROM overtimes LIMIT 5")
rows = [dict(r) for r in cursor.fetchall()]

print(f"Fetched {len(rows)} overtimes for export test")

if rows:
    # Generate Excel bytes
    excel_bio = generate_overtime_excel(rows)
    excel_bytes = excel_bio.getvalue()
    print(f"Generated Excel size: {len(excel_bytes)} bytes")

    # Save to disk for inspection
    with open("test_import_sample.xlsx", "wb") as f:
        f.write(excel_bytes)

    # Test importing back
    res = import_overtimes_from_excel(excel_bytes, target_teams=None, admin_emp_id="ADMIN")
    print("\n=== Excel Import Result ===")
    print(res)
else:
    print("No overtime records in DB to test")

conn.close()
