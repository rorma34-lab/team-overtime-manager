import os
import sys
import json
import sqlite3
import asyncio
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from main import app, toggle_overtime_pre_deduct, get_user_overtime_stats
from schemas import OvertimePreDeductRequest
from database import get_db_connection, init_db
from exporter import aggregate_user_holidays, generate_overtime_excel

def test_v142_all():
    print("==================================================")
    print("🚀 TEAM OVERTIME MANAGER v1.42 VERIFICATION TEST")
    print("==================================================")

    # 1. Version check
    assert app.version == "v1.42", f"App version mismatch: {app.version}"
    print("✅ 1. FastAPI App Version is v1.42")

    # 2. Database & Sample test data setup
    init_db()
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM overtimes WHERE emp_id = 'test9999'")
        cursor.execute("DELETE FROM users WHERE emp_id = 'test9999'")
        cursor.execute("""
            INSERT INTO users (emp_id, name, team, position, is_admin, is_super, created_at)
            VALUES ('test9999', '테스트맨', '기술연구소', '책임', 0, 0, '2026-09-01 09:00:00')
        """)

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Case A: 일반휴일 2일 (2026-09-05 ~ 2026-09-06), 사전차감 1 (출장기간 외)
        cursor.execute("""
            INSERT INTO overtimes (emp_id, user_name, team, category, start_date, end_date, project_no, location, reason,
                                  sub_holiday_date, sub_holiday_used, is_pre_deduct, trip_start_date, trip_end_date, is_confirmed, bonus_granted, created_at, updated_at)
            VALUES ('test9999', '테스트맨', '기술연구소', '일반휴일', '2026-09-05', '2026-09-06', 'PRJ-01', '본사', '연구개발',
                    '', 0.0, 1, '', '', 1, 0, ?, ?)
        """, (now_str, now_str))
        ot_id_1 = cursor.lastrowid

        # Case B: 대체휴무 1일 (2026-09-15), 출장기간 2026-09-10 ~ 2026-09-20
        cursor.execute("""
            INSERT INTO overtimes (emp_id, user_name, team, category, start_date, end_date, project_no, location, reason,
                                  sub_holiday_date, sub_holiday_used, is_pre_deduct, trip_start_date, trip_end_date, is_confirmed, bonus_granted, created_at, updated_at)
            VALUES ('test9999', '테스트맨', '기술연구소', '대체휴무', '2026-09-15', '2026-09-15', 'PRJ-01', '본사', '출장휴무',
                    '', 0.0, 0, '2026-09-10', '2026-09-20', 1, 0, ?, ?)
        """, (now_str, now_str))

        # Case C: 일반휴일 1일 (2026-09-12), 출장기간 내 사전차감 1
        cursor.execute("""
            INSERT INTO overtimes (emp_id, user_name, team, category, start_date, end_date, project_no, location, reason,
                                  sub_holiday_date, sub_holiday_used, is_pre_deduct, trip_start_date, trip_end_date, is_confirmed, bonus_granted, created_at, updated_at)
            VALUES ('test9999', '테스트맨', '기술연구소', '일반휴일', '2026-09-12', '2026-09-12', 'PRJ-01', '현장', '출장 중 특근',
                    '', 0.0, 1, '', '', 1, 0, ?, ?)
        """, (now_str, now_str))
        ot_id_3 = cursor.lastrowid

        conn.commit()

    print(f"✅ 2. Test data inserted: ot_1={ot_id_1}, ot_3={ot_id_3}")

    # 3. Test Pre-deduct toggle function
    req0 = OvertimePreDeductRequest(admin_emp_id="ps37082", is_pre_deduct=0)
    res0 = toggle_overtime_pre_deduct(ot_id_1, req0)
    assert res0["is_pre_deduct"] == 0, f"Expected 0, got {res0}"

    req1 = OvertimePreDeductRequest(admin_emp_id="ps37082", is_pre_deduct=1)
    res1 = toggle_overtime_pre_deduct(ot_id_1, req1)
    assert res1["is_pre_deduct"] == 1, f"Expected 1, got {res1}"
    print("✅ 3. toggle_overtime_pre_deduct endpoint logic verified successfully")

    # 4. Test get_user_overtime_stats
    stats_data = get_user_overtime_stats("test9999")
    assert "recent_months" in stats_data, "Missing recent_months in stats"
    assert "quarters" in stats_data, "Missing quarters in stats"
    assert "halves" in stats_data, "Missing halves in stats"
    assert "by_year" in stats_data, "Missing by_year in stats"

    # Find 9월 in recent_months dict
    sep_month = stats_data["recent_months"].get("9월")
    assert sep_month is not None, f"9월 not found in recent_months: {stats_data['recent_months'].keys()}"
    print(f"📊 9월 Bucket: {sep_month}")
    assert sep_month["일반휴일"] == 3, f"Expected 3 normal days, got {sep_month['일반휴일']}"
    assert sep_month["대체휴무"] == 1, f"Expected 1 sub rest day, got {sep_month['대체휴무']}"
    assert sep_month["최종실특근"] == 1.0, f"Expected 최종실특근 1.0, got {sep_month['최종실특근']}"
    assert sep_month["사전차감"] == 2, f"Expected 2 total pre-deduct, got {sep_month['사전차감']}"
    assert sep_month["출장내사전차감"] == 1, f"Expected 1 trip pre-deduct, got {sep_month['출장내사전차감']}"
    assert sep_month["사전차감잔여"] == 1, f"Expected 1 pre-deduct remaining, got {sep_month['사전차감잔여']}"
    print("✅ 4. User stats buckets & mathematical formula verified")

    # 5. Test Exporter logic & headers
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM overtimes WHERE emp_id = 'test9999'")
        records = [dict(r) for r in cursor.fetchall()]
        user_sums = aggregate_user_holidays(records, {"test9999": "책임"})
        u = user_sums[0]
        assert u["normal_holiday_days"] == 3
        assert u["sub_holiday_days"] == 1
        assert u["actual_overtime_days"] == 1.0
        assert u["pre_deduct_count"] == 2
        assert u["trip_pre_deduct_count"] == 1
        assert u["pre_deduct_remaining"] == 1

        excel_bytes = generate_overtime_excel(records, {"test9999": "책임"})
        assert len(excel_bytes.getvalue()) > 1000, "Generated Excel is too small"

    print("✅ 5. Exporter aggregate_user_holidays & generate_overtime_excel verified")

    # Cleanup test data
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM overtimes WHERE emp_id = 'test9999'")
        cursor.execute("DELETE FROM users WHERE emp_id = 'test9999'")
        conn.commit()

    print("\n🎉 ALL v1.42 BACKEND & BUSINESS LOGIC TESTS PASSED PERFECTLY!\n")

if __name__ == "__main__":
    test_v142_all()
