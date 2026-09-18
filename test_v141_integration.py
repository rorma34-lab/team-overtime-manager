import urllib.request
import urllib.parse
import json

BASE = "http://localhost:8000"

def req(url, method="GET", data=None):
    if data:
        body = json.dumps(data).encode("utf-8")
        request = urllib.request.Request(url, data=body, method=method, headers={"Content-Type": "application/json"})
    else:
        request = urllib.request.Request(url, method=method)
    with urllib.request.urlopen(request) as res:
        res_data = res.read().decode("utf-8")
        return res.status, json.loads(res_data) if res_data else {}

def test_v141():
    print("=== [1] user_preferences API & Auto-save Test ===")
    emp_id = "test_user_v141"
    
    # 1. 회원 임시 등록
    try:
        req(f"{BASE}/api/users", method="POST", data={"emp_id": emp_id, "name": "테스트사원", "team": "제어실", "position": "선임"})
    except Exception as e:
        print("user create error/exists:", e)
        
    # 2. 특근 신청 (프로젝트번호와 근무장소 포함)
    ot_data = {
        "emp_id": emp_id,
        "category": "일반휴일",
        "start_date": "2026-09-20",
        "end_date": "2026-09-20",
        "project_no": "PRJ-2026-V141",
        "location": "대전테스트센터",
        "reason": "v1.41 기능 검증",
        "is_pre_deduct": 1
    }
    status, res = req(f"{BASE}/api/overtimes", method="POST", data=ot_data)
    assert status == 200, f"Overtime creation failed: {res}"
    ot_id = res["id"]
    print(f"✓ Overtime created with id: {ot_id}")

    # 3. user_preferences 확인
    status, prefs = req(f"{BASE}/api/users/{emp_id}/preferences")
    assert status == 200, f"Get preferences failed: {prefs}"
    print("✓ Preferences response:", prefs)
    assert prefs["exists"] is True
    assert prefs["preferences"]["last_project_no"] == "PRJ-2026-V141"
    assert prefs["preferences"]["last_location"] == "대전테스트센터"
    print("✓ User preferences auto-saved correctly!")

    # 4. 특근 수정 (사전차감 및 출장기간)
    print("\n=== [2] Overtime Update (Pre-deduct & Trip Dates) ===")
    update_data = {
        "changed_by": "ps37082",
        "category": "대체휴무",
        "start_date": "2026-09-20",
        "end_date": "2026-09-20",
        "project_no": "PRJ-2026-V141",
        "location": "대전테스트센터",
        "reason": "사전차감 및 대체휴무 수정",
        "is_pre_deduct": 1,
        "trip_start_date": "2026-09-15",
        "trip_end_date": "2026-09-18"
    }
    status, res = req(f"{BASE}/api/overtimes/{ot_id}", method="PUT", data=update_data)
    assert status == 200, f"Overtime update failed: {res}"
    
    # 5. 조회 확인
    status, res = req(f"{BASE}/api/overtimes/{ot_id}")
    assert status == 200
    item = res["overtime"]
    assert item["is_pre_deduct"] == 1
    assert item["trip_start_date"] == "2026-09-15"
    assert item["trip_end_date"] == "2026-09-18"
    assert item["category"] == "대체휴무"
    print("✓ Overtime updated with pre_deduct=1, trip_dates successfully!")

    # 6. 약식 통계 API 검증
    print("\n=== [3] User Stats Summary API ===")
    status, stats = req(f"{BASE}/api/users/{emp_id}/overtime-stats")
    assert status == 200, f"Stats API failed: {stats}"
    print("✓ User stats summary:", json.dumps(stats, ensure_ascii=False, indent=2))
    assert "periods" in stats
    assert "by_year" in stats
    print("✓ Overtime stats returned accurately!")

    # 7. 정리 (테스트 데이터 삭제)
    try:
        req(f"{BASE}/api/overtimes/{ot_id}", method="DELETE")
        req(f"{BASE}/api/users/{emp_id}", method="DELETE")
    except Exception as e:
        pass
    print("\n✓ Cleaned up test record.")
    print("\n🎉 ALL v1.41 INTEGRATION TESTS PASSED!")

if __name__ == "__main__":
    test_v141()
