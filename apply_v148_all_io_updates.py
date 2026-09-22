import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

# 1. Update exportExcelBtn
old_export_excel = """document.getElementById('exportExcelBtn').addEventListener('click', async () => {
  let targetList = [];
  if (selectedOvertimeIds.size > 0) {
    targetList = lastFetchedOvertimes.filter(o => selectedOvertimeIds.has(o.id));
  } else {
    targetList = [...lastFetchedOvertimes];
  }

  if (!targetList || targetList.length === 0) {
    showToast('내보낼 특근 내역이 없습니다.', 'error');
    return;
  }

  const exportBtn = document.getElementById('exportExcelBtn');
  const restoreExportBtn = setButtonLoading(exportBtn, '엑셀 생성 중...');
  showToast('엑셀 파일을 생성 중입니다...');"""

new_export_excel = """document.getElementById('exportExcelBtn').addEventListener('click', async () => {
  const exportBtn = document.getElementById('exportExcelBtn');
  const restoreExportBtn = setButtonLoading(exportBtn, '⏳ 특근 엑셀 생성 중...');
  showToast('엑셀 파일을 생성 중입니다...');

  let targetList = [];
  if (selectedOvertimeIds.size > 0) {
    targetList = lastFetchedOvertimes.filter(o => selectedOvertimeIds.has(o.id));
  } else {
    targetList = [...lastFetchedOvertimes];
  }

  if (!targetList || targetList.length === 0) {
    alert("🚨 [특근 엑셀 내보내기 안내]\\n\\n내보낼 특근 내역이 없습니다. 검색 필터를 확인해 주세요.");
    showToast('내보낼 특근 내역이 없습니다.', 'error');
    restoreExportBtn();
    return;
  }"""

if old_export_excel in app_code:
    app_code = app_code.replace(old_export_excel, new_export_excel)
    print("Updated exportExcelBtn start")
else:
    print("exportExcelBtn pattern not matched exactly")

# 2. Update exportUsersBtn and importUsersBtn block
old_users_block = """// 요구사항 12: 팀원 명부 엑셀 내보내기
document.getElementById('exportUsersBtn').addEventListener('click', () => {
  if (!allUsersCache || allUsersCache.length === 0) {
    showToast('등록된 팀원 정보가 없습니다.', 'error');
    return;
  }
  if (typeof XLSX === 'undefined') {
    showToast('엑셀 라이브러리를 로드 중입니다.', 'error');
    return;
  }

  const rows = allUsersCache.map((u, idx) => ({
    "순번": idx + 1,
    "사원번호": u.emp_id,
    "성명": u.name,
    "소속팀": u.team,
    "직급": u.position || '팀원',
    "관리자여부": (u.is_admin || u.is_super) ? '관리자' : '일반',
    "등록일시": u.created_at || '-'
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows);
  worksheet['!cols'] = [
    { wch: 6 }, { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 20 }
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "팀원명부");
  XLSX.writeFile(workbook, `팀원명부_${getTodayStr()}.xlsx`);
  showToast('팀원 명부 엑셀 다운로드가 완료되었습니다!');
});

// 요구사항 12: 팀원 명부 엑셀 가져오기 (업로드 파싱)
const userExcelFileInput = document.getElementById('userExcelFileInput');
document.getElementById('importUsersBtn').addEventListener('click', () => {
  userExcelFileInput.value = '';
  userExcelFileInput.click();
});

userExcelFileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  if (typeof XLSX === 'undefined') {
    showToast('엑셀 처리 도구를 로드 중입니다.', 'error');
    return;
  }

  showToast('엑셀 파일을 읽고 있습니다...');
  const reader = new FileReader();
  reader.onload = async (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const jsonRows = XLSX.utils.sheet_to_json(sheet);

      if (!jsonRows || jsonRows.length === 0) {
        showToast('엑셀 파일에 데이터가 없습니다.', 'error');
        return;
      }

      let successCount = 0;
      let errorCount = 0;

      for (const row of jsonRows) {
        const empId = String(row['사원번호'] || row['사번'] || row['emp_id'] || '').trim();
        const name = String(row['성명'] || row['이름'] || row['name'] || '').trim();
        const team = String(row['소속팀'] || row['부서'] || row['team'] || '').trim();
        const position = String(row['직급'] || row['position'] || '팀원').trim();
        const adminStr = String(row['관리자여부'] || row['관리자'] || row['is_admin'] || '').trim();
        const isAdmin = (adminStr === '관리자' || adminStr === '1' || adminStr.toLowerCase() === 'true') ? 1 : 0;

        if (!empId || !name || !team) {
          errorCount++;
          continue;
        }

        try {
          // 기존에 있는지 확인
          const exists = allUsersCache.some(u => u.emp_id.toLowerCase() === empId.toLowerCase());
          if (exists) {
            // 정보 업데이트
            await fetch(`/api/users/${empId}`, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ name, team, position, is_admin: isAdmin })
            });
          } else {
            // 신규 등록
            await fetch('/api/users/register', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ emp_id: empId, name, team, position, is_admin: isAdmin })
            });
          }
          successCount++;
        } catch (err) {
          errorCount++;
        }
      }

      showToast(`총 ${successCount}명의 팀원 정보가 등록/갱신되었습니다! (오류: ${errorCount}건)`);
      loadAdminUserTable();
    } catch (err) {
      console.error(err);
      showToast('엑셀 파일 파싱 중 오류가 발생했습니다.', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
});"""

new_users_block = """// 요구사항 12: 팀원 명부 엑셀 내보내기
document.getElementById('exportUsersBtn')?.addEventListener('click', () => {
  const btn = document.getElementById('exportUsersBtn');
  const restoreBtn = setButtonLoading(btn, '⏳ 팀원명부 엑셀 생성 중...');

  try {
    if (!allUsersCache || allUsersCache.length === 0) {
      alert("🚨 [팀원 명부 내보내기 안내]\\n\\n등록된 팀원 정보가 없습니다.");
      showToast('등록된 팀원 정보가 없습니다.', 'error');
      return;
    }
    if (typeof XLSX === 'undefined') {
      alert("🚨 [엑셀 모듈 오류]\\n\\nSheetJS 엑셀 처리 라이브러리를 로드할 수 없습니다. 페이지를 새로고침 해 주세요.");
      showToast('엑셀 라이브러리를 로드 중입니다.', 'error');
      return;
    }

    const rows = allUsersCache.map((u, idx) => ({
      "순번": idx + 1,
      "사원번호": u.emp_id,
      "성명": u.name,
      "소속팀": u.team,
      "직급": u.position || '팀원',
      "관리자여부": (u.is_admin || u.is_super) ? '관리자' : '일반',
      "등록일시": u.created_at || '-'
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet['!cols'] = [
      { wch: 6 }, { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 20 }
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "팀원명부");
    XLSX.writeFile(workbook, `팀원명부_${getTodayStr()}.xlsx`);
    showToast('팀원 명부 엑셀 다운로드가 완료되었습니다!');
  } catch (err) {
    console.error(err);
    alert("🚨 [팀원 명부 엑셀 내보내기 오류]\\n\\n" + err.message);
    showToast('팀원 명부 내보내기 실패', 'error');
  } finally {
    restoreBtn();
  }
});

// 요구사항 12: 팀원 명부 엑셀 가져오기 (전역 핸들러)
window.triggerUserExcelImport = function() {
  const input = document.getElementById('userExcelFileInput');
  if (input) {
    input.value = '';
    input.click();
  }
};

window.handleUserExcelFileSelected = async function(input) {
  const file = input?.files?.[0];
  if (!file) return;

  const btn = document.getElementById('importUsersBtn');
  const restoreBtn = setButtonLoading(btn, '⏳ 엑셀 읽는 중...');

  if (typeof XLSX === 'undefined') {
    alert("🚨 [엑셀 라이브러리 미로드]\\n\\nSheetJS 엑셀 파싱 모듈이 로드되지 않았습니다. 페이지를 새로고침 해 주세요.");
    showToast('엑셀 처리 도구를 로드 중입니다.', 'error');
    restoreBtn();
    return;
  }

  showToast('엑셀 파일을 읽고 있습니다...');
  const reader = new FileReader();
  reader.onload = async (evt) => {
    try {
      setButtonLoading(btn, '⏳ 팀원 정보 갱신 중...');
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const jsonRows = XLSX.utils.sheet_to_json(sheet);

      if (!jsonRows || jsonRows.length === 0) {
        alert("🚨 [팀원 명부 엑셀 가져오기 오류]\\n\\n선택하신 엑셀 파일에 데이터 행이 없습니다.");
        showToast('엑셀 파일에 데이터가 없습니다.', 'error');
        restoreBtn();
        return;
      }

      let successCount = 0;
      let errorCount = 0;

      for (const row of jsonRows) {
        const empId = String(row['사원번호'] || row['사번'] || row['emp_id'] || '').trim();
        const name = String(row['성명'] || row['이름'] || row['name'] || '').trim();
        const team = String(row['소속팀'] || row['부서'] || row['team'] || '').trim();
        const position = String(row['직급'] || row['position'] || '팀원').trim();
        const adminStr = String(row['관리자여부'] || row['관리자'] || row['is_admin'] || '').trim();
        const isAdmin = (adminStr === '관리자' || adminStr === '1' || adminStr.toLowerCase() === 'true') ? 1 : 0;

        if (!empId || !name || !team) {
          errorCount++;
          continue;
        }

        try {
          const exists = allUsersCache.some(u => u.emp_id.toLowerCase() === empId.toLowerCase());
          if (exists) {
            await fetch(`/api/users/${empId}`, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ name, team, position, is_admin: isAdmin })
            });
          } else {
            await fetch('/api/users/register', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ emp_id: empId, name, team, position, is_admin: isAdmin })
            });
          }
          successCount++;
        } catch (err) {
          errorCount++;
        }
      }

      showToast(`총 ${successCount}명의 팀원 정보가 등록/갱신되었습니다! (오류: ${errorCount}건)`);
      await loadAdminUserTable();
    } catch (err) {
      console.error(err);
      alert("🚨 [팀원 명부 엑셀 파싱 오류]\\n\\n" + err.message + "\\n\\n💡 [해결 조치]\\n올바른 팀원 명부 엑셀 서식인지 확인하세요.");
      showToast('엑셀 파일 파싱 중 오류가 발생했습니다.', 'error');
    } finally {
      restoreBtn();
    }
  };
  reader.onerror = (err) => {
    alert("🚨 [엑셀 파일 읽기 오류]\\n\\n" + (err?.message || '파일을 읽을 수 없습니다.'));
    restoreBtn();
  };
  reader.readAsArrayBuffer(file);
};

document.getElementById('importUsersBtn')?.addEventListener('click', window.triggerUserExcelImport);
document.getElementById('userExcelFileInput')?.addEventListener('change', function() { window.handleUserExcelFileSelected(this); });"""

if old_users_block in app_code:
    app_code = app_code.replace(old_users_block, new_users_block)
    print("Updated user excel import/export block")
else:
    print("user excel import/export block not matched exactly")

# Save updated app.js
with open('app.js', 'w', encoding='utf-8') as f:
    f.write(app_code)

print("Saved app.js")
