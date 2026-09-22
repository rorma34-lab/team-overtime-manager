import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

# 1. exportSummaryExcelBtn
old_summary_export = """// 실특근 정산표 엑셀 내보내기 - openpyxl 서버 엔진 전용
const exportSummaryExcelBtn = document.getElementById('exportSummaryExcelBtn');
if (exportSummaryExcelBtn) {
  exportSummaryExcelBtn.addEventListener('click', async () => {
    const users = lastFetchedSummary.user_summary || [];
    if (users.length === 0) {
      showToast('내보낼 정산 데이터가 없습니다.', 'error');
      return;
    }

    const sDate = document.getElementById('summaryStartDate')?.value || '';
    const eDate = document.getElementById('summaryEndDate')?.value || '';
    const tFilter = selectedDeptFilter || document.getElementById('summaryTeamFilter')?.value || document.getElementById('filterTeam')?.value || '';

    showToast('정산표 엑셀을 생성 중입니다...');

    try {
      const res = await fetch('/api/overtimes/export-settlement', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ start_date: sDate, end_date: eDate, team: tFilter })
      });

      if (res.ok) {
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const period = `${sDate || '전체'}_${eDate || '전체'}`;
        const today = new Date();
        const todayStr = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
        a.download = `실특근정산표_${period}_${todayStr}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
        showToast('실특근 정산표 엑셀 다운로드가 완료되었습니다!');
        return;
      }

      const errData = await res.json().catch(() => ({}));
      showToast(errData.detail || '정산표 엑셀 생성에 실패했습니다.', 'error');
    } catch (err) {
      console.error('Settlement export error:', err);
      showToast('엑셀 다운로드 중 오류가 발생했습니다.', 'error');
    }
  });
}"""

new_summary_export = """// 실특근 정산표 엑셀 내보내기 - openpyxl 서버 엔진 전용
const exportSummaryExcelBtn = document.getElementById('exportSummaryExcelBtn');
if (exportSummaryExcelBtn) {
  exportSummaryExcelBtn.addEventListener('click', async () => {
    const restoreBtn = setButtonLoading(exportSummaryExcelBtn, '⏳ 정산표 엑셀 생성 중...');
    const users = lastFetchedSummary.user_summary || [];
    if (users.length === 0) {
      alert("🚨 [실특근 정산표 내보내기 안내]\\n\\n내보낼 정산 데이터가 없습니다. 먼저 정산 조회를 진행해 주세요.");
      showToast('내보낼 정산 데이터가 없습니다.', 'error');
      restoreBtn();
      return;
    }

    const sDate = document.getElementById('summaryStartDate')?.value || '';
    const eDate = document.getElementById('summaryEndDate')?.value || '';
    const tFilter = (selectedDeptFilters && selectedDeptFilters.length > 0) ? selectedDeptFilters.join(',') : (selectedDeptFilter || document.getElementById('summaryTeamFilter')?.value || document.getElementById('filterTeam')?.value || '');

    showToast('정산표 엑셀을 생성 중입니다...');

    try {
      const res = await fetch('/api/overtimes/export-settlement', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ start_date: sDate, end_date: eDate, team: tFilter })
      });

      if (res.ok) {
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const period = `${sDate || '전체'}_${eDate || '전체'}`;
        const today = new Date();
        const todayStr = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
        a.download = `실특근정산표_${period}_${todayStr}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
        showToast('실특근 정산표 엑셀 다운로드가 완료되었습니다!');
        return;
      }

      const errData = await res.json().catch(() => ({}));
      alert("🚨 [실특근 정산표 엑셀 내보내기 오류]\\n\\n" + (errData.detail || '정산표 엑셀 생성에 실패했습니다.') + "\\n\\n💡 [해결 조치]\\nstart_server.bat 서버 및 openpyxl 라이브러리를 확인하세요.");
      showToast(errData.detail || '정산표 엑셀 생성에 실패했습니다.', 'error');
    } catch (err) {
      console.error('Settlement export error:', err);
      alert("🚨 [실특근 정산표 엑셀 통신 오류]\\n\\n" + err.message);
      showToast('엑셀 다운로드 중 오류가 발생했습니다.', 'error');
    } finally {
      restoreBtn();
    }
  });
}"""

if old_summary_export in app_code:
    app_code = app_code.replace(old_summary_export, new_summary_export)
    print("Updated exportSummaryExcelBtn")
else:
    print("exportSummaryExcelBtn pattern not matched exactly")

# 2. saveWebBackup
old_save_backup = """async function saveWebBackup() {
  const nameInput = document.getElementById('backupCustomNameInput');
  const customName = nameInput ? nameInput.value.trim() : '';
  const btn = document.getElementById('btnExecuteWebBackupSave');
  const restoreBtn = setButtonLoading(btn, '스냅샷 생성 저장 중...');

  try {
    const res = await fetch('/api/backup/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: customName, description: '웹 브라우저 수동 스냅샷 백업' })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '스냅샷 저장 실패', 'error');
      return;
    }

    showToast(`스냅샷 '${data.name}' 저장이 완료되었습니다! (사원: ${data.counts.users}명, 특근: ${data.counts.overtimes}건)`);
    if (nameInput) nameInput.value = '';
    await loadBackupList();
  } catch (err) {
    console.error(err);
    showToast('백업 저장 중 통신 오류가 발생했습니다.', 'error');
  } finally {
    restoreBtn();
  }
}"""

new_save_backup = """async function saveWebBackup() {
  const nameInput = document.getElementById('backupCustomNameInput');
  const customName = nameInput ? nameInput.value.trim() : '';
  const btn = document.getElementById('btnExecuteWebBackupSave');
  const restoreBtn = setButtonLoading(btn, '⏳ 스냅샷 저장 중...');

  try {
    const res = await fetch('/api/backup/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: customName, description: '웹 브라우저 수동 스냅샷 백업' })
    });
    const data = await res.json();
    if (!res.ok) {
      alert("🚨 [스냅샷 백업 저장 오류]\\n\\n" + (data.detail || '스냅샷 저장이 실패했습니다.') + "\\n\\n💡 [해결 조치]\\nstart_server.bat 백엔드 서버 상태 및 overtime.db 권한을 확인하세요.");
      showToast(data.detail || '스냅샷 저장 실패', 'error');
      return;
    }

    showToast(`스냅샷 '${data.name}' 저장이 완료되었습니다! (사원: ${data.counts.users}명, 특근: ${data.counts.overtimes}건)`);
    if (nameInput) nameInput.value = '';
    await loadBackupList();
  } catch (err) {
    console.error(err);
    alert("🚨 [스냅샷 백업 통신 오류]\\n\\n" + err.message);
    showToast('백업 저장 중 통신 오류가 발생했습니다.', 'error');
  } finally {
    restoreBtn();
  }
}"""

if old_save_backup in app_code:
    app_code = app_code.replace(old_save_backup, new_save_backup)
    print("Updated saveWebBackup")
else:
    print("saveWebBackup pattern not matched exactly")

# 3. exportAccessLogsBtn
old_access_logs = """// 엑셀 내보내기 이벤트 바인딩
const exportAccessLogsBtn = document.getElementById('exportAccessLogsBtn');
if (exportAccessLogsBtn) {
  exportAccessLogsBtn.addEventListener('click', () => {
    if (!currentUser) return;
    const startDate = document.getElementById('accessFilterStartDate')?.value || '';
    const endDate = document.getElementById('accessFilterEndDate')?.value || '';
    const actionType = document.getElementById('accessFilterAction')?.value || '';
    const status = document.getElementById('accessFilterStatus')?.value || '';
    const search = document.getElementById('accessFilterSearch')?.value || '';

    const params = [`admin_emp_id=${encodeURIComponent(currentUser.emp_id)}`];
    if (startDate) params.push(`start_date=${encodeURIComponent(startDate)}`);
    if (endDate) params.push(`end_date=${encodeURIComponent(endDate)}`);
    if (actionType) params.push(`action_type=${encodeURIComponent(actionType)}`);
    if (status) params.push(`status=${encodeURIComponent(status)}`);
    if (search) params.push(`search=${encodeURIComponent(search)}`);

    showToast('감사 로그 엑셀 다운로드를 시작합니다...');
    window.location.href = `/api/admin/access-logs/export?${params.join('&')}`;
  });
}"""

new_access_logs = """// 엑셀 내보내기 이벤트 바인딩
const exportAccessLogsBtn = document.getElementById('exportAccessLogsBtn');
if (exportAccessLogsBtn) {
  exportAccessLogsBtn.addEventListener('click', async () => {
    if (!currentUser) return;
    const restoreBtn = setButtonLoading(exportAccessLogsBtn, '⏳ 감사로그 엑셀 생성 중...');
    const startDate = document.getElementById('accessFilterStartDate')?.value || '';
    const endDate = document.getElementById('accessFilterEndDate')?.value || '';
    const actionType = document.getElementById('accessFilterAction')?.value || '';
    const status = document.getElementById('accessFilterStatus')?.value || '';
    const search = document.getElementById('accessFilterSearch')?.value || '';

    const params = [`admin_emp_id=${encodeURIComponent(currentUser.emp_id)}`];
    if (startDate) params.push(`start_date=${encodeURIComponent(startDate)}`);
    if (endDate) params.push(`end_date=${encodeURIComponent(endDate)}`);
    if (actionType) params.push(`action_type=${encodeURIComponent(actionType)}`);
    if (status) params.push(`status=${encodeURIComponent(status)}`);
    if (search) params.push(`search=${encodeURIComponent(search)}`);

    showToast('감사 로그 엑셀 다운로드를 시작합니다...');

    try {
      const res = await fetch(`/api/admin/access-logs/export?${params.join('&')}`);
      if (res.ok) {
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `보안감사로그_${getTodayStr()}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
        showToast('보안 감사 로그 다운로드가 완료되었습니다!');
        return;
      }
      const data = await res.json().catch(() => ({}));
      alert("🚨 [감사 로그 내보내기 오류]\\n\\n" + (data.detail || '감사 로그 엑셀 내보내기 실패') + "\\n\\n💡 [해결 조치]\\n관리자 권한 및 백엔드 서버 상태를 확인하세요.");
      showToast(data.detail || '다운로드 실패', 'error');
    } catch (err) {
      console.error(err);
      alert("🚨 [감사 로그 내보내기 통신 오류]\\n\\n" + err.message);
      showToast('감사 로그 다운로드 중 오류 발생', 'error');
    } finally {
      restoreBtn();
    }
  });
}"""

if old_access_logs in app_code:
    app_code = app_code.replace(old_access_logs, new_access_logs)
    print("Updated exportAccessLogsBtn")
else:
    print("exportAccessLogsBtn pattern not matched exactly")

with open('app.js', 'w', encoding='utf-8') as f:
    f.write(app_code)

print("Saved app.js")
