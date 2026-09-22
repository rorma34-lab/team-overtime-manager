import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

old_summary_import = """if (importSummaryExcelBtn && summaryExcelFileInput) {
  importSummaryExcelBtn.addEventListener('click', () => {
    summaryExcelFileInput.value = '';
    summaryExcelFileInput.click();
  });

  summaryExcelFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const targetTeam = selectedDeptFilter || document.getElementById('summaryTeamFilter')?.value || document.getElementById('filterTeam')?.value || '';

    // 1. 진행중 안내 모달 표시 (요구사항 4)
    const progressDeptBadge = document.getElementById('importProgressDeptBadge');
    if (progressDeptBadge) {
      progressDeptBadge.textContent = `🏢 부서 현황: ${targetTeam ? targetTeam : '전체 부서 (모든 부서 반영)'}`;
    }
    openModal('settlementImportProgressModal');

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const rawResult = evt.target.result || '';
        const base64Content = rawResult.includes(',') ? rawResult.split(',')[1] : rawResult;

        const res = await fetch('/api/overtimes/import-settlement', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            file_base64: base64Content,
            filename: file.name,
            target_team: targetTeam,
            admin_emp_id: currentUser ? currentUser.emp_id : ''
          })
        });

        closeModal('settlementImportProgressModal');

        const data = await res.json().catch(() => ({}));

      // 2. 결과 / 실패 사유 모달 표시 (요구사항 4)
      const banner = document.getElementById('importResultBanner');
      const statTotal = document.getElementById('importStatTotal');
      const statProcessed = document.getElementById('importStatProcessed');
      const statSkipped = document.getElementById('importStatSkipped');
      const statErrors = document.getElementById('importStatErrors');
      const deptRuleBox = document.getElementById('importResultDeptRuleBox');
      const errSection = document.getElementById('importErrorDetailsSection');
      const errList = document.getElementById('importErrorDetailsList');

      if (statTotal) statTotal.textContent = data.total_rows || 0;
      if (statProcessed) statProcessed.textContent = data.processed_count || 0;
      if (statSkipped) statSkipped.textContent = data.skipped_other_dept_count || 0;
      if (statErrors) statErrors.textContent = data.error_count || (res.ok ? 0 : 1);

      if (res.ok && data.success !== false) {
        if (banner) {
          banner.style.background = '#ecfdf5';
          banner.style.color = '#047857';
          banner.style.border = '1px solid #10b981';
          banner.innerHTML = `<span>✅</span> <span>정산표 엑셀 가져오기가 성공적으로 완료되었습니다!</span>`;
        }
        if (deptRuleBox) {
          if (targetTeam) {
            deptRuleBox.innerHTML = `🛡️ <b>부서 데이터 격리 적용</b>: 선택된 부서(<b>${escapeHtml(targetTeam)}</b>)의 <b>${data.processed_count || 0}건</b>만 반영되었으며, 타 부서 <b>${data.skipped_other_dept_count || 0}건</b>의 정보는 100% 변경 없이 안전하게 보존되었습니다.`;
          } else {
            deptRuleBox.innerHTML = `🛡️ <b>전체 부서 데이터 반영</b>: 엑셀 파일 내 모든 부서의 데이터가 수집/반영되었습니다.`;
          }
        }
        showToast('엑셀 정산 데이터가 정상 반영되었습니다!');
        loadSettlementSummary();
        loadAdminData();
      } else {
        if (banner) {
          banner.style.background = '#fef2f2';
          banner.style.color = '#b91c1c';
          banner.style.border = '1px solid #ef4444';
          banner.innerHTML = `<span>❌</span> <span>정산표 엑셀 가져오기 실패</span>`;
        }
        if (deptRuleBox) {
          const detailMsg = data.detail || data.message || '파일 처리 중 오류가 발생했습니다.';
          deptRuleBox.innerHTML = `⚠️ <b>실패 사유</b>: <span style="color:#b91c1c;">${escapeHtml(detailMsg)}</span>`;
        }
        showToast(data.detail || '가져오기 실패', 'error');
      }

      const errors = data.errors || [];
      if (errors.length > 0 && errSection && errList) {
        errSection.style.display = 'block';
        errList.innerHTML = errors.map(err => `<div>• ${escapeHtml(err)}</div>`).join('');
      } else if (errSection) {
        errSection.style.display = 'none';
      }

      openModal('settlementImportResultModal');
    } catch (err) {
      console.error('Import error:', err);
      closeModal('settlementImportProgressModal');
      showToast('엑셀 업로드 통신 오류가 발생했습니다.', 'error');
    }
    };
    reader.readAsDataURL(file);
  });
}"""

new_summary_import = """if (importSummaryExcelBtn && summaryExcelFileInput) {
  importSummaryExcelBtn.addEventListener('click', () => {
    summaryExcelFileInput.value = '';
    summaryExcelFileInput.click();
  });

  summaryExcelFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const restoreBtn = setButtonLoading(importSummaryExcelBtn, '⏳ 정산표 읽는 중...');
    const targetTeam = (selectedDeptFilters && selectedDeptFilters.length > 0) ? selectedDeptFilters.join(',') : (selectedDeptFilter || document.getElementById('summaryTeamFilter')?.value || document.getElementById('filterTeam')?.value || '');

    const progressDeptBadge = document.getElementById('importProgressDeptBadge');
    if (progressDeptBadge) {
      progressDeptBadge.textContent = `🏢 부서 현황: ${targetTeam ? targetTeam : '전체 부서 (모든 부서 반영)'}`;
    }
    openModal('settlementImportProgressModal');

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        setButtonLoading(importSummaryExcelBtn, '⏳ 정산 데이터 반영 중...');
        const rawResult = evt.target.result || '';
        const base64Content = rawResult.includes(',') ? rawResult.split(',')[1] : rawResult;

        const res = await fetch('/api/overtimes/import-settlement', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            file_base64: base64Content,
            filename: file.name,
            target_team: targetTeam,
            admin_emp_id: currentUser ? currentUser.emp_id : ''
          })
        });

        closeModal('settlementImportProgressModal');

        const data = await res.json().catch(() => ({}));

        const banner = document.getElementById('importResultBanner');
        const statTotal = document.getElementById('importStatTotal');
        const statProcessed = document.getElementById('importStatProcessed');
        const statSkipped = document.getElementById('importStatSkipped');
        const statErrors = document.getElementById('importStatErrors');
        const deptRuleBox = document.getElementById('importResultDeptRuleBox');
        const errSection = document.getElementById('importErrorDetailsSection');
        const errList = document.getElementById('importErrorDetailsList');

        if (statTotal) statTotal.textContent = data.total_rows || 0;
        if (statProcessed) statProcessed.textContent = data.processed_count || 0;
        if (statSkipped) statSkipped.textContent = data.skipped_other_dept_count || 0;
        if (statErrors) statErrors.textContent = data.error_count || (res.ok ? 0 : 1);

        if (res.ok && data.success !== false) {
          if (banner) {
            banner.style.background = '#ecfdf5';
            banner.style.color = '#047857';
            banner.style.border = '1px solid #10b981';
            banner.innerHTML = `<span>✅</span> <span>정산표 엑셀 가져오기가 성공적으로 완료되었습니다!</span>`;
          }
          if (deptRuleBox) {
            if (targetTeam) {
              deptRuleBox.innerHTML = `🛡️ <b>부서 데이터 격리 적용</b>: 선택된 부서(<b>${escapeHtml(targetTeam)}</b>)의 <b>${data.processed_count || 0}건</b>만 반영되었으며, 타 부서 <b>${data.skipped_other_dept_count || 0}건</b>의 정보는 100% 변경 없이 안전하게 보존되었습니다.`;
            } else {
              deptRuleBox.innerHTML = `🛡️ <b>전체 부서 데이터 반영</b>: 엑셀 파일 내 모든 부서의 데이터가 수집/반영되었습니다.`;
            }
          }
          showToast('엑셀 정산 데이터가 정상 반영되었습니다!');
          await loadSettlementSummary();
          await loadAdminData();
        } else {
          const detailMsg = data.detail || data.message || '파일 처리 중 오류가 발생했습니다.';
          alert("🚨 [실특근 정산표 가져오기 오류]\\n\\n" + detailMsg + "\\n\\n💡 [해결 조치]\\n업로드한 엑셀 파일이 실특근 정산표 양식인지 확인해 주세요.");
          if (banner) {
            banner.style.background = '#fef2f2';
            banner.style.color = '#b91c1c';
            banner.style.border = '1px solid #ef4444';
            banner.innerHTML = `<span>❌</span> <span>정산표 엑셀 가져오기 실패</span>`;
          }
          if (deptRuleBox) {
            deptRuleBox.innerHTML = `⚠️ <b>실패 사유</b>: <span style="color:#b91c1c;">${escapeHtml(detailMsg)}</span>`;
          }
          showToast(detailMsg, 'error');
        }

        const errors = data.errors || [];
        if (errors.length > 0 && errSection && errList) {
          errSection.style.display = 'block';
          errList.innerHTML = errors.map(err => `<div>• ${escapeHtml(err)}</div>`).join('');
        } else if (errSection) {
          errSection.style.display = 'none';
        }

        openModal('settlementImportResultModal');
      } catch (err) {
        console.error('Import error:', err);
        closeModal('settlementImportProgressModal');
        alert("🚨 [실특근 정산표 업로드 통신 오류]\\n\\n" + err.message + "\\n\\n💡 [해결 조치]\\nstart_server.bat 백엔드 서버가 켜져 있는지 확인하세요.");
        showToast('엑셀 업로드 통신 오류가 발생했습니다.', 'error');
      } finally {
        restoreBtn();
      }
    };
    reader.readAsDataURL(file);
  });
}"""

if old_summary_import in app_code:
    app_code = app_code.replace(old_summary_import, new_summary_import)
    print("Updated summary import block")
else:
    print("summary import block not matched exactly")

with open('app.js', 'w', encoding='utf-8') as f:
    f.write(app_code)

print("Saved app.js")
