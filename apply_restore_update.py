import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('app.js', 'r', encoding='utf-8') as f:
    app_code = f.read()

old_restore_block = """    tbody.querySelectorAll('.btn-restore-backup').forEach(btn => {
      btn.addEventListener('click', async () => {
        const filename = btn.getAttribute('data-filename');
        const name = btn.getAttribute('data-name');
        if (!confirm(`'${name}' 스냅샷을 열어 현재 데이터베이스를 복원하시겠습니까?\\n\\n* 현재 데이터가 해당 스냅샷 시점으로 덮어쓰기됩니다.`)) {
          return;
        }

        try {
          const rRes = await fetch('/api/backup/load', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filename })
          });
          const rData = await rRes.json();
          if (!rRes.ok) {
            showToast(rData.detail || '복원 실패', 'error');
            return;
          }

          showToast(rData.message || '스냅샷이 성공적으로 복원되었습니다!');
          closeModal('backupModal');

          // 복원 후 화면 및 세션 갱신
          await loadTeams();
          if (currentUser) {
            const chkRes = await fetch('/api/users/login', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ emp_id: currentUser.emp_id })
            });
            const chkData = await chkRes.json();
            if (chkData.exists) {
              setUserSession(chkData.user);
            } else {
              showToast('복원된 데이터에 현재 로그인 정보가 없어 다시 로그인합니다.', 'warning');
              currentUser = null;
              try { localStorage.removeItem('overtime_session'); } catch (e) {}
              document.getElementById('mainDashboard').style.display = 'none';
              document.getElementById('loginScreen').style.display = 'flex';
            }
          }
          if (currentMode === 'admin') {
            await loadAdminData();
            await loadAdminUserTable();
          }
        } catch (err) {
          console.error(err);
          showToast('복원 중 통신 오류가 발생했습니다.', 'error');
        }
      });
    });"""

new_restore_block = """    tbody.querySelectorAll('.btn-restore-backup').forEach(btn => {
      btn.addEventListener('click', async () => {
        const filename = btn.getAttribute('data-filename');
        const name = btn.getAttribute('data-name');
        if (!confirm(`'${name}' 스냅샷을 열어 현재 데이터베이스를 복원하시겠습니까?\\n\\n* 현재 데이터가 해당 스냅샷 시점으로 덮어쓰기됩니다.`)) {
          return;
        }

        const restoreBtn = setButtonLoading(btn, '⏳ 스냅샷 복원(열기) 중...');

        try {
          const rRes = await fetch('/api/backup/load', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filename })
          });
          const rData = await rRes.json();
          if (!rRes.ok) {
            alert("🚨 [스냅샷 복원(열기) 오류]\\n\\n" + (rData.detail || '스냅샷 복원에 실패했습니다.') + "\\n\\n💡 [해결 조치]\\n해당 스냅샷 파일(data/backups/)이 정상 존재하는지 확인해 주세요.");
            showToast(rData.detail || '복원 실패', 'error');
            return;
          }

          showToast(rData.message || '스냅샷이 성공적으로 복원되었습니다!');
          closeModal('backupModal');

          // 복원 후 화면 및 세션 갱신
          await loadTeams();
          if (currentUser) {
            const chkRes = await fetch('/api/users/login', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ emp_id: currentUser.emp_id })
            });
            const chkData = await chkRes.json();
            if (chkData.exists) {
              setUserSession(chkData.user);
            } else {
              showToast('복원된 데이터에 현재 로그인 정보가 없어 다시 로그인합니다.', 'warning');
              currentUser = null;
              try { localStorage.removeItem('overtime_session'); } catch (e) {}
              document.getElementById('mainDashboard').style.display = 'none';
              document.getElementById('loginScreen').style.display = 'flex';
            }
          }
          if (currentMode === 'admin') {
            await loadAdminData();
            await loadAdminUserTable();
          }
        } catch (err) {
          console.error(err);
          alert("🚨 [스냅샷 복원 통신 오류]\\n\\n" + err.message);
          showToast('복원 중 통신 오류가 발생했습니다.', 'error');
        } finally {
          restoreBtn();
        }
      });
    });"""

if old_restore_block in app_code:
    app_code = app_code.replace(old_restore_block, new_restore_block)
    print("Updated restore backup block")
else:
    print("restore backup block not matched exactly")

with open('app.js', 'w', encoding='utf-8') as f:
    f.write(app_code)

print("Saved app.js")
