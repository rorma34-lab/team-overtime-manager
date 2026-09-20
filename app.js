// ===== 전역 오류 핸들러 (JS 오류 발생 시 화면에 표시) =====
window.onerror = function(msg, src, line, col, err) {
  var box = document.createElement('div');
  box.style.cssText = 'position:fixed;top:0;left:0;right:0;background:#ef4444;color:#fff;padding:12px;z-index:99999;font-size:13px;word-break:break-all;';
  box.innerHTML = '<b>JS 오류 발생!</b> ' + msg + ' (L' + line + ':' + col + ') ' + src;
  document.body.appendChild(box);
  return false;
};

// ===== 버튼 로딩 스피너 및 진행 피드백 유틸리티 =====
function setButtonLoading(btn, loadingText = '처리 중...') {
  if (!btn) return () => {};
  const originalHtml = btn.innerHTML;
  const originalDisabled = btn.disabled;
  btn.disabled = true;
  btn.innerHTML = `<span style="display:inline-block;width:0.85rem;height:0.85rem;border:2px solid currentColor;border-right-color:transparent;border-radius:50%;animation:spin 0.6s linear infinite;vertical-align:middle;margin-right:6px;"></span>${loadingText}`;
  return function restore() {
    btn.disabled = originalDisabled;
    btn.innerHTML = originalHtml;
  };
}

// ===== 상태 변수 =====
let currentUser = null;
let currentMode = 'user'; // 'user' 또는 'admin'
let adminViewMode = 'calendar'; // 'calendar' 또는 'table'

// 달력 상태 (사용자 신청용)
let calYear = new Date().getFullYear();
let calMonth = new Date().getMonth();
let calStartDate = null;
let calEndDate = null;

// 관리자 월간 캘린더 상태
let adminCalYear = new Date().getFullYear();
let adminCalMonth = new Date().getMonth();
let adminSelectedDate = null; // 'YYYY-MM-DD'

// 부서별 필터 상태
let selectedDeptFilter = ''; // ''이면 전체

// 정렬 상태
let otSortCol = 'start_date';
let otSortAsc = false;
let userSortCol = 'emp_id';
let userSortAsc = true;
let summarySortCol = 'actual_overtime_days';
let summarySortAsc = false;

// 캐시 및 선택 상태
let selectedOvertimeIds = new Set();
let allUsersCache = [];
let lastFetchedOvertimes = [];
let lastFetchedSummary = { user_summary: [], team_summary: [], total_records: 0 };
let teamsCache = [];
let isAddingMemberFromAdmin = false;
let userOvertimesData = [];
let userCalYear = new Date().getFullYear();
let userCalMonth = new Date().getMonth();

// ===== 소속팀(부서) 관리 함수 (요구사항 25) =====
async function loadTeams() {
  try {
    const res = await fetch('/api/teams');
    if (!res.ok) return;
    const data = await res.json();
    teamsCache = (data.teams || []).map(t => typeof t === 'string' ? t : t.name);

    // 1. 회원가입 리스트박스 채우기 (요구사항 25)
    const regTeamSel = document.getElementById('regTeam');
    if (regTeamSel) {
      const curVal = regTeamSel.value;
      regTeamSel.innerHTML = '<option value="">소속팀을 선택하세요</option>';
      teamsCache.forEach(teamName => {
        const opt = document.createElement('option');
        opt.value = teamName;
        opt.textContent = teamName;
        regTeamSel.appendChild(opt);
      });
      if (curVal && teamsCache.includes(curVal)) {
        regTeamSel.value = curVal;
      } else {
        regTeamSel.value = '';
      }
    }

    // 2. 관리자 소속팀 관리 목록 채우기
    renderTeamManageList();

    // 3. 필터 셀렉트박스 채우기
    const filterTeam = document.getElementById('filterTeam');
    if (filterTeam) {
      const cur = filterTeam.value;
      filterTeam.innerHTML = '<option value="">전체 소속팀</option>';
      teamsCache.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t;
        opt.textContent = t;
        filterTeam.appendChild(opt);
      });
      if (cur) filterTeam.value = cur;
    }

    const summaryTeam = document.getElementById('summaryTeamFilter');
    if (summaryTeam) {
      const cur = summaryTeam.value;
      summaryTeam.innerHTML = '<option value="">전체 부서</option>';
      teamsCache.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t;
        opt.textContent = t;
        summaryTeam.appendChild(opt);
      });
      if (cur) summaryTeam.value = cur;
    }
  } catch (err) {
    console.error('Failed to load teams:', err);
  }
}

function renderTeamManageList() {
  const ul = document.getElementById('teamManageList');
  const countSpan = document.getElementById('teamListCount');
  if (!ul) return;
  if (countSpan) countSpan.textContent = teamsCache.length;
  ul.innerHTML = '';
  if (teamsCache.length === 0) {
    ul.innerHTML = '<li style="padding:1rem; text-align:center; color:var(--text-muted);">등록된 소속팀이 없습니다.</li>';
    return;
  }
  teamsCache.forEach(tName => {
    const li = document.createElement('li');
    li.style.cssText = 'display:flex; justify-content:space-between; align-items:center; padding:0.6rem 0.85rem; border-bottom:1px solid var(--border-color); font-size:0.9rem;';
    li.innerHTML = `
      <span style="font-weight:600; color:var(--text-main);">🏢 ${escapeHtml(tName)}</span>
      <button type="button" class="btn btn-danger btn-sm delete-team-btn" style="padding:3px 10px; font-size:0.75rem;" data-team="${escapeHtml(tName)}">삭제</button>
    `;
    ul.appendChild(li);
  });

  // 소속팀 삭제 버튼 이벤트 바인딩
  ul.querySelectorAll('.delete-team-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const tName = btn.getAttribute('data-team');
      await handleDeleteTeam(tName);
    });
  });
}

async function handleDeleteTeam(teamName) {
  if (!confirm(`'${teamName}' 소속팀을 정말 삭제하시겠습니까?\n(해당 팀에 등록된 팀원이 없어야 삭제 가능합니다.)`)) return;
  if (!currentUser || (!currentUser.is_admin && !currentUser.is_super)) {
    showToast('관리자 권한이 필요합니다.', 'error');
    return;
  }
  try {
    const res = await fetch(`/api/teams/${encodeURIComponent(teamName)}?admin_emp_id=${encodeURIComponent(currentUser.emp_id)}`, {
      method: 'DELETE'
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '삭제 실패', 'error');
      return;
    }
    showToast(data.message);
    await loadTeams();
    renderTeamManageList();
    if (currentMode === 'admin') {
      await loadAdminUserTable();
    }
  } catch (err) {
    console.error(err);
    showToast('통신 오류', 'error');
  }
}
window.handleDeleteTeam = handleDeleteTeam;


// ===== DOM 유틸리티 및 토스트 =====
function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `<span>${type === 'success' ? '✓' : '⚠️'}</span> <span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 2700);
}

function openModal(id) {
  const m = document.getElementById(id);
  if (m) m.classList.add('active');
}

function closeModal(id) {
  const m = document.getElementById(id);
  if (m) m.classList.remove('active');
}

document.querySelectorAll('[data-close]').forEach(btn => {
  btn.addEventListener('click', () => {
    closeModal(btn.getAttribute('data-close'));
  });
});

document.querySelectorAll('.modal-overlay').forEach(modal => {
  modal.addEventListener('click', (e) => {
    if (e.target === modal) modal.classList.remove('active');
  });
});

function formatDateStr(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function getTodayStr() {
  return formatDateStr(new Date());
}

// 해당 주의 토요일 자동 산출 (요구사항 4)
function getThisSaturdayStr() {
  const now = new Date();
  const day = now.getDay(); // 0: 일, 1: 월, ..., 6: 토
  const diff = 6 - day;
  const sat = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diff);
  return formatDateStr(sat);
}

// 사원번호 클라이언트 검증: 1 또는 2로 시작하는 6자리 숫자 (슈퍼관리자 ps37082 제외)
// 규칙 안내 문구는 노출하지 않고, 불일치 시 '입력이 올바르지 않습니다.'만 표시
function validateEmpIdClient(empId) {
  empId = (empId || '').trim();
  if (empId.toLowerCase() === 'ps37082') return true;
  const regex = /^[12]\d{5}$/;
  if (!regex.test(empId)) {
    showToast('입력이 올바르지 않습니다.', 'warning');
    return false;
  }
  return true;
}

// 직전 작성 특근 내역 가이드 자동 채우기 (요구사항 5)
async function loadLatestOvertimeGuide(empId) {
  try {
    const res = await fetch(`/api/users/${encodeURIComponent(empId)}/latest-overtime`);
    if (!res.ok) return;
    const data = await res.json();
    const noticeEl = document.getElementById('recentGuideNotice');
    const pInput = document.getElementById('projectNoInput');
    const lInput = document.getElementById('locationInput');
    const rInput = document.getElementById('reasonInput');

    if (data.has_previous && data.latest) {
      if (pInput && !pInput.value) pInput.value = data.latest.project_no || '';
      if (lInput && !lInput.value) lInput.value = data.latest.location || '';
      if (rInput && !rInput.value) rInput.value = data.latest.reason || '';
      if (noticeEl) noticeEl.style.display = 'block';
    } else {
      if (noticeEl) noticeEl.style.display = 'none';
    }
  } catch (err) {
    console.error('직전 작성 가이드 로딩 실패:', err);
  }
}

function escapeHtml(str) {
  return (str || '').replace(/[&<>"']/g, function(m) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
  });
}

// ===== 1. 로그인 플로우 =====
const loginForm = document.getElementById('loginForm');
loginForm.addEventListener('submit', (e) => {
  e.preventDefault();
  const empId = document.getElementById('empIdInput').value.trim();
  if (empId) handleLogin(empId);
});

async function handleLogin(empId) {
  empId = (empId || '').trim();
  if (!validateEmpIdClient(empId)) return;

  try {
    const res = await fetch('/api/users/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ emp_id: empId })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '로그인 중 오류가 발생했습니다.', 'error');
      return;
    }

    if (data.exists) {
      setUserSession(data.user);
      showToast(`${data.user.name}님 환영합니다!`);
    } else {
      isAddingMemberFromAdmin = false;
      const modal = document.getElementById('registerModal');
      const title = modal.querySelector('.modal-title');
      if (title) title.textContent = '🎉 신규 팀원 등록';
      const desc = modal.querySelector('.modal-body > p');
      if (desc) desc.textContent = '등록되지 않은 사번입니다. 정보를 입력하시면 즉시 회원 등록 및 입장됩니다.';
      const submitBtn = modal.querySelector('button[type="submit"]');
      if (submitBtn) submitBtn.textContent = '등록 및 입장하기';

      const regEmpId = document.getElementById('regEmpId');
      regEmpId.value = empId;
      regEmpId.setAttribute('readonly', 'true');
      document.getElementById('regName').value = '';
      await loadTeams();
      document.getElementById('regTeam').value = '';
      document.getElementById('regPosition').value = '팀원';
      openModal('registerModal');
    }
  } catch (err) {
    console.error(err);
    showToast('서버와 통신할 수 없습니다.', 'error');
  }
}

const registerForm = document.getElementById('registerForm');
registerForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const empId = document.getElementById('regEmpId').value.trim();
  const name = document.getElementById('regName').value.trim();
  const team = document.getElementById('regTeam').value.trim();
  const position = document.getElementById('regPosition').value.trim();

  if (!empId) {
    showToast('사원번호를 입력해주세요.', 'error');
    return;
  }
  if (!validateEmpIdClient(empId)) return;

  if (!name) {
    showToast('성명을 입력해주세요.', 'error');
    return;
  }
  if (!team) {
    showToast('소속팀을 선택해주세요.', 'error');
    return;
  }

  try {
    const res = await fetch('/api/users/register', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ emp_id: empId, name, team, position })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '등록 실패', 'error');
      return;
    }
    closeModal('registerModal');

    if (isAddingMemberFromAdmin) {
      showToast(`팀원 '${data.user.name}(${data.user.emp_id})' 등록이 완료되었습니다.`);
      await loadAdminUserTable();
      await loadTeams();
    } else {
      setUserSession(data.user);
      showToast(`팀원 등록이 완료되었습니다. 환영합니다, ${data.user.name}님!`);
    }
  } catch (err) {
    console.error(err);
    showToast('등록 중 오류 발생', 'error');
  }
});

function setUserSession(user) {
  currentUser = user;
  // 세션 영구 보존 (모바일 새로고침/창 전환 시에도 유지 - 요구사항 1, 6)
  try {
    localStorage.setItem('overtime_session', JSON.stringify(user));
  } catch (e) {
    console.warn('localStorage 접근 불가:', e);
  }
  
  document.getElementById('displayUserName').textContent = user.name;
  document.getElementById('displayEmpId').textContent = user.emp_id;
  document.getElementById('displayTeam').textContent = user.team;
  document.getElementById('displayPosition').textContent = user.position || '팀원';
  document.getElementById('avatarInitial').textContent = user.name ? user.name[0] : 'U';

  const adminBadge = document.getElementById('adminBadge');
  const modeTabsWrap = document.getElementById('modeTabsWrap');
  const manageTeamsTopBtnEl = document.getElementById('manageTeamsTopBtn');
  const manageTeamsBtnEl = document.getElementById('manageTeamsBtn');

  // 관리자 2등급제 구분 (요구사항 3, 9, 10)
  const isSuper = user.is_super === 1 || user.emp_id.toLowerCase() === 'ps37082';
  const isAdmin = user.is_admin === 1;

  if (isSuper) {
    adminBadge.style.display = 'inline-block';
    adminBadge.textContent = '👑 총괄 슈퍼관리자';
    adminBadge.style.background = 'linear-gradient(135deg, #4338ca, #6366f1)';
    modeTabsWrap.style.display = 'flex';
  } else if (isAdmin) {
    adminBadge.style.display = 'inline-block';
    adminBadge.textContent = `🛡️ 팀관리자 (${user.team})`;
    adminBadge.style.background = 'linear-gradient(135deg, #0284c7, #38bdf8)';
    modeTabsWrap.style.display = 'flex';
  } else {
    adminBadge.style.display = 'none';
    modeTabsWrap.style.display = 'none';
  }

  // 소속팀 관리는 오직 슈퍼관리자만 가능
  if (manageTeamsBtnEl) {
    manageTeamsBtnEl.style.display = isSuper ? 'inline-block' : 'none';
  }
  if (manageTeamsTopBtnEl) {
    manageTeamsTopBtnEl.style.display = isSuper ? 'inline-block' : 'none';
  }
  
  // 슈퍼관리자 전용 보안 감사 로그 탭 제어 (신규 요구사항 5)
  const adminSubTabAccessLogs = document.getElementById('adminSubTabAccessLogs');
  if (adminSubTabAccessLogs) {
    adminSubTabAccessLogs.style.display = isSuper ? 'inline-block' : 'none';
  }

  loadTeams();

  // 로그인 성공 시 상단 헤더 액션 버튼 표시 (로그인 전 화면에서는 숨김)
  const headerNavActions = document.getElementById('headerNavActions');
  if (headerNavActions) {
    headerNavActions.style.display = 'flex';
  }

  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('mainDashboard').style.display = 'block';
  document.getElementById('logoutBtn').style.display = 'inline-flex';

  switchMode('user');

  // 신규 신청 기본값 설정: 해당 주의 토요일로 세팅 (요구사항 4)
  const thisSat = getThisSaturdayStr();
  document.getElementById('startDateInput').value = thisSat;
  document.getElementById('endDateInput').value = thisSat;
  calStartDate = thisSat;
  calEndDate = thisSat;
  renderCalendar();

  // 신규 특근 신청 시 '일반휴일' 기본 선택 (요구사항 3)
  const cat2 = document.getElementById('cat2');
  if (cat2) cat2.checked = true;

  loadUserOvertimes();
}

document.getElementById('logoutBtn').addEventListener('click', async () => {
  // 로그아웃 감사 로그 기록
  if (currentUser && currentUser.emp_id) {
    try {
      fetch('/api/users/logout', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ emp_id: currentUser.emp_id })
      });
    } catch (e) {
      console.warn('Logout log error:', e);
    }
  }

  currentUser = null;
  try {
    localStorage.removeItem('overtime_session');
  } catch (e) {}

  // 로그인 전 상태로 초기화: 상단 헤더 액션 도구 완전 숨김
  const headerNavActions = document.getElementById('headerNavActions');
  if (headerNavActions) {
    headerNavActions.style.display = 'none';
  }

  document.getElementById('mainDashboard').style.display = 'none';
  document.getElementById('loginScreen').style.display = 'flex';
  document.getElementById('logoutBtn').style.display = 'none';
  document.getElementById('empIdInput').value = '';
  const topManualBtn = document.getElementById('topManualBtn');
  if (topManualBtn) {
    topManualBtn.href = '/api/manual/user';
    topManualBtn.textContent = '📖 사용자 매뉴얼';
    topManualBtn.title = '사용자 매뉴얼 다운로드 (.pptx)';
  }
});

document.getElementById('brandHomeBtn').addEventListener('click', () => {
  if (currentUser) switchMode('user');
});

// ===== 2. 모드 전환 (사용자 ↔ 관리자) =====
const tabUserMode = document.getElementById('tabUserMode');
const tabAdminMode = document.getElementById('tabAdminMode');

tabUserMode.addEventListener('click', () => switchMode('user'));
tabAdminMode.addEventListener('click', () => switchMode('admin'));

function switchMode(mode) {
  currentMode = mode;
  const topManualBtn = document.getElementById('topManualBtn');
  if (mode === 'user') {
    if (topManualBtn) {
      topManualBtn.href = '/api/manual/user';
      topManualBtn.textContent = '📖 사용자 매뉴얼';
      topManualBtn.title = '사용자 매뉴얼 다운로드 (.pptx)';
    }
    tabUserMode.classList.add('active');
    tabAdminMode.classList.remove('active');
    document.getElementById('userModeSection').style.display = 'block';
    document.getElementById('adminModeSection').style.display = 'none';
    document.querySelectorAll('.admin-only-tool').forEach(el => el.style.display = 'none');
    loadUserOvertimes();
  } else {
    if (topManualBtn) {
      topManualBtn.href = '/api/manual/admin';
      topManualBtn.textContent = '📖 관리자 매뉴얼';
      topManualBtn.title = '관리자 운영 매뉴얼 다운로드 (.pptx)';
    }
    tabAdminMode.classList.add('active');
    tabUserMode.classList.remove('active');
    document.getElementById('userModeSection').style.display = 'none';
    document.getElementById('adminModeSection').style.display = 'block';
    document.querySelectorAll('.admin-only-tool').forEach(el => el.style.display = 'inline-block');

    // 관리자 등급별 상단 안내 바 및 권한 설정 (요구사항 2, 3)
    const isSuper = currentUser && (currentUser.is_super === 1 || currentUser.emp_id.toLowerCase() === 'ps37082');
    const roleBadge = document.getElementById('adminRoleHeaderBadge');
    const roleDesc = document.getElementById('adminRoleHeaderDesc');
    const deptInfo = document.getElementById('adminCurrentDeptInfo');
    const superTeamSec = document.getElementById('superAdminTeamSection');

    if (isSuper) {
      if (roleBadge) {
        roleBadge.textContent = '👑 총괄 슈퍼관리자';
        roleBadge.style.background = '#fbbf24';
        roleBadge.style.color = '#78350f';
      }
      if (roleDesc) roleDesc.textContent = '전 부서의 특근 승인, 소속팀 관리 및 팀관리자 지정 권한을 총괄 제어합니다.';
      if (deptInfo) deptInfo.textContent = '권한: 전체 부서 총괄 통제';
      if (superTeamSec) superTeamSec.style.display = 'block';
    } else {
      if (roleBadge) {
        roleBadge.textContent = `🛡️ 팀관리자 (${currentUser ? currentUser.team : ''})`;
        roleBadge.style.background = '#38bdf8';
        roleBadge.style.color = '#0c4a6e';
      }
      if (roleDesc) roleDesc.textContent = `본인 소속팀(${currentUser ? currentUser.team : ''})의 특근만 승인/관리할 수 있습니다.`;
      if (deptInfo) deptInfo.textContent = `소속: ${currentUser ? currentUser.team : ''}`;
      if (superTeamSec) superTeamSec.style.display = 'none';
    }

    loadAdminDashboard();
  }
}

// ===== 3. 사용자 신청용 인터랙티브 달력 =====
const calPrevBtn = document.getElementById('calPrevBtn');
const calNextBtn = document.getElementById('calNextBtn');
const calMonthTitle = document.getElementById('calMonthTitle');
const calendarGrid = document.getElementById('calendarGrid');

calPrevBtn.addEventListener('click', () => {
  calMonth--;
  if (calMonth < 0) { calMonth = 11; calYear--; }
  renderCalendar();
});

calNextBtn.addEventListener('click', () => {
  calMonth++;
  if (calMonth > 11) { calMonth = 0; calYear++; }
  renderCalendar();
});

function renderCalendar() {
  calMonthTitle.textContent = `${calYear}년 ${calMonth + 1}월`;
  while (calendarGrid.children.length > 7) {
    calendarGrid.removeChild(calendarGrid.lastChild);
  }

  const firstDayIndex = new Date(calYear, calMonth, 1).getDay();
  const lastDate = new Date(calYear, calMonth + 1, 0).getDate();
  const prevLastDate = new Date(calYear, calMonth, 0).getDate();
  const todayStr = getTodayStr();

  for (let x = firstDayIndex; x > 0; x--) {
    const cell = document.createElement('div');
    cell.className = 'cal-cell disabled';
    cell.textContent = prevLastDate - x + 1;
    calendarGrid.appendChild(cell);
  }

  for (let d = 1; d <= lastDate; d++) {
    const dateStr = `${calYear}-${String(calMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dateObj = new Date(calYear, calMonth, d);
    const dayOfWeek = dateObj.getDay();

    const cell = document.createElement('div');
    cell.className = 'cal-cell';
    cell.textContent = d;
    cell.setAttribute('data-date', dateStr);

    if (dayOfWeek === 0) cell.classList.add('sun');
    if (dayOfWeek === 6) cell.classList.add('sat');
    if (dateStr === todayStr) cell.classList.add('today');

    if (calStartDate && dateStr === calStartDate) cell.classList.add('selected-start');
    if (calEndDate && dateStr === calEndDate) cell.classList.add('selected-end');
    if (calStartDate && calEndDate && dateStr > calStartDate && dateStr < calEndDate) cell.classList.add('in-range');

    cell.addEventListener('click', () => handleCalCellClick(dateStr));
    calendarGrid.appendChild(cell);
  }
}

function handleCalCellClick(dateStr) {
  if (!calStartDate || (calStartDate && calEndDate && calStartDate !== calEndDate)) {
    // 1단계: 첫 번째 날짜 선택 -> 시작일과 종료일 모두 해당 날짜로 우선 자동 설정
    calStartDate = dateStr;
    calEndDate = dateStr;
  } else if (calStartDate && calStartDate === calEndDate) {
    // 2단계: 두 번째 날짜 선택 -> 범위 설정
    if (dateStr >= calStartDate) {
      calEndDate = dateStr;
    } else {
      calStartDate = dateStr;
    }
  }
  document.getElementById('startDateInput').value = calStartDate;
  document.getElementById('endDateInput').value = calEndDate;
  renderCalendar();
}

// 시작일 변경 시: 종료일도 동일 날짜로 우선 자동 변경 (요구사항 2)
document.getElementById('startDateInput').addEventListener('change', (e) => {
  const newDate = e.target.value;
  if (!newDate) return;
  
  if (calStartDate && calEndDate && calStartDate === calEndDate) {
    // 이미 1개 날짜만 선택된 상태에서 시작일을 바꿨다면 -> 종료일도 동일하게 변경
    calStartDate = newDate;
    calEndDate = newDate;
    document.getElementById('endDateInput').value = newDate;
  } else if (calStartDate && calEndDate && calStartDate !== calEndDate) {
    // 이미 범위가 설정된 상태에서 시작일을 바꿀 때:
    // 만약 새 시작일이 종료일보다 늦으면 종료일도 시작일로 자동 변경
    calStartDate = newDate;
    if (calStartDate > calEndDate) {
      calEndDate = newDate;
      document.getElementById('endDateInput').value = newDate;
    }
  } else {
    // 기본: 시작일과 종료일 동일 동기화
    calStartDate = newDate;
    calEndDate = newDate;
    document.getElementById('endDateInput').value = newDate;
  }

  syncCalYearMonth(calStartDate);
  renderCalendar();
});

// 종료일 변경 시: 첫 선택 시 시작일도 동기화, 또는 두 번째 날짜로 범위 설정 (요구사항 2)
document.getElementById('endDateInput').addEventListener('change', (e) => {
  const newDate = e.target.value;
  if (!newDate) return;

  if (calStartDate && calEndDate && calStartDate === calEndDate) {
    // 시작일과 동일했던 상태에서 사용자가 종료일을 두 번째 날짜로 변경한 경우
    if (newDate >= calStartDate) {
      calEndDate = newDate; // 범위 설정 완료
    } else {
      // 종료일을 시작일보다 이전 날짜로 선택한 경우: 시작일도 함께 변경
      calStartDate = newDate;
      calEndDate = newDate;
      document.getElementById('startDateInput').value = newDate;
    }
  } else if (!calStartDate) {
    // 시작일이 없던 경우: 시작일도 동일하게 자동 설정
    calStartDate = newDate;
    calEndDate = newDate;
    document.getElementById('startDateInput').value = newDate;
  } else {
    // 이미 범위가 있던 상태에서 종료일을 새로 고른 경우
    if (newDate < calStartDate) {
      calStartDate = newDate;
      calEndDate = newDate;
      document.getElementById('startDateInput').value = newDate;
    } else {
      calEndDate = newDate;
    }
  }

  syncCalYearMonth(calEndDate);
  renderCalendar();
});

function syncCalYearMonth(dateStr) {
  if (!dateStr) return;
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    calYear = parseInt(parts[0], 10);
    calMonth = parseInt(parts[1], 10) - 1;
  }
}

// ===== 4. 특근 신청 저장 및 버전 모달 바인딩 =====
const versionBadgeBtn = document.getElementById('versionBadgeBtn');
if (versionBadgeBtn) {
  versionBadgeBtn.addEventListener('click', () => openModal('versionModal'));
}

// 요구사항 4: 실시간 전체 데이터 새로고침 (동시 접속 동기화)
const globalSyncBtn = document.getElementById('globalSyncBtn');
if (globalSyncBtn) {
  globalSyncBtn.addEventListener('click', async () => {
    showToast('최신 데이터를 동기화하는 중입니다...', 'info');
    try {
      await loadTeams();
      if (currentMode === 'admin') {
        await loadAdminData();
        await loadAdminUserTable();
        if (typeof loadSettlementSummary === 'function') await loadSettlementSummary();
      } else {
        await loadUserOvertimes();
      }
      showToast('최신 정보로 100% 동기화되었습니다!');
    } catch (e) {
      console.error(e);
      showToast('동기화 중 오류가 발생했습니다.', 'error');
    }
  });
}

const editSubHolidayDate = document.getElementById('editSubHolidayDate');
const editSubHolidayUsed = document.getElementById('editSubHolidayUsed');
if (editSubHolidayDate && editSubHolidayUsed) {
  editSubHolidayDate.addEventListener('change', (e) => {
    if (e.target.value) {
      const cur = parseFloat(editSubHolidayUsed.value || 0);
      if (cur <= 0) editSubHolidayUsed.value = '1';
    }
  });
}

const overtimeForm = document.getElementById('overtimeForm');
overtimeForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (!currentUser) {
    showToast('로그인이 필요합니다. 사원번호로 먼저 입장해주세요.', 'warning');
    document.getElementById('mainDashboard').style.display = 'none';
    document.getElementById('loginScreen').style.display = 'flex';
    return;
  }

  const catEl = document.querySelector('input[name="category"]:checked');
  const category = catEl ? catEl.value : '일반휴일';
  const startDate = document.getElementById('startDateInput').value;
  const endDate = document.getElementById('endDateInput').value;
  const projectNo = document.getElementById('projectNoInput').value.trim();
  const location = document.getElementById('locationInput').value.trim();
  const reason = document.getElementById('reasonInput').value.trim();

  if (!startDate || !endDate) {
    showToast('특근 시작일과 종료일을 입력해주세요.', 'error');
    return;
  }
  if (startDate > endDate) {
    showToast('종료일은 시작일보다 빠를 수 없습니다.', 'error');
    return;
  }
  if (!reason) {
    showToast('특근 사유를 입력해주세요.', 'error');
    return;
  }

  // 요구사항 5: 동일/겹치는 날짜 중복 신청 사전 차단
  if (Array.isArray(userOvertimesData) && userOvertimesData.length > 0) {
    const conflict = userOvertimesData.find(item => item.start_date <= endDate && item.end_date >= startDate);
    if (conflict) {
      showToast(`이미 해당 기간(${conflict.start_date} ~ ${conflict.end_date})에 신청된 특근이 존재합니다. 중복 신청은 불가능합니다.`, 'error');
      return;
    }
  }

  const submitBtn = overtimeForm.querySelector('button[type="submit"]');
  const restoreBtn = setButtonLoading(submitBtn, '신청 저장 중...');

  try {
    const res = await fetch('/api/overtimes', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        emp_id: currentUser.emp_id,
        category,
        start_date: startDate,
        end_date: endDate,
        project_no: projectNo,
        location,
        reason,
        sub_holiday_date: '',
        sub_holiday_used: 0.0
      })
    });
    let data = {};
    try {
      data = await res.json();
    } catch (parseErr) {
      const text = await res.text().catch(() => '');
      data = { detail: text || `서버 응답 오류 (${res.status})` };
    }
    if (!res.ok) {
      showToast(data.detail || `특근 신청 저장에 실패했습니다. (${res.status})`, 'error');
      return;
    }

    showToast('특근 신청이 정상 저장되었습니다!');
    // 입력 필드 초기화
    document.getElementById('projectNoInput').value = '';
    document.getElementById('locationInput').value = '';
    document.getElementById('reasonInput').value = '';
    await loadUserOvertimes();
  } catch (err) {
    console.error(err);
    showToast(`통신 오류가 발생했습니다: ${err.message || '네트워크를 확인해주세요.'}`, 'error');
  } finally {
    restoreBtn();
  }
});

// ===== 5. 본인 특근 내역 조회 및 통계 / 달력 (요구사항 4) =====
document.getElementById('userRefreshBtn').addEventListener('click', loadUserOvertimes);

let userViewMode = 'list'; // 'list' | 'calendar'

// 개인 특근 뷰 전환 탭 바인딩
const myViewTabList = document.getElementById('myViewTabList');
const myViewTabCalendar = document.getElementById('myViewTabCalendar');
const userRecordListView = document.getElementById('userRecordListView');
const userCalendarView = document.getElementById('userCalendarView');

if (myViewTabList && myViewTabCalendar) {
  myViewTabList.addEventListener('click', () => {
    userViewMode = 'list';
    myViewTabList.classList.add('active');
    myViewTabCalendar.classList.remove('active');
    if (userRecordListView) userRecordListView.style.display = 'block';
    if (userCalendarView) userCalendarView.style.display = 'none';
  });

  myViewTabCalendar.addEventListener('click', () => {
    userViewMode = 'calendar';
    myViewTabCalendar.classList.add('active');
    myViewTabList.classList.remove('active');
    if (userRecordListView) userRecordListView.style.display = 'none';
    if (userCalendarView) userCalendarView.style.display = 'block';
    renderUserCalendar(userCalYear, userCalMonth);
  });
}

// 개인 달력 이전/다음 달 이동
const userCalPrevBtn = document.getElementById('userCalPrevBtn');
const userCalNextBtn = document.getElementById('userCalNextBtn');

if (userCalPrevBtn) {
  userCalPrevBtn.addEventListener('click', () => {
    userCalMonth--;
    if (userCalMonth < 0) {
      userCalMonth = 11;
      userCalYear--;
    }
    renderUserCalendar(userCalYear, userCalMonth);
  });
}

if (userCalNextBtn) {
  userCalNextBtn.addEventListener('click', () => {
    userCalMonth++;
    if (userCalMonth > 11) {
      userCalMonth = 0;
      userCalYear++;
    }
    renderUserCalendar(userCalYear, userCalMonth);
  });
}

// ===== 사용자 모드 다차원 검색 및 필터링 (요구사항 5) =====
function updateUserFilterStats(list) {
  const totalCount = list.length;
  let totalDays = 0;
  let confirmedCount = 0;
  let pendingCount = 0;
  let normalCount = 0;
  let legalCount = 0;
  let subCount = 0;

  list.forEach(item => {
    if (item.is_confirmed === 1) confirmedCount++;
    else pendingCount++;

    if (item.category === '일반휴일') normalCount++;
    else if (item.category === '법정휴일') legalCount++;
    else if (item.category === '대체근무') subCount++;

    const s = item.start_date;
    const e = item.end_date || item.start_date;
    if (s && e) {
      const d1 = new Date(s);
      const d2 = new Date(e);
      const diff = Math.max(1, Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1);
      totalDays += diff;
    } else {
      totalDays += 1;
    }
  });

  const setTxt = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  };

  setTxt('userStatTotal', totalCount);
  setTxt('userStatTotalDays', totalDays);
  setTxt('userStatConfirmed', confirmedCount);
  setTxt('userStatPending', pendingCount);
  setTxt('userStatNormal', normalCount);
  setTxt('userStatLegal', legalCount);
  setTxt('userStatSub', subCount);
  setTxt('userFilterMatchCount', totalCount);
}

function getFilteredUserOvertimes() {
  let list = (userOvertimesData || []).slice();
  const sDate = document.getElementById('userFilterStartDate') ? document.getElementById('userFilterStartDate').value : '';
  const eDate = document.getElementById('userFilterEndDate') ? document.getElementById('userFilterEndDate').value : '';
  const cat = document.getElementById('userFilterCategory') ? document.getElementById('userFilterCategory').value : '';
  const st = document.getElementById('userFilterStatus') ? document.getElementById('userFilterStatus').value : '';
  const q = document.getElementById('userFilterSearch') ? document.getElementById('userFilterSearch').value.trim().toLowerCase() : '';

  if (sDate) {
    list = list.filter(item => (item.end_date || item.start_date) >= sDate);
  }
  if (eDate) {
    list = list.filter(item => item.start_date <= eDate);
  }
  if (cat) {
    list = list.filter(item => item.category === cat);
  }
  if (st !== '') {
    const isConf = parseInt(st, 10);
    list = list.filter(item => item.is_confirmed === isConf);
  }
  if (q) {
    list = list.filter(item => {
      const p = (item.project_no || '').toLowerCase();
      const l = (item.location || '').toLowerCase();
      const r = (item.reason || '').toLowerCase();
      return p.includes(q) || l.includes(q) || r.includes(q);
    });
  }

  updateUserFilterStats(list);

  return list;
}

function renderUserRecords() {
  const container = document.getElementById('userRecordList');
  if (!container) return;
  const list = getFilteredUserOvertimes();

  if (list.length === 0) {
    container.innerHTML = `
      <div style="text-align: center; color: var(--text-muted); padding: 2.5rem 1rem;">
        <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">🔍</div>
        조건에 일치하는 특근 내역이 없습니다.<br>검색 조건을 변경해보세요!
      </div>`;
    return;
  }

  container.innerHTML = '';
  list.forEach(item => {
    const card = document.createElement('div');
    card.className = 'record-card';

    const isConf = item.is_confirmed === 1;
    const statusHtml = isConf
      ? `<span class="status-badge confirmed">✓ 확인완료 (${item.confirmed_by || '관리자'})</span>`
      : `<span class="status-badge pending">⏳ 승인대기 (미확인)</span>`;

    let daysCount = 1;
    try {
      const d1 = new Date(item.start_date);
      const d2 = new Date(item.end_date);
      daysCount = Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1;
    } catch (e) {}

    card.innerHTML = `
      <div class="record-main">
        <div class="record-header-line">
          <span class="cat-badge ${item.category}">${item.category}</span>
          ${statusHtml}
          <span style="font-size: 0.76rem; color: var(--text-light); margin-left: auto;">신청: ${item.created_at}</span>
        </div>
        <div class="record-date-range">
          ${item.start_date} ~ ${item.end_date} 
          <span style="font-size: 0.85rem; font-weight: 500; color: var(--primary);">(${daysCount}일간)</span>
        </div>
        <div class="record-details">
          <span>🏷️ 프로젝트: <b>${item.project_no || '미지정'}</b></span>
          <span>📍 장소: <b>${item.location || '미지정'}</b></span>
        </div>
        <div class="record-reason">
          💬 ${escapeHtml(item.reason || '사유 없음')}
        </div>
        ${item.sub_holiday_date ? `
          <div style="margin-top: 0.45rem; padding: 0.35rem 0.65rem; background: #fef3c7; border-radius: 6px; border: 1px solid #fde68a; font-size: 0.82rem; color: #b45309; font-weight: 700; display: inline-flex; align-items: center; gap: 4px;">
            🏖️ 대체휴일 사용: <span>${escapeHtml(item.sub_holiday_date)}</span> (${item.sub_holiday_used}일)
          </div>
        ` : ''}
      </div>
      <div class="record-actions">
        <button class="btn btn-secondary btn-sm" onclick="openHistoryModal(${item.id})">이력</button>
        ${(!isConf || (currentUser && (currentUser.is_admin === 1 || currentUser.is_super === 1))) ? `
          <button class="btn btn-secondary btn-sm" onclick="openUserOvertimeEditModal(${item.id})">수정</button>
          <button class="btn btn-danger btn-sm" onclick="handleDeleteOvertime(${item.id})">삭제</button>
        ` : `
          <span class="badge" style="background: #ecfdf5; color: #065f46; font-size: 0.75rem; padding: 4px 8px; border-radius: 6px; font-weight: 700; border: 1px solid #a7f3d0;" title="승인 완료된 특근은 관리자만 수정 및 삭제할 수 있습니다.">🔒 승인완료 (수정/삭제 불가)</span>
        `}
      </div>
    `;
    container.appendChild(card);
  });
}

async function loadUserOvertimes() {
  if (!currentUser) return;
  const container = document.getElementById('userRecordList');
  if (container) container.innerHTML = '<div style="text-align: center; color: var(--text-muted); padding: 1.5rem;">데이터를 불러오는 중입니다...</div>';

  try {
    const res = await fetch(`/api/overtimes?emp_id=${encodeURIComponent(currentUser.emp_id)}`);
    const data = await res.json();
    if (!res.ok) {
      if (container) container.innerHTML = '<div style="text-align: center; color: var(--danger); padding: 1rem;">조회 실패</div>';
      return;
    }

    userOvertimesData = data.overtimes || [];
    renderUserRecords();
    renderUserCalendar(userCalYear, userCalMonth);
  } catch (err) {
    console.error(err);
    if (container) container.innerHTML = '<div style="text-align: center; color: var(--danger); padding: 1rem;">통신 오류</div>';
  }
}

// 사용자 모드 다차원 검색 및 필터 이벤트 리스너 바인딩 (요구사항 5)
const uFilterStart = document.getElementById('userFilterStartDate');
const uFilterEnd = document.getElementById('userFilterEndDate');
const uFilterCat = document.getElementById('userFilterCategory');
const uFilterStatus = document.getElementById('userFilterStatus');
const uFilterSearch = document.getElementById('userFilterSearch');
const uFilterReset = document.getElementById('userSearchResetBtn');

const onUserFilterChange = () => {
  renderUserRecords();
  renderUserCalendar(userCalYear, userCalMonth);
};

if (uFilterStart) uFilterStart.addEventListener('change', onUserFilterChange);
if (uFilterEnd) uFilterEnd.addEventListener('change', onUserFilterChange);
if (uFilterCat) uFilterCat.addEventListener('change', onUserFilterChange);
if (uFilterStatus) uFilterStatus.addEventListener('change', onUserFilterChange);
if (uFilterSearch) uFilterSearch.addEventListener('input', onUserFilterChange);
if (uFilterReset) {
  uFilterReset.addEventListener('click', () => {
    if (uFilterStart) uFilterStart.value = '';
    if (uFilterEnd) uFilterEnd.value = '';
    if (uFilterCat) uFilterCat.value = '';
    if (uFilterStatus) uFilterStatus.value = '';
    if (uFilterSearch) uFilterSearch.value = '';
    onUserFilterChange();
  });
}

// 요구사항 4: 개인 특근 통계(월간/반기별/연간) 계산 및 렌더링
function renderUserMyStats(list) {
  const now = new Date();
  const curYear = now.getFullYear();
  const curMonth = now.getMonth() + 1; // 1~12

  // 0. 대체휴무 출장기간 수집
  const tripRanges = [];
  list.forEach(item => {
    const cat = (item.category || '').trim();
    const ts = (item.trip_start_date || '').trim();
    const te = (item.trip_end_date || '').trim();
    if ((cat === '대체휴무' || cat === '대체휴일') && ts && te) {
      tripRanges.push({ start: ts, end: te });
    }
  });

  let monthTotal = 0, monthConfirmed = 0;
  let monthNorm = 0, monthPre = 0, monthSub = 0, monthTripPre = 0;

  let halfTotal = 0, halfConfirmed = 0;
  let halfNorm = 0, halfPre = 0, halfSub = 0, halfTripPre = 0;
  const isFirstHalf = (curMonth <= 6);

  let yearTotal = 0, yearConfirmed = 0;
  let yearNorm = 0, yearPre = 0, yearSub = 0, yearTripPre = 0;

  list.forEach(item => {
    let d1, d2;
    try {
      d1 = new Date(item.start_date);
      d2 = new Date(item.end_date || item.start_date);
      if (isNaN(d1.getTime())) return;
      if (isNaN(d2.getTime()) || d2 < d1) d2 = d1;
    } catch (e) {
      return;
    }

    const itemYear = d1.getFullYear();
    const itemMonth = d1.getMonth() + 1;
    const days = Math.max(1, Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1);
    const subUsed = Number(item.sub_holiday_used || 0);
    const isConf = item.is_confirmed === 1;
    const isPre = item.is_pre_deduct === 1;

    const sStr = item.start_date;
    const eStr = item.end_date || item.start_date;
    let inTrip = false;
    if (isPre) {
      for (const tr of tripRanges) {
        if (!(eStr < tr.start || sStr > tr.end)) {
          inTrip = true;
          break;
        }
      }
    }

    const otDays = (item.category === '일반휴일') ? days : 0;
    const subDays = (item.category === '대체휴무' || item.category === '대체휴일') ? days : subUsed;
    const preCount = isPre ? days : 0;
    const tripPreCount = inTrip ? days : 0;

    // 1. 연간
    if (itemYear === curYear) {
      yearTotal++;
      yearNorm += otDays;
      yearPre += preCount;
      yearSub += subDays;
      yearTripPre += tripPreCount;
      if (isConf) yearConfirmed++;

      // 2. 반기별
      if (isFirstHalf && (itemMonth >= 1 && itemMonth <= 6)) {
        halfTotal++;
        halfNorm += otDays;
        halfPre += preCount;
        halfSub += subDays;
        halfTripPre += tripPreCount;
        if (isConf) halfConfirmed++;
      } else if (!isFirstHalf && (itemMonth >= 7 && itemMonth <= 12)) {
        halfTotal++;
        halfNorm += otDays;
        halfPre += preCount;
        halfSub += subDays;
        halfTripPre += tripPreCount;
        if (isConf) halfConfirmed++;
      }

      // 3. 월간
      if (itemMonth === curMonth) {
        monthTotal++;
        monthNorm += otDays;
        monthPre += preCount;
        monthSub += subDays;
        monthTripPre += tripPreCount;
        if (isConf) monthConfirmed++;
      }
    }
  });

  // 최종 실특근 = 일반특근 - 사전차감 - (대체휴무 - 대체휴무시 작성한 출장기간 이내의 사전차감)
  const calcAct = (norm, pre, sub, tripPre) => Math.max(0, Math.round((norm - pre - (sub - tripPre)) * 10) / 10);
  const monthActDays = calcAct(monthNorm, monthPre, monthSub, monthTripPre);
  const halfActDays = calcAct(halfNorm, halfPre, halfSub, halfTripPre);
  const yearActDays = calcAct(yearNorm, yearPre, yearSub, yearTripPre);

  // UI 엘리먼트 갱신
  const elMonthDays = document.getElementById('myStatMonthDays');
  const elMonthTotal = document.getElementById('myStatMonthTotalCount');
  const elMonthConf = document.getElementById('myStatMonthConfirmed');
  if (elMonthDays) elMonthDays.textContent = monthActDays.toFixed(1).replace(/\.0$/, '');
  if (elMonthTotal) elMonthTotal.textContent = monthTotal;
  if (elMonthConf) elMonthConf.textContent = monthConfirmed;

  const elHalfTitle = document.getElementById('myStatHalfTitle');
  const elHalfBadge = document.getElementById('myStatHalfBadge');
  const elHalfDays = document.getElementById('myStatHalfDays');
  const elHalfCount = document.getElementById('myStatHalfCount');
  const elHalfConf = document.getElementById('myStatHalfConfirmed');
  if (elHalfTitle) elHalfTitle.textContent = isFirstHalf ? '🌓 반기별 누적 (상반기)' : '🌓 반기별 누적 (하반기)';
  if (elHalfBadge) elHalfBadge.textContent = isFirstHalf ? '1~6월' : '7~12월';
  if (elHalfDays) elHalfDays.textContent = halfActDays.toFixed(1).replace(/\.0$/, '');
  if (elHalfCount) elHalfCount.textContent = halfTotal;
  if (elHalfConf) elHalfConf.textContent = halfConfirmed;

  const elYearTitle = document.getElementById('myStatYearTitle');
  const elYearDays = document.getElementById('myStatYearDays');
  const elYearCount = document.getElementById('myStatYearCount');
  const elYearRate = document.getElementById('myStatYearRate');
  if (elYearTitle) elYearTitle.textContent = `🏆 연간 종합 (${curYear}년)`;
  if (elYearDays) elYearDays.textContent = yearActDays.toFixed(1).replace(/\.0$/, '');
  if (elYearCount) elYearCount.textContent = yearTotal;
  const rate = yearTotal > 0 ? Math.round((yearConfirmed / yearTotal) * 100) : 0;
  if (elYearRate) elYearRate.textContent = `${rate}%`;
}

// 요구사항 4: 사용자 전용 월간 달력 렌더링
function renderUserCalendar(year, month) {
  const grid = document.getElementById('userCalendarGrid');
  const title = document.getElementById('userCalMonthTitle');
  if (!grid || !title) return;

  title.textContent = `${year}년 ${month + 1}월`;
  grid.innerHTML = '';

  const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
  dayNames.forEach((d, idx) => {
    const el = document.createElement('div');
    el.className = `cal-day-name ${idx === 0 ? 'sun' : (idx === 6 ? 'sat' : '')}`;
    el.textContent = d;
    grid.appendChild(el);
  });

  const firstDay = new Date(year, month, 1);
  const startDayOfWeek = firstDay.getDay();
  const lastDate = new Date(year, month + 1, 0).getDate();
  const prevLastDate = new Date(year, month, 0).getDate();

  const today = new Date();
  const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;

  // 이전 달 날짜 셀
  for (let i = startDayOfWeek - 1; i >= 0; i--) {
    const cell = document.createElement('div');
    cell.className = 'user-cal-cell other-month';
    cell.innerHTML = `<div class="user-cal-date-num" style="color:var(--text-light);">${prevLastDate - i}</div>`;
    grid.appendChild(cell);
  }

  // 당월 날짜 셀
  for (let date = 1; date <= lastDate; date++) {
    const cell = document.createElement('div');
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(date).padStart(2, '0')}`;
    const dayOfWeek = new Date(year, month, date).getDay();

    const isToday = (dateStr === todayStr);
    cell.className = `user-cal-cell ${isToday ? 'today' : ''}`;
    cell.style.cursor = 'pointer';

    // 캘린더 선택 날짜를 특근 신청 폼의 기본값으로 자동 연동 (요구사항 1)
    cell.addEventListener('click', () => {
      calStartDate = dateStr;
      calEndDate = dateStr;
      const sInput = document.getElementById('startDateInput');
      const eInput = document.getElementById('endDateInput');
      if (sInput) sInput.value = dateStr;
      if (eInput) eInput.value = dateStr;
      renderCalendar();
      showToast(`📅 ${dateStr} 선택됨 (특근 신청일로 자동 지정)`);
    });

    let dateNumColor = 'var(--text-main)';
    if (dayOfWeek === 0) dateNumColor = 'var(--danger)';
    if (dayOfWeek === 6) dateNumColor = 'var(--primary)';

    const numRow = document.createElement('div');
    numRow.className = 'user-cal-date-num';
    numRow.style.color = dateNumColor;
    numRow.innerHTML = `<span>${date}</span> ${isToday ? '<span style="font-size:0.68rem; background:var(--primary); color:#fff; padding:1px 4px; border-radius:3px;">오늘</span>' : ''}`;
    cell.appendChild(numRow);

    // 해당 일자에 걸쳐있는 내 특근 찾기 (필터 적용)
    const dayOvertimes = getFilteredUserOvertimes().filter(item => {
      const s = item.start_date;
      const e = item.end_date || item.start_date;
      return dateStr >= s && dateStr <= e;
    });

    dayOvertimes.forEach(ot => {
      const badge = document.createElement('div');
      const isConf = ot.is_confirmed === 1;
      let badgeClass = isConf ? 'confirmed' : 'pending';
      if (ot.category !== '일반휴일') badgeClass = 'sub-legal';

      badge.className = `user-cal-badge ${badgeClass}`;
      const prefix = isConf ? '✓' : '⏳';
      badge.textContent = `${prefix} [${ot.category}] ${ot.reason || '특근'}`;
      badge.title = `[${ot.category}] ${ot.start_date}~${ot.end_date} | 사유: ${ot.reason || '없음'} | 상태: ${isConf ? '승인완료' : '승인대기'}`;

      badge.addEventListener('click', (e) => {
        e.stopPropagation();
        showToast(`[${ot.category}] ${ot.start_date}~${ot.end_date} | 사유: ${ot.reason || '없음'} (${isConf ? '승인완료' : '승인대기'})`);
      });

      cell.appendChild(badge);
    });

    grid.appendChild(cell);
  }

  // 다음 달 날짜 채우기 (42셀 맞춤)
  const totalCells = startDayOfWeek + lastDate;
  const remaining = (totalCells % 7 === 0) ? 0 : 7 - (totalCells % 7);
  for (let d = 1; d <= remaining; d++) {
    const cell = document.createElement('div');
    cell.className = 'user-cal-cell other-month';
    cell.innerHTML = `<div class="user-cal-date-num" style="color:var(--text-light);">${d}</div>`;
    grid.appendChild(cell);
  }
}

// ===== 6. 수정/삭제/이력 모달 =====
// =========================================================================
// 2-A. 사용자 모드: 특근 내역 수정 (사전차감 / 보너스 일체 미노출)
// =========================================================================
async function openUserOvertimeEditModal(itemId) {
  try {
    const res = await fetch(`/api/overtimes/${itemId}`);
    const data = await res.json();
    if (!res.ok) {
      showToast('상세 내역 조회 실패', 'error');
      return;
    }
    const item = data.overtime;
    const isUserAdmin = currentUser && (currentUser.is_admin === 1 || currentUser.is_super === 1);
    if (item.is_confirmed === 1 && !isUserAdmin && currentMode === 'user') {
      showToast('승인 완료된 특근은 관리자만 수정할 수 있습니다.', 'error');
      return;
    }
    document.getElementById('userEditId').value = item.id;
    document.getElementById('userEditCategory').value = item.category;
    document.getElementById('userEditStartDate').value = item.start_date;
    document.getElementById('userEditEndDate').value = item.end_date;
    document.getElementById('userEditProjectNo').value = item.project_no || '';
    document.getElementById('userEditLocation').value = item.location || '';
    document.getElementById('userEditReason').value = item.reason || '';

    // 대체휴무 출장기간 표시 제어
    const tripRow = document.getElementById('userEditTripDatesRow');
    if (tripRow) {
      tripRow.style.display = item.category === '대체휴무' ? 'grid' : 'none';
    }
    const tripStartEl = document.getElementById('userEditTripStartDate');
    const tripEndEl = document.getElementById('userEditTripEndDate');
    if (tripStartEl) tripStartEl.value = item.trip_start_date || '';
    if (tripEndEl) tripEndEl.value = item.trip_end_date || '';

    openModal('userOvertimeEditModal');
  } catch (err) {
    console.error(err);
    showToast('통신 오류', 'error');
  }
}

const userEditCategoryEl = document.getElementById('userEditCategory');
if (userEditCategoryEl) {
  userEditCategoryEl.addEventListener('change', () => {
    const tripRow = document.getElementById('userEditTripDatesRow');
    if (tripRow) {
      tripRow.style.display = userEditCategoryEl.value === '대체휴무' ? 'grid' : 'none';
    }
  });
}

const userEditOvertimeForm = document.getElementById('userEditOvertimeForm');
if (userEditOvertimeForm) {
  userEditOvertimeForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const id = document.getElementById('userEditId').value;
    const category = document.getElementById('userEditCategory').value;
    const startDate = document.getElementById('userEditStartDate').value;
    const endDate = document.getElementById('userEditEndDate').value;
    const projectNo = document.getElementById('userEditProjectNo').value.trim();
    const location = document.getElementById('userEditLocation').value.trim();
    const reason = document.getElementById('userEditReason').value.trim();

    const tripStartDate = (document.getElementById('userEditTripStartDate')?.value || '').trim();
    const tripEndDate = (document.getElementById('userEditTripEndDate')?.value || '').trim();

    if (startDate > endDate) {
      showToast('종료일은 시작일보다 빠를 수 없습니다.', 'error');
      return;
    }

    // 사용자 모드는 사전차감/보너스 정보를 수정하거나 전송하지 않음
    const payload = {
      changed_by: currentUser ? currentUser.emp_id : '',
      category,
      start_date: startDate,
      end_date: endDate,
      project_no: projectNo,
      location,
      reason,
      trip_start_date: tripStartDate,
      trip_end_date: tripEndDate
    };

    const submitBtn = userEditOvertimeForm.querySelector('button[type="submit"]');
    const restoreBtn = setButtonLoading(submitBtn, '수정 저장 중...');

    try {
      const res = await fetch(`/api/overtimes/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || '수정 실패', 'error');
        return;
      }
      closeModal('userOvertimeEditModal');
      showToast('특근/휴무 내역이 수정되었습니다.');
      loadUserOvertimes();
    } catch (err) {
      console.error(err);
      showToast('수정 중 오류 발생', 'error');
    } finally {
      restoreBtn();
    }
  });
}

// =========================================================================
// 2-B. 관리자 모드: 팀원 특근 내역 수정 (사전차감 및 보너스 부여 토글 지원)
// =========================================================================
function setAdminEditPreDeduct(active) {
  const btn = document.getElementById('adminEditPreDeductToggleBtn');
  const input = document.getElementById('adminEditIsPreDeduct');
  if (!btn || !input) return;
  input.value = active ? '1' : '0';
  if (active) {
    btn.textContent = '적용 ON';
    btn.className = 'btn btn-sm btn-primary';
    btn.style.background = 'linear-gradient(135deg, #0284c7 0%, #0369a1 100%)';
    btn.style.color = '#fff';
    btn.style.borderColor = '#0284c7';
  } else {
    btn.textContent = 'OFF';
    btn.className = 'btn btn-sm btn-secondary';
    btn.style.background = '';
    btn.style.color = '';
    btn.style.borderColor = '';
  }
}

function setAdminEditBonus(active) {
  const btn = document.getElementById('adminEditBonusToggleBtn');
  const input = document.getElementById('adminEditBonusGranted');
  if (!btn || !input) return;
  input.value = active ? '1' : '0';
  if (active) {
    btn.textContent = '🎁 부여 ON';
    btn.className = 'btn btn-sm btn-primary';
    btn.style.background = 'linear-gradient(135deg, #9333ea 0%, #7e22ce 100%)';
    btn.style.color = '#fff';
    btn.style.borderColor = '#7e22ce';
  } else {
    btn.textContent = '미부여 OFF';
    btn.className = 'btn btn-sm btn-secondary';
    btn.style.background = '';
    btn.style.color = '';
    btn.style.borderColor = '';
  }
}

const adminEditPreDeductToggleBtn = document.getElementById('adminEditPreDeductToggleBtn');
if (adminEditPreDeductToggleBtn) {
  adminEditPreDeductToggleBtn.addEventListener('click', () => {
    const currentVal = document.getElementById('adminEditIsPreDeduct')?.value === '1';
    setAdminEditPreDeduct(!currentVal);
  });
}

const adminEditBonusToggleBtn = document.getElementById('adminEditBonusToggleBtn');
if (adminEditBonusToggleBtn) {
  adminEditBonusToggleBtn.addEventListener('click', () => {
    const currentVal = document.getElementById('adminEditBonusGranted')?.value === '1';
    setAdminEditBonus(!currentVal);
  });
}

const adminEditCategoryEl = document.getElementById('adminEditCategory');
if (adminEditCategoryEl) {
  adminEditCategoryEl.addEventListener('change', () => {
    const tripRow = document.getElementById('adminEditTripDatesRow');
    if (tripRow) {
      tripRow.style.display = adminEditCategoryEl.value === '대체휴무' ? 'grid' : 'none';
    }
  });
}

async function openAdminEditModal(itemId) {
  try {
    const viewerParam = currentUser ? `?viewer_emp_id=${encodeURIComponent(currentUser.emp_id)}` : '';
    const res = await fetch(`/api/overtimes/${itemId}${viewerParam}`);
    const data = await res.json();
    if (!res.ok) {
      showToast('상세 내역 조회 실패', 'error');
      return;
    }
    const item = data.overtime;
    document.getElementById('adminEditId').value = item.id;
    document.getElementById('adminEditCategory').value = item.category;
    document.getElementById('adminEditStartDate').value = item.start_date;
    document.getElementById('adminEditEndDate').value = item.end_date;
    document.getElementById('adminEditProjectNo').value = item.project_no || '';
    document.getElementById('adminEditLocation').value = item.location || '';
    document.getElementById('adminEditReason').value = item.reason || '';

    // 대체휴무 출장기간 표시 제어
    const tripRow = document.getElementById('adminEditTripDatesRow');
    if (tripRow) {
      tripRow.style.display = item.category === '대체휴무' ? 'grid' : 'none';
    }
    const tripStartEl = document.getElementById('adminEditTripStartDate');
    const tripEndEl = document.getElementById('adminEditTripEndDate');
    if (tripStartEl) tripStartEl.value = item.trip_start_date || '';
    if (tripEndEl) tripEndEl.value = item.trip_end_date || '';

    // 사전차감 토글 상태 세팅
    setAdminEditPreDeduct(item.is_pre_deduct === 1);
    // 보너스 부여 토글 상태 세팅
    setAdminEditBonus(item.bonus_granted === 1);

    openModal('adminEditModal');
  } catch (err) {
    console.error(err);
    showToast('통신 오류', 'error');
  }
}

const adminEditOvertimeForm = document.getElementById('adminEditOvertimeForm');
if (adminEditOvertimeForm) {
  adminEditOvertimeForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const id = document.getElementById('adminEditId').value;
    const category = document.getElementById('adminEditCategory').value;
    const startDate = document.getElementById('adminEditStartDate').value;
    const endDate = document.getElementById('adminEditEndDate').value;
    const projectNo = document.getElementById('adminEditProjectNo').value.trim();
    const location = document.getElementById('adminEditLocation').value.trim();
    const reason = document.getElementById('adminEditReason').value.trim();

    const tripStartDate = (document.getElementById('adminEditTripStartDate')?.value || '').trim();
    const tripEndDate = (document.getElementById('adminEditTripEndDate')?.value || '').trim();
    const isPreDeduct = document.getElementById('adminEditIsPreDeduct')?.value === '1' ? 1 : 0;
    const bonusGranted = document.getElementById('adminEditBonusGranted')?.value === '1' ? 1 : 0;

    if (startDate > endDate) {
      showToast('종료일은 시작일보다 빠를 수 없습니다.', 'error');
      return;
    }

    const payload = {
      changed_by: currentUser ? currentUser.emp_id : '',
      category,
      start_date: startDate,
      end_date: endDate,
      project_no: projectNo,
      location,
      reason,
      is_pre_deduct: isPreDeduct,
      bonus_granted: bonusGranted,
      trip_start_date: tripStartDate,
      trip_end_date: tripEndDate
    };

    const submitBtn = adminEditOvertimeForm.querySelector('button[type="submit"]');
    const restoreBtn = setButtonLoading(submitBtn, '수정 저장 중...');

    try {
      const res = await fetch(`/api/overtimes/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || '수정 실패', 'error');
        return;
      }
      closeModal('adminEditModal');
      showToast('특근/휴무 내역이 수정되었습니다.');
      loadAdminData();
    } catch (err) {
      console.error(err);
      showToast('수정 중 오류 발생', 'error');
    } finally {
      restoreBtn();
    }
  });
}

// 통합 라우터 (호환성 유지)
async function openEditModal(itemId) {
  if (currentMode === 'admin' || (currentUser && (currentUser.is_admin === 1 || currentUser.is_super === 1))) {
    return openAdminEditModal(itemId);
  } else {
    return openUserOvertimeEditModal(itemId);
  }
}
window.openUserOvertimeEditModal = openUserOvertimeEditModal;
window.openAdminEditModal = openAdminEditModal;
window.openEditModal = openEditModal;

async function handleDeleteOvertime(itemId) {
  const isUserAdmin = currentUser && (currentUser.is_admin === 1 || currentUser.is_super === 1);
  if (currentMode === 'user' && !isUserAdmin) {
    const cachedItem = userOvertimesData.find(x => x.id === itemId);
    if (cachedItem && cachedItem.is_confirmed === 1) {
      showToast('승인 완료된 특근은 관리자만 삭제할 수 있습니다.', 'error');
      return;
    }
  }
  if (!confirm('정말 이 특근 내역을 삭제하시겠습니까? (삭제 이력은 보존됩니다)')) return;

  try {
    const res = await fetch(`/api/overtimes/${itemId}?changed_by=${encodeURIComponent(currentUser.emp_id)}`, {
      method: 'DELETE'
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '삭제 실패', 'error');
      return;
    }
    showToast('특근 내역이 삭제되었습니다.');
    if (currentMode === 'user') loadUserOvertimes();
    else loadAdminData();
  } catch (err) {
    console.error(err);
    showToast('삭제 중 오류 발생', 'error');
  }
}

async function openHistoryModal(itemId) {
  try {
    const viewerParam = currentUser ? `?viewer_emp_id=${encodeURIComponent(currentUser.emp_id)}` : '';
    const res = await fetch(`/api/overtimes/${itemId}${viewerParam}`);
    const data = await res.json();
    if (!res.ok) {
      showToast('이력 조회 실패', 'error');
      return;
    }

    const historyTimeline = document.getElementById('historyTimeline');
    historyTimeline.innerHTML = '';
    const histories = data.history || [];

    if (histories.length === 0) {
      historyTimeline.innerHTML = '<div style="color: var(--text-muted);">기록된 변경 이력이 없습니다.</div>';
    } else {
      histories.forEach(h => {
        const item = document.createElement('div');
        item.className = 'history-item';
        item.innerHTML = `
          <div class="history-dot"></div>
          <div class="history-time">${h.created_at}</div>
          <div class="history-action">[${h.action}] 작업자: ${h.changed_by_name || h.changed_by} (${h.changed_by})</div>
        `;
        historyTimeline.appendChild(item);
      });
    }
    openModal('historyModal');
  } catch (err) {
    console.error(err);
    showToast('통신 오류', 'error');
  }
}

// ===== 7. 관리자 모드: 대시보드 & 부서별 필터 & 뷰 전환 =====
const adminSubTabOvertimes = document.getElementById('adminSubTabOvertimes');
const adminSubTabUsers = document.getElementById('adminSubTabUsers');
const adminSubTabSummary = document.getElementById('adminSubTabSummary');
const adminSubTabAccessLogs = document.getElementById('adminSubTabAccessLogs');
const adminOvertimeView = document.getElementById('adminOvertimeView');
const adminUserView = document.getElementById('adminUserView');
const adminSummaryView = document.getElementById('adminSummaryView');
const adminAccessLogsView = document.getElementById('adminAccessLogsView');

const viewTabCalendar = document.getElementById('viewTabCalendar');
const viewTabTable = document.getElementById('viewTabTable');
const adminCalendarView = document.getElementById('adminCalendarView');
const adminTableView = document.getElementById('adminTableView');

adminSubTabOvertimes.addEventListener('click', () => {
  adminSubTabOvertimes.className = 'btn btn-primary btn-sm';
  adminSubTabUsers.className = 'btn btn-secondary btn-sm';
  if (adminSubTabSummary) adminSubTabSummary.className = 'btn btn-secondary btn-sm';
  if (adminSubTabAccessLogs) adminSubTabAccessLogs.className = 'btn btn-secondary btn-sm';
  adminOvertimeView.style.display = 'block';
  adminUserView.style.display = 'none';
  if (adminSummaryView) adminSummaryView.style.display = 'none';
  if (adminAccessLogsView) adminAccessLogsView.style.display = 'none';
  loadAdminData();
});

adminSubTabUsers.addEventListener('click', () => {
  adminSubTabUsers.className = 'btn btn-primary btn-sm';
  adminSubTabOvertimes.className = 'btn btn-secondary btn-sm';
  if (adminSubTabSummary) adminSubTabSummary.className = 'btn btn-secondary btn-sm';
  if (adminSubTabAccessLogs) adminSubTabAccessLogs.className = 'btn btn-secondary btn-sm';
  adminOvertimeView.style.display = 'none';
  adminUserView.style.display = 'block';
  if (adminSummaryView) adminSummaryView.style.display = 'none';
  if (adminAccessLogsView) adminAccessLogsView.style.display = 'none';
  loadAdminUserTable();
});

if (adminSubTabSummary) {
  adminSubTabSummary.addEventListener('click', () => {
    adminSubTabSummary.className = 'btn btn-primary btn-sm';
    adminSubTabOvertimes.className = 'btn btn-secondary btn-sm';
    adminSubTabUsers.className = 'btn btn-secondary btn-sm';
    if (adminSubTabAccessLogs) adminSubTabAccessLogs.className = 'btn btn-secondary btn-sm';
    adminOvertimeView.style.display = 'none';
    adminUserView.style.display = 'none';
    adminSummaryView.style.display = 'block';
    if (adminAccessLogsView) adminAccessLogsView.style.display = 'none';
    initSummaryDateFilter();
    loadSettlementSummary();
  });
}

if (adminSubTabAccessLogs) {
  adminSubTabAccessLogs.addEventListener('click', () => {
    adminSubTabAccessLogs.className = 'btn btn-primary btn-sm';
    adminSubTabOvertimes.className = 'btn btn-secondary btn-sm';
    adminSubTabUsers.className = 'btn btn-secondary btn-sm';
    if (adminSubTabSummary) adminSubTabSummary.className = 'btn btn-secondary btn-sm';
    adminOvertimeView.style.display = 'none';
    adminUserView.style.display = 'none';
    if (adminSummaryView) adminSummaryView.style.display = 'none';
    if (adminAccessLogsView) adminAccessLogsView.style.display = 'block';
    initAccessLogsFilter();
    loadAccessLogs();
  });
}

viewTabCalendar.addEventListener('click', () => {
  adminViewMode = 'calendar';
  viewTabCalendar.classList.add('active');
  viewTabTable.classList.remove('active');
  adminCalendarView.style.display = 'block';
  adminTableView.style.display = 'none';
  renderAdminCalendar();
});

viewTabTable.addEventListener('click', () => {
  adminViewMode = 'table';
  viewTabTable.classList.add('active');
  viewTabCalendar.classList.remove('active');
  adminCalendarView.style.display = 'none';
  adminTableView.style.display = 'block';
  renderAdminOvertimeTable();
});

function loadAdminDashboard() {
  loadAdminData();
}

async function loadAdminData() {
  try {
    // 1. 팀원 목록 로드 (관리자 등급에 따른 필터링)
    const adminParam = currentUser ? `?admin_emp_id=${encodeURIComponent(currentUser.emp_id)}` : '';
    const userRes = await fetch('/api/users' + adminParam);
    const userData = await userRes.json();
    if (userRes.ok) {
      allUsersCache = userData.users || [];
      document.getElementById('statUserCount').textContent = allUsersCache.length;
    }

    // 2. 특근 목록 로드
    let url = '/api/overtimes';
    const params = [];
    const startDate = document.getElementById('filterStartDate').value;
    const endDate = document.getElementById('filterEndDate').value;
    const team = selectedDeptFilter || document.getElementById('filterTeam').value;
    const category = document.getElementById('filterCategory').value;
    const status = document.getElementById('filterStatus').value;
    const search = document.getElementById('filterSearch').value.trim();

    if (startDate) params.push(`start_date=${encodeURIComponent(startDate)}`);
    if (endDate) params.push(`end_date=${encodeURIComponent(endDate)}`);
    if (team) params.push(`team=${encodeURIComponent(team)}`);
    if (category) params.push(`category=${encodeURIComponent(category)}`);
    if (status !== '') params.push(`is_confirmed=${encodeURIComponent(status)}`);
    if (search) params.push(`search=${encodeURIComponent(search)}`);
    if (currentUser) params.push(`admin_emp_id=${encodeURIComponent(currentUser.emp_id)}`);
    if (params.length > 0) url += '?' + params.join('&');

    const otRes = await fetch(url);
    if (otRes.ok) {
      const otData = await otRes.json();
      let rawList = otData.overtimes || [];
      // 요구사항 2: 팀관리자는 슈퍼관리자의 특근 일정을 열람할 수 없음
      if (currentUser && currentUser.is_super !== 1 && currentUser.emp_id.toLowerCase() !== 'ps37082') {
        rawList = rawList.filter(o => o.emp_id.toLowerCase() !== 'ps37082');
      }
      lastFetchedOvertimes = rawList;
      updateAdminStats(lastFetchedOvertimes);
      renderDeptFilterPills();
      if (adminViewMode === 'calendar') {
        renderAdminCalendar();
      } else {
        renderAdminOvertimeTable();
      }
    }
  } catch (err) {
    console.error(err);
  }
}

function updateAdminStats(list) {
  const total = list.length;
  const conf = list.filter(i => i.is_confirmed === 1).length;
  const pend = total - conf;

  document.getElementById('statTotalCount').textContent = total;
  document.getElementById('statConfirmedCount').textContent = conf;
  document.getElementById('statPendingCount').textContent = pend;

  // 요구사항 2: 팀관리자는 본인 소속 1개의 상단 요약 카드만 표시되고, 슈퍼관리자는 팀별로 구분 표시
  const isSuper = currentUser && (currentUser.is_super === 1 || currentUser.emp_id.toLowerCase() === 'ps37082');
  const superTeamSec = document.getElementById('superAdminTeamSection');
  const superTeamGrid = document.getElementById('superAdminTeamGrid');

  if (isSuper && superTeamSec && superTeamGrid) {
    superTeamSec.style.display = 'block';
    renderSuperAdminTeamCards(list);
  } else if (superTeamSec) {
    superTeamSec.style.display = 'none';
  }
}

// 요구사항 2: 슈퍼관리자 전용 팀별 구분 현황 카드 그리드 렌더링
function renderSuperAdminTeamCards(list) {
  const grid = document.getElementById('superAdminTeamGrid');
  if (!grid) return;
  grid.innerHTML = '';

  // 1. 전체 부서별 집계 데이터 생성
  const teamStats = {};
  
  // 팀원 수 집계
  allUsersCache.forEach(u => {
    const t = u.team || '기타';
    if (!teamStats[t]) {
      teamStats[t] = { team: t, users: 0, total: 0, conf: 0, pend: 0 };
    }
    teamStats[t].users++;
  });

  // 특근 건수 집계
  list.forEach(item => {
    const t = item.team || '기타';
    if (!teamStats[t]) {
      teamStats[t] = { team: t, users: 0, total: 0, conf: 0, pend: 0 };
    }
    teamStats[t].total++;
    if (item.is_confirmed === 1) {
      teamStats[t].conf++;
    } else {
      teamStats[t].pend++;
    }
  });

  const teams = Object.keys(teamStats).sort();
  if (teams.length === 0) {
    grid.innerHTML = '<div style="color:var(--text-muted); font-size:0.85rem; padding:0.5rem;">등록된 팀 정보가 없습니다.</div>';
    return;
  }

  teams.forEach(tName => {
    const s = teamStats[tName];
    const isSelected = (selectedDeptFilter === tName);
    const card = document.createElement('div');
    card.className = `super-team-card ${isSelected ? 'active-team' : ''}`;
    
    card.innerHTML = `
      <div class="team-title">
        <span>🏢 ${escapeHtml(tName)}</span>
        <span class="badge" style="font-size:0.7rem; background:#e0e7ff; color:#3730a3;">👥 ${s.users}명</span>
      </div>
      <div class="team-stats-row">
        <span>신청 건수:</span>
        <b style="color:var(--text-main);">${s.total}건</b>
      </div>
      <div class="team-stats-row">
        <span>승인 완료:</span>
        <b style="color:var(--success);">${s.conf}건</b>
      </div>
      <div class="team-stats-row">
        <span>승인 대기:</span>
        <b style="color:${s.pend > 0 ? '#d97706' : 'var(--text-muted)'};">${s.pend}건</b>
      </div>
    `;

    card.addEventListener('click', () => {
      if (selectedDeptFilter === tName) {
        selectedDeptFilter = ''; // 토글 해제
      } else {
        selectedDeptFilter = tName;
      }
      loadAdminData();
    });

    grid.appendChild(card);
  });
}

// ===== 8. 요구사항 15: 부서별 빠른 칩 필터 렌더링 =====
function renderDeptFilterPills() {
  const container = document.getElementById('deptPillBar');
  const sel = document.getElementById('filterTeam');
  if (!container) return;

  // 부서별 통계 집계
  const deptCounts = {};
  allUsersCache.forEach(u => {
    if (u.team) {
      if (!deptCounts[u.team]) deptCounts[u.team] = { users: 0, overtimes: 0 };
      deptCounts[u.team].users++;
    }
  });

  lastFetchedOvertimes.forEach(o => {
    if (o.team) {
      if (!deptCounts[o.team]) deptCounts[o.team] = { users: 0, overtimes: 0 };
      deptCounts[o.team].overtimes++;
    }
  });

  const teams = Object.keys(deptCounts);

  // 셀렉트 박스 동기화
  const curSelVal = sel.value;
  sel.innerHTML = '<option value="">전체 소속팀</option>';
  teams.forEach(t => {
    const opt = document.createElement('option');
    opt.value = t;
    opt.textContent = t;
    sel.appendChild(opt);
  });
  sel.value = curSelVal;

  // 칩 렌더링
  container.innerHTML = '';

  // 1) 전체 칩
  const allPill = document.createElement('div');
  allPill.className = `dept-pill ${selectedDeptFilter === '' ? 'active' : ''}`;
  allPill.innerHTML = `전체 부서 <span class="count-badge">${lastFetchedOvertimes.length}건</span>`;
  allPill.addEventListener('click', () => {
    selectedDeptFilter = '';
    sel.value = '';
    loadAdminData();
  });
  container.appendChild(allPill);

  // 2) 부서별 칩
  teams.forEach(t => {
    const p = document.createElement('div');
    const isAct = selectedDeptFilter === t;
    p.className = `dept-pill ${isAct ? 'active' : ''}`;
    const info = deptCounts[t];
    p.innerHTML = `${t} <span class="count-badge">${info.overtimes}건 / ${info.users}명</span>`;
    p.addEventListener('click', () => {
      selectedDeptFilter = isAct ? '' : t;
      sel.value = selectedDeptFilter;
      loadAdminData();
    });
    container.appendChild(p);
  });
}

// ===== 9. 요구사항 14: 관리자 월간 캘린더 & 날짜 클릭 작업자 상세 표시 =====
const adminCalPrevBtn = document.getElementById('adminCalPrevBtn');
const adminCalNextBtn = document.getElementById('adminCalNextBtn');
const adminCalMonthTitle = document.getElementById('adminCalMonthTitle');
const adminCalGrid = document.getElementById('adminCalGrid');

adminCalPrevBtn.addEventListener('click', () => {
  adminCalMonth--;
  if (adminCalMonth < 0) { adminCalMonth = 11; adminCalYear--; }
  renderAdminCalendar();
});

adminCalNextBtn.addEventListener('click', () => {
  adminCalMonth++;
  if (adminCalMonth > 11) { adminCalMonth = 0; adminCalYear++; }
  renderAdminCalendar();
});

function renderAdminCalendar() {
  adminCalMonthTitle.textContent = `${adminCalYear}년 ${adminCalMonth + 1}월`;
  adminCalGrid.innerHTML = '';

  const firstDayIndex = new Date(adminCalYear, adminCalMonth, 1).getDay();
  const lastDate = new Date(adminCalYear, adminCalMonth + 1, 0).getDate();
  const prevLastDate = new Date(adminCalYear, adminCalMonth, 0).getDate();
  const todayStr = getTodayStr();

  // 이전 달 날짜들
  for (let x = firstDayIndex; x > 0; x--) {
    const cell = document.createElement('div');
    cell.className = 'admin-cal-cell disabled';
    cell.innerHTML = `<span class="admin-cal-date-num">${prevLastDate - x + 1}</span>`;
    adminCalGrid.appendChild(cell);
  }

  // 이번 달 날짜들
  for (let d = 1; d <= lastDate; d++) {
    const dateStr = `${adminCalYear}-${String(adminCalMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dateObj = new Date(adminCalYear, adminCalMonth, d);
    const dayOfWeek = dateObj.getDay();

    const cell = document.createElement('div');
    cell.className = 'admin-cal-cell';
    if (adminSelectedDate === dateStr) cell.classList.add('selected');

    let dateNumClass = '';
    if (dayOfWeek === 0) dateNumClass = 'sun';
    if (dayOfWeek === 6) dateNumClass = 'sat';

    // 해당 날짜에 해당하는 특근 목록 조회
    const dayOvertimes = lastFetchedOvertimes.filter(o => o.start_date <= dateStr && o.end_date >= dateStr);
    const daySubHolidays = lastFetchedOvertimes.filter(o => o.sub_holiday_date === dateStr && (parseFloat(o.sub_holiday_used) || 0) > 0);
    const workerCount = dayOvertimes.length;
    const subHolidayCount = daySubHolidays.length;
    const bonusCount = dayOvertimes.filter(o => o.bonus_granted === 1).length;

    // 요구사항 4: 축약(+N명) 없이 모든 특근 인원 전체 표시
    let workerChipsHtml = '';
    if (workerCount > 0) {
      dayOvertimes.forEach(o => {
        const hasBonus = o.bonus_granted === 1;
        const hasPre = o.is_pre_deduct === 1;
        const bonusStyle = hasBonus ? 'border: 1.5px solid #8b5cf6; box-shadow: 0 0 0 1px #a78bfa; font-weight: 700;' : '';
        const bonusPrefix = hasBonus ? '<span style="font-size:0.75rem; margin-right:1px;" title="보너스 부여">🎁</span>' : '';
        const prePrefix = hasPre ? '<span style="font-size:0.72rem; margin-right:1px;" title="사전차감">⚡</span>' : '';
        const bonusTitle = hasBonus ? ' [🎁보너스 부여]' : '';
        const preTitle = hasPre ? ' [⚡사전차감]' : '';
        workerChipsHtml += `<div class="worker-chip ${o.category}${hasBonus ? ' has-bonus' : ''}" style="${bonusStyle}" title="${escapeHtml(o.user_name)} (${o.category})${bonusTitle}${preTitle}">${bonusPrefix}${prePrefix}${escapeHtml(o.user_name)}</div>`;
      });
    }
    // 요구사항 3 & 4: 대체휴무(대휴) 사용 인원도 고유 클래스로 눈에 띄게 전체 표시
    if (subHolidayCount > 0) {
      daySubHolidays.forEach(s => {
        workerChipsHtml += `<div class="worker-chip 대체휴무" title="대체휴무 사용: ${escapeHtml(s.user_name)} (${s.sub_holiday_used}일)">🌿 ${escapeHtml(s.user_name)}</div>`;
      });
    }

    cell.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center;">
        <span class="admin-cal-date-num ${dateNumClass}">${d}</span>
        <div style="display:flex; gap:3px; align-items:center;">
          ${workerCount > 0 ? `<span style="font-size:0.7rem; font-weight:700; color:var(--primary);">${workerCount}명</span>` : ''}
          ${bonusCount > 0 ? `<span style="font-size:0.68rem; font-weight:800; color:#6d28d9; background:#ede9fe; padding:1px 4px; border-radius:4px; border:1px solid #c4b5fd;" title="보너스 부여 ${bonusCount}명">🎁${bonusCount}</span>` : ''}
          ${subHolidayCount > 0 ? `<span style="font-size:0.68rem; font-weight:800; color:#065f46; background:#d1fae5; padding:1px 4px; border-radius:4px; border:1px solid #a7f3d0;" title="대체휴무 사용 ${subHolidayCount}명">🌿${subHolidayCount}</span>` : ''}
        </div>
      </div>
      <div class="worker-chips-container">
        ${workerChipsHtml}
      </div>
    `;

    // 날짜 클릭 이벤트 -> 해당일 작업자 상세 표시
    cell.addEventListener('click', () => {
      adminSelectedDate = dateStr;
      renderAdminCalendar(); // 선택 하이라이트 갱신
      showDailyWorkers(dateStr, dayOvertimes);
    });

    adminCalGrid.appendChild(cell);
  }

  // 기존에 선택된 날짜가 있었다면 패널 자동 갱신
  if (adminSelectedDate) {
    const dayOvertimes = lastFetchedOvertimes.filter(o => o.start_date <= adminSelectedDate && o.end_date >= adminSelectedDate);
    showDailyWorkers(adminSelectedDate, dayOvertimes);
  }
}

function showDailyWorkers(dateStr, list) {
  const panel = document.getElementById('dailyWorkersPanel');
  const title = document.getElementById('dailyPanelTitle');
  const countText = document.getElementById('dailyWorkerCountText');
  const tbody = document.getElementById('dailyWorkersTbody');

  // 요구사항 26: 해당 일자에 대체휴일을 사용하는 인원 표시
  const subHolidayUsers = lastFetchedOvertimes.filter(o => o.sub_holiday_date === dateStr && (parseFloat(o.sub_holiday_used) || 0) > 0);
  const subHolNotice = document.getElementById('dailySubHolidayNotice');
  if (subHolNotice) {
    if (subHolidayUsers.length > 0) {
      subHolNotice.style.display = 'flex';
      subHolNotice.innerHTML = `
        <span style="font-size:1.3rem;">🏖️</span>
        <div>
          <b style="color:#b45309;">대체휴일 사용일 인원 (${subHolidayUsers.length}명):</b>
          ${subHolidayUsers.map(u => `<span style="display:inline-block; background:#fff; border:1px solid #f59e0b; padding:2px 8px; border-radius:9999px; margin:2px 4px; font-weight:700; font-size:0.8rem; color:#b45309;">${escapeHtml(u.user_name)} (${escapeHtml(u.team)}, ${u.sub_holiday_used}일 대휴)</span>`).join('')}
        </div>
      `;
    } else {
      subHolNotice.style.display = 'none';
    }
  }

  panel.style.display = 'block';
  title.textContent = `📅 ${dateStr} 특근 작업자 상세 현황`;
  countText.textContent = `특근 ${list.length}명 ${subHolidayUsers.length > 0 ? `| 대체휴일 ${subHolidayUsers.length}명` : ''}`;
  tbody.innerHTML = '';

  if (list.length === 0) {
    tbody.innerHTML = `<tr><td colspan="12" style="text-align:center; padding:1.5rem; color:var(--text-muted);">${dateStr}에 신청된 특근자가 없습니다.</td></tr>`;
    return;
  }

  list.forEach(item => {
    const tr = document.createElement('tr');
    const isConf = item.is_confirmed === 1;

    tr.innerHTML = `
      <td>
        <a href="javascript:void(0)" onclick="showUserStatsPopup('${item.emp_id}', '${escapeHtml(item.user_name)}')" style="color:var(--primary); font-weight:700; text-decoration:underline; cursor:pointer;" title="클릭 시 ${escapeHtml(item.user_name)} 님의 상세 통계 팝업 열기">
          ${escapeHtml(item.user_name)}
        </a>
        <span style="font-size:0.76rem; color:var(--text-light);">(${item.emp_id})</span>
      </td>
      <td>${item.team}</td>
      <td>
        <span class="cat-badge ${item.category}">${item.category}</span>
        ${item.is_pre_deduct === 1 ? `<span class="badge" style="background:#0284c7; color:#fff; font-size:0.72rem; padding:1px 5px; border-radius:4px; font-weight:700; margin-left:3px;">⚡사전차감</span>` : ''}
      </td>
      <td>${item.start_date} ~ ${item.end_date}</td>
      <td>${item.project_no || '-'}</td>
      <td>${item.location || '-'}</td>
      <td style="max-width:200px; word-break:break-all;">${escapeHtml(item.reason || '-')}</td>
      <td style="text-align:center;">
        ${item.bonus_granted === 1 ? '<span class="badge" style="background:#8b5cf6; color:white; font-size:0.75rem; padding:2px 8px; border-radius:9999px; font-weight:700;">🎁 부여</span>' : '<span style="color:#94a3b8; font-size:0.8rem;">-</span>'}
      </td>
      <td style="text-align:center;">
        <label class="toggle-switch" title="사전차감 원클릭 전환">
          <input type="checkbox" class="pre-deduct-toggle" data-id="${item.id}" ${item.is_pre_deduct === 1 ? 'checked' : ''}>
          <span class="toggle-slider"></span>
        </label>
      </td>
      <td>
        <label class="toggle-switch" title="확인(승인) 원클릭 전환">
          <input type="checkbox" class="confirm-toggle" data-id="${item.id}" ${isConf ? 'checked' : ''}>
          <span class="toggle-slider"></span>
        </label>
      </td>
      <td style="font-size:0.8rem; color:var(--text-muted);">${item.confirmed_by || '-'}</td>
      <td>
        <div style="display:flex; gap:0.25rem;">
          <button class="btn btn-secondary btn-sm" onclick="openHistoryModal(${item.id})">이력</button>
          <button class="btn btn-secondary btn-sm" onclick="openAdminEditModal(${item.id})">수정</button>
          <button class="btn btn-danger btn-sm" onclick="handleDeleteOvertime(${item.id})">삭제</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll('.confirm-toggle').forEach(tg => {
    tg.addEventListener('change', async () => {
      const id = parseInt(tg.getAttribute('data-id'), 10);
      const isConfirmed = tg.checked ? 1 : 0;
      await toggleConfirmOvertime(id, isConfirmed);
    });
  });

  tbody.querySelectorAll('.pre-deduct-toggle').forEach(tg => {
    tg.addEventListener('change', async () => {
      const id = parseInt(tg.getAttribute('data-id'), 10);
      await togglePreDeductOvertime(id, tg.checked);
    });
  });

  const dailyBatchConfirmBtn = document.getElementById('dailyBatchConfirmBtn');
  const dailyBatchUnconfirmBtn = document.getElementById('dailyBatchUnconfirmBtn');
  if (dailyBatchConfirmBtn) {
    dailyBatchConfirmBtn.onclick = () => handleDailyBatch(list, 1, dailyBatchConfirmBtn);
  }
  if (dailyBatchUnconfirmBtn) {
    dailyBatchUnconfirmBtn.onclick = () => handleDailyBatch(list, 0, dailyBatchUnconfirmBtn);
  }

  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

async function handleDailyBatch(list, isConfirmed, triggerBtn = null) {
  if (!list || list.length === 0) {
    showToast('처리할 작업자가 없습니다.', 'error');
    return;
  }
  const actionName = isConfirmed === 1 ? '일괄 확인(승인)' : '일괄 확인 취소';
  if (!confirm(`당일(${adminSelectedDate}) 특근자 ${list.length}명을 모두 ${actionName} 처리하시겠습니까?`)) return;

  const btn = triggerBtn || (isConfirmed === 1 ? document.getElementById('dailyBatchConfirmBtn') : document.getElementById('dailyBatchUnconfirmBtn'));
  const restoreBtn = setButtonLoading(btn, isConfirmed === 1 ? '일괄 승인 중...' : '일괄 취소 중...');

  const ids = list.map(item => item.id);
  try {
    const res = await fetch('/api/overtimes/batch-confirm', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ids: ids,
        admin_emp_id: currentUser ? currentUser.emp_id : '',
        is_confirmed: isConfirmed
      })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '일괄 처리에 실패했습니다.', 'error');
      return;
    }
    showToast(data.message || `${list.length}건 ${actionName} 완료!`);
    await loadAdminData();
  } catch (err) {
    console.error(err);
    showToast('일괄 처리 중 통신 오류가 발생했습니다.', 'error');
  } finally {
    restoreBtn();
  }
}

// ===== 10. 특근 테이블 뷰 & 요구사항 13: 소팅 지원 =====
function renderAdminOvertimeTable() {
  const tbody = document.getElementById('adminOvertimeTbody');
  tbody.innerHTML = '';
  selectedOvertimeIds.clear();
  updateSelectedCountText();
  const selectAllEl = document.getElementById('selectAllCheckbox');
  if (selectAllEl) selectAllEl.checked = false;

  let list = [...lastFetchedOvertimes];

  // 정렬 적용
  list.sort((a, b) => {
    let vA = a[otSortCol] ?? '';
    let vB = b[otSortCol] ?? '';
    if (typeof vA === 'string') {
      return otSortAsc ? vA.localeCompare(vB) : vB.localeCompare(vA);
    }
    return otSortAsc ? (vA > vB ? 1 : -1) : (vA < vB ? 1 : -1);
  });

  if (list.length === 0) {
    tbody.innerHTML = '<tr><td colspan="14" style="text-align:center; padding: 2rem; color:var(--text-muted);">조건에 일치하는 특근 내역이 없습니다.</td></tr>';
    return;
  }

  list.forEach(item => {
    const tr = document.createElement('tr');
    const isConf = item.is_confirmed === 1;

    tr.innerHTML = `
      <td style="text-align: center;">
        <input type="checkbox" class="row-checkbox" data-id="${item.id}">
      </td>
      <td>
        <a href="javascript:void(0)" onclick="showUserStatsPopup('${item.emp_id}', '${escapeHtml(item.user_name)}')" style="color:var(--primary); font-weight:700; text-decoration:underline; cursor:pointer;" title="클릭 시 ${escapeHtml(item.user_name)} 님의 상세 통계 팝업 열기">
          ${escapeHtml(item.user_name)}
        </a>
        <span style="font-size:0.78rem; color:var(--text-light);">(${item.emp_id})</span>
      </td>
      <td>${item.team}</td>
      <td><span class="cat-badge ${item.category}">${item.category}</span></td>
      <td>${item.start_date} ~ ${item.end_date}</td>
      <td>
        ${item.is_pre_deduct === 1 ? `<span class="badge" style="background:#0284c7; color:#fff; font-size:0.75rem; padding:2px 6px; border-radius:4px; font-weight:700; margin-right:4px;">⚡ 사전차감</span>` : ''}
        ${item.category === '대체휴무' && item.trip_start_date ? `<span style="background:#ecfdf5; color:#065f46; border:1px solid #a7f3d0; font-weight:700; padding:2px 6px; border-radius:4px; font-size:0.75rem; display:inline-block; white-space:nowrap;">✈️ ${escapeHtml(item.trip_start_date)}~${escapeHtml(item.trip_end_date)}</span>` : ''}
        ${!item.is_pre_deduct && !(item.category === '대체휴무' && item.trip_start_date) && !item.sub_holiday_date ? `<span style="color:#94a3b8; font-size:0.8rem;">-</span>` : ''}
        ${item.sub_holiday_date ? `<span style="background:#fef3c7; color:#b45309; border:1px solid #fde68a; font-weight:700; padding:2px 8px; border-radius:6px; font-size:0.75rem; display:inline-block; white-space:nowrap;">🏖️ ${escapeHtml(item.sub_holiday_date)} (${item.sub_holiday_used}일)</span>` : ''}
      </td>
      <td>${item.project_no || '-'}</td>
      <td>${item.location || '-'}</td>
      <td style="max-width: 220px; word-break: break-all;">${escapeHtml(item.reason || '-')}</td>
      <td style="text-align: center;">
        ${item.bonus_granted === 1 ? `<span class="badge" style="background:#8b5cf6; color:white; font-size:0.75rem; padding:2px 8px; border-radius:9999px; font-weight:700; display:inline-block;">🎁 부여</span>` : `<span style="color:#94a3b8; font-size:0.8rem;">-</span>`}
      </td>
      <td style="text-align: center;">
        <label class="toggle-switch" title="사전차감 원클릭 전환">
          <input type="checkbox" class="pre-deduct-toggle" data-id="${item.id}" ${item.is_pre_deduct === 1 ? 'checked' : ''}>
          <span class="toggle-slider"></span>
        </label>
      </td>
      <td>
        <label class="toggle-switch" title="확인(승인) 원클릭 전환">
          <input type="checkbox" class="confirm-toggle" data-id="${item.id}" ${isConf ? 'checked' : ''}>
          <span class="toggle-slider"></span>
        </label>
      </td>
      <td style="font-size: 0.8rem; color: var(--text-muted);">${item.confirmed_by || '-'}</td>
      <td>
        <div style="display: flex; gap: 0.3rem;">
          <button class="btn btn-secondary btn-sm" onclick="openHistoryModal(${item.id})">이력</button>
          <button class="btn btn-secondary btn-sm" onclick="openAdminEditModal(${item.id})">수정</button>
          <button class="btn btn-danger btn-sm" onclick="handleDeleteOvertime(${item.id})">삭제</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll('.row-checkbox').forEach(cb => {
    cb.addEventListener('change', () => {
      const id = parseInt(cb.getAttribute('data-id'), 10);
      if (cb.checked) selectedOvertimeIds.add(id);
      else selectedOvertimeIds.delete(id);
      updateSelectedCountText();
    });
  });

  tbody.querySelectorAll('.confirm-toggle').forEach(tg => {
    tg.addEventListener('change', async () => {
      const id = parseInt(tg.getAttribute('data-id'), 10);
      const isConfirmed = tg.checked ? 1 : 0;
      await toggleConfirmOvertime(id, isConfirmed);
    });
  });

  tbody.querySelectorAll('.pre-deduct-toggle').forEach(tg => {
    tg.addEventListener('change', async () => {
      const id = parseInt(tg.getAttribute('data-id'), 10);
      await togglePreDeductOvertime(id, tg.checked);
    });
  });
}

// 사전차감 원클릭 토글 함수 (요구사항 6)
async function togglePreDeductOvertime(itemId, isPreDeduct) {
  try {
    const res = await fetch(`/api/overtimes/${itemId}/pre-deduct`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        admin_emp_id: currentUser ? currentUser.emp_id : '',
        is_pre_deduct: isPreDeduct ? 1 : 0
      })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '사전차감 상태 변경 실패', 'error');
      await loadAdminData();
      return;
    }
    showToast(data.message || '사전차감 상태가 변경되었습니다.');
    await loadAdminData();
  } catch (err) {
    console.error(err);
    showToast('통신 오류가 발생했습니다.', 'error');
  }
}

// 특근 테이블 헤더 소팅 클릭 바인딩
document.querySelectorAll('#adminTableView th.sortable').forEach(th => {
  th.addEventListener('click', () => {
    const col = th.getAttribute('data-sort');
    if (otSortCol === col) {
      otSortAsc = !otSortAsc;
    } else {
      otSortCol = col;
      otSortAsc = true;
    }
    // 소팅 아이콘 업데이트
    document.querySelectorAll('#adminTableView th.sortable .sort-icon').forEach(icon => icon.textContent = '↕');
    th.querySelector('.sort-icon').textContent = otSortAsc ? '▲' : '▼';
    renderAdminOvertimeTable();
  });
});

async function toggleConfirmOvertime(itemId, isConfirmed) {
  try {
    const res = await fetch(`/api/overtimes/${itemId}/confirm`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        admin_emp_id: currentUser.emp_id,
        is_confirmed: isConfirmed
      })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '상태 변경 실패', 'error');
      loadAdminData();
      return;
    }
    showToast(data.message);
    loadAdminData();
  } catch (err) {
    console.error(err);
    showToast('통신 오류', 'error');
  }
}

const selectAllCheckbox = document.getElementById('selectAllCheckbox');
if (selectAllCheckbox) {
  selectAllCheckbox.addEventListener('change', () => {
    const isChecked = selectAllCheckbox.checked;
    document.querySelectorAll('#adminOvertimeTbody .row-checkbox').forEach(cb => {
      cb.checked = isChecked;
      const id = parseInt(cb.getAttribute('data-id'), 10);
      if (isChecked) selectedOvertimeIds.add(id);
      else selectedOvertimeIds.delete(id);
    });
    updateSelectedCountText();
  });
}

function updateSelectedCountText() {
  const el = document.getElementById('selectedCountText');
  if (el) el.textContent = `${selectedOvertimeIds.size}개 선택됨`;
}

const batchConfirmBtn = document.getElementById('batchConfirmBtn');
if (batchConfirmBtn) {
  batchConfirmBtn.addEventListener('click', () => handleBatchConfirm(1));
}

const batchUnconfirmBtn = document.getElementById('batchUnconfirmBtn');
if (batchUnconfirmBtn) {
  batchUnconfirmBtn.addEventListener('click', () => handleBatchConfirm(0));
}

const batchDeleteBtn = document.getElementById('batchDeleteBtn');
if (batchDeleteBtn) {
  batchDeleteBtn.addEventListener('click', handleBatchDelete);
}

async function handleBatchDelete() {
  if (selectedOvertimeIds.size === 0) {
    showToast('삭제할 특근 항목의 체크박스를 선택해주세요.', 'error');
    return;
  }
  if (!confirm(`선택한 ${selectedOvertimeIds.size}건의 특근/휴무 내역을 정말로 일괄 삭제하시겠습니까?\n\n⚠️ 삭제된 데이터는 복구할 수 없으며 변경 이력(Audit Log)에 기록됩니다.`)) return;

  const targetBtn = document.getElementById('batchDeleteBtn');
  const restoreBtn = setButtonLoading(targetBtn, '일괄 삭제 중...');

  try {
    const res = await fetch('/api/overtimes/batch-delete', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ids: Array.from(selectedOvertimeIds),
        admin_emp_id: currentUser ? currentUser.emp_id : ''
      })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '일괄 삭제 실패', 'error');
      return;
    }
    showToast(data.message || `${selectedOvertimeIds.size}건 일괄 삭제 완료!`);
    selectedOvertimeIds.clear();
    updateSelectedCountText();
    const selectAllEl = document.getElementById('selectAllCheckbox');
    if (selectAllEl) selectAllEl.checked = false;
    await loadAdminData();
  } catch (err) {
    console.error(err);
    showToast('일괄 삭제 중 오류 발생', 'error');
  } finally {
    restoreBtn();
  }
}

async function handleBatchConfirm(isConfirmed) {
  if (selectedOvertimeIds.size === 0) {
    if (adminViewMode === 'calendar') {
      if (adminSelectedDate) {
        const dayOvertimes = lastFetchedOvertimes.filter(o => o.start_date <= adminSelectedDate && o.end_date >= adminSelectedDate);
        if (dayOvertimes.length > 0) {
          const actionName = isConfirmed === 1 ? '일괄 확인(승인)' : '일괄 취소';
          if (confirm(`달력에서 선택된 날짜(${adminSelectedDate})의 특근자 ${dayOvertimes.length}명을 모두 ${actionName} 처리하시겠습니까?\n\n(여러 건을 체크박스로 개별 선택하시려면 '취소'를 누르시면 테이블 목록 뷰로 이동합니다)`)) {
            handleDailyBatch(dayOvertimes, isConfirmed);
            return;
          }
        }
      }
      showToast('📋 [테이블 목록 뷰]로 이동합니다. 원하는 항목의 체크박스를 선택해주세요.');
      document.getElementById('viewTabTable').click();
      return;
    }
    showToast('처리할 항목의 체크박스를 선택해주세요.', 'error');
    return;
  }
  const actionName = isConfirmed === 1 ? '확인(승인)' : '승인 취소';
  if (!confirm(`선택한 ${selectedOvertimeIds.size}건을 일괄 ${actionName} 처리하시겠습니까?`)) return;

  const targetBtn = isConfirmed === 1 ? document.getElementById('batchConfirmBtn') : document.getElementById('batchUnconfirmBtn');
  const restoreBtn = setButtonLoading(targetBtn, isConfirmed === 1 ? '일괄 승인 중...' : '일괄 취소 중...');

  try {
    const res = await fetch('/api/overtimes/batch-confirm', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ids: Array.from(selectedOvertimeIds),
        admin_emp_id: currentUser ? currentUser.emp_id : '',
        is_confirmed: isConfirmed
      })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '일괄 처리 실패', 'error');
      return;
    }
    showToast(data.message || `${selectedOvertimeIds.size}건 일괄 처리 완료!`);
    selectedOvertimeIds.clear();
    updateSelectedCountText();
    const selectAllEl = document.getElementById('selectAllCheckbox');
    if (selectAllEl) selectAllEl.checked = false;
    await loadAdminData();
  } catch (err) {
    console.error(err);
    showToast('일괄 처리 중 오류 발생', 'error');
  } finally {
    restoreBtn();
  }
}

// 필터 바 이벤트
['filterStartDate', 'filterEndDate', 'filterTeam', 'filterCategory', 'filterStatus'].forEach(id => {
  document.getElementById(id).addEventListener('change', loadAdminData);
});
document.getElementById('filterSearch').addEventListener('input', debounce(loadAdminData, 300));

document.getElementById('filterResetBtn').addEventListener('click', () => {
  document.getElementById('filterStartDate').value = '';
  document.getElementById('filterEndDate').value = '';
  document.getElementById('filterTeam').value = '';
  document.getElementById('filterCategory').value = '';
  document.getElementById('filterStatus').value = '';
  document.getElementById('filterSearch').value = '';
  selectedDeptFilter = '';
  loadAdminData();
});

function debounce(func, wait) {
  let timeout;
  return function(...args) {
    clearTimeout(timeout);
    timeout = setTimeout(() => func.apply(this, args), wait);
  };
}

// 요구사항 5: 관리자 팀원 특근/대체휴무 대리 신청
function setProxyPreDeduct(active) {
  const btn = document.getElementById('proxyPreDeductToggleBtn');
  const chk = document.getElementById('proxyIsPreDeduct');
  if (!btn || !chk) return;
  chk.checked = !!active;
  if (active) {
    btn.textContent = 'ON';
    btn.className = 'btn btn-sm btn-primary';
    btn.style.background = 'linear-gradient(135deg, #0284c7 0%, #0369a1 100%)';
    btn.style.color = '#fff';
  } else {
    btn.textContent = 'OFF';
    btn.className = 'btn btn-sm btn-secondary';
    btn.style.background = '';
    btn.style.color = '';
  }
}

const proxyPreDeductToggleBtn = document.getElementById('proxyPreDeductToggleBtn');
if (proxyPreDeductToggleBtn) {
  proxyPreDeductToggleBtn.addEventListener('click', () => {
    const chk = document.getElementById('proxyIsPreDeduct');
    setProxyPreDeduct(!chk?.checked);
  });
}

const proxyCategoryEl = document.getElementById('proxyCategory');
if (proxyCategoryEl) {
  proxyCategoryEl.addEventListener('change', () => {
    const tripRow = document.getElementById('proxyTripDatesRow');
    if (tripRow) {
      tripRow.style.display = proxyCategoryEl.value === '대체휴무' ? 'grid' : 'none';
    }
  });
}

const adminProxyOvertimeBtn = document.getElementById('adminProxyOvertimeBtn');
if (adminProxyOvertimeBtn) {
  adminProxyOvertimeBtn.addEventListener('click', async () => {
    // 팀원 목록 불러오기
    if (!allUsersCache || allUsersCache.length === 0) {
      try {
        const adminParam = currentUser ? `?admin_emp_id=${encodeURIComponent(currentUser.emp_id)}` : '';
        const res = await fetch('/api/users' + adminParam);
        const data = await res.json();
        if (res.ok && data.users) {
          allUsersCache = data.users;
        }
      } catch (e) {
        console.error(e);
      }
    }

    const proxySelect = document.getElementById('proxyUserSelect');
    if (proxySelect) {
      proxySelect.innerHTML = '<option value="">팀원을 선택하세요</option>';
      
      // ps37082 사원번호만 제외하고 모든 팀원 표시 (요구사항 2)
      let candidateUsers = (allUsersCache || []).filter(u => {
        return (u.emp_id || '').toLowerCase() !== 'ps37082';
      });

      candidateUsers.sort((a, b) => (a.name || '').localeCompare(b.name || ''));

      candidateUsers.forEach(u => {
        const opt = document.createElement('option');
        opt.value = u.emp_id;
        opt.textContent = `${u.name} (${u.emp_id} / ${u.team} ${u.position || '팀원'})`;
        proxySelect.appendChild(opt);
      });
    }

    // 기본값 설정 (캘린더 선택 날짜 우선 반영, 없으면 토요일 기본값) (요구사항 1)
    const defaultDate = adminSelectedDate || getThisSaturdayStr();
    const startDateInput = document.getElementById('proxyStartDate');
    const endDateInput = document.getElementById('proxyEndDate');
    if (startDateInput) startDateInput.value = defaultDate;
    if (endDateInput) endDateInput.value = defaultDate;

    const proxyCategory = document.getElementById('proxyCategory');
    if (proxyCategory) proxyCategory.value = '일반휴일';

    const tripRow = document.getElementById('proxyTripDatesRow');
    if (tripRow) tripRow.style.display = 'none';
    const pTripStart = document.getElementById('proxyTripStartDate');
    const pTripEnd = document.getElementById('proxyTripEndDate');
    if (pTripStart) pTripStart.value = '';
    if (pTripEnd) pTripEnd.value = '';
    setProxyPreDeduct(false);

    const pNo = document.getElementById('proxyProjectNo');
    const pLoc = document.getElementById('proxyLocation');
    const pReason = document.getElementById('proxyReason');
    const pBonus = document.getElementById('proxyBonusGranted');

    if (pNo) pNo.value = 'BT2601-L1';
    if (pLoc) pLoc.value = '본사5층';
    if (pReason) pReason.value = '프로그램 개발';
    if (pBonus) pBonus.checked = false;

    openModal('adminProxyOvertimeModal');
  });
}

const adminProxyOvertimeForm = document.getElementById('adminProxyOvertimeForm');
if (adminProxyOvertimeForm) {
  adminProxyOvertimeForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const empId = document.getElementById('proxyUserSelect').value;
    const category = document.getElementById('proxyCategory').value;
    const startDate = document.getElementById('proxyStartDate').value;
    const endDate = document.getElementById('proxyEndDate').value;
    const projectNo = (document.getElementById('proxyProjectNo').value || '').trim();
    const location = (document.getElementById('proxyLocation').value || '').trim();
    const reason = (document.getElementById('proxyReason').value || '').trim();
    const bonusGranted = document.getElementById('proxyBonusGranted').checked ? 1 : 0;
    const isPreDeduct = document.getElementById('proxyIsPreDeduct')?.checked ? 1 : 0;
    const tripStartDate = (document.getElementById('proxyTripStartDate')?.value || '').trim();
    const tripEndDate = (document.getElementById('proxyTripEndDate')?.value || '').trim();

    if (!empId) {
      showToast('신청 대상 팀원을 선택해주세요.', 'error');
      return;
    }
    if (!startDate || !endDate) {
      showToast('특근/휴무 시작일과 종료일을 입력해주세요.', 'error');
      return;
    }
    if (startDate > endDate) {
      showToast('종료일은 시작일보다 빠를 수 없습니다.', 'error');
      return;
    }

    const submitBtn = adminProxyOvertimeForm.querySelector('button[type="submit"]');
    const restoreBtn = setButtonLoading(submitBtn, '등록 저장 중...');

    try {
      const res = await fetch('/api/overtimes', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          emp_id: empId,
          category,
          start_date: startDate,
          end_date: endDate,
          project_no: projectNo,
          location,
          reason,
          bonus_granted: bonusGranted,
          is_pre_deduct: isPreDeduct,
          trip_start_date: tripStartDate,
          trip_end_date: tripEndDate
        })
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || '특근/휴무 등록에 실패했습니다.', 'error');
        return;
      }

      showToast(`팀원의 특근/대체휴무가 성공적으로 등록되었습니다! ${bonusGranted === 1 ? '(🎁 보너스 포함)' : ''}`);
      closeModal('adminProxyOvertimeModal');
      await loadAdminData();
    } catch (err) {
      console.error(err);
      showToast('통신 오류가 발생했습니다. 네트워크 상태를 확인해주세요.', 'error');
    } finally {
      restoreBtn();
    }
  });
}

// ===== 11. 엑셀 (.xlsx) 내보내기 =====
document.getElementById('exportExcelBtn').addEventListener('click', async () => {
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
  showToast('엑셀 파일을 생성 중입니다...');

  // 1. 고품질 다중 시트 openpyxl 백엔드 API 우선 호출 (서버 환경)
  try {
    const reqBody = {
      ids: selectedOvertimeIds.size > 0 ? Array.from(selectedOvertimeIds) : undefined,
      start_date: document.getElementById('filterStartDate')?.value || undefined,
      end_date: document.getElementById('filterEndDate')?.value || undefined,
      team: selectedDeptFilter || document.getElementById('filterTeam')?.value || undefined,
      category: document.getElementById('filterCategory')?.value || undefined,
      is_confirmed: document.getElementById('filterStatus')?.value !== '' ? parseInt(document.getElementById('filterStatus').value) : undefined,
      search: document.getElementById('filterSearch')?.value?.trim() || undefined,
      admin_emp_id: currentUser ? currentUser.emp_id : undefined
    };
    const res = await fetch('/api/overtimes/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(reqBody)
    });
    if (res.ok) {
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `특근현황_일자별및개인별정산_${getTodayStr()}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      showToast('엑셀 다운로드가 완료되었습니다!');
      restoreExportBtn();
      return;
    }
  } catch (apiErr) {
    console.warn('API export fallback to client-side SheetJS:', apiErr);
  } finally {
    restoreExportBtn();
  }

  // 2. 오프라인 또는 Vercel 환경 클라이언트 사이드 SheetJS 폴백
  if (typeof XLSX !== 'undefined') {
    try {
      // 1. 휴일일자별_특근현황 시트 (요구사항 27-1, 27-2)
      const WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"];
      const expandedRows = [];
      targetList.forEach(item => {
        let d1, d2;
        try {
          d1 = new Date(item.start_date);
          d2 = new Date(item.end_date || item.start_date);
          if (isNaN(d1.getTime())) d1 = new Date();
          if (isNaN(d2.getTime()) || d2 < d1) d2 = d1;
        } catch (e) {
          d1 = new Date();
          d2 = d1;
        }

        const totalDays = Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1;
        const subUsed = Number(item.sub_holiday_used || 0);
        const subDate = item.sub_holiday_date || '-';
        const isConf = item.is_confirmed === 1;

        let curr = new Date(d1);
        let dayIdx = 0;
        while (curr <= d2) {
          const y = curr.getFullYear();
          const m = String(curr.getMonth() + 1).padStart(2, '0');
          const d = String(curr.getDate()).padStart(2, '0');
          const dateStr = `${y}-${m}-${d}`;
          const wStr = WEEKDAYS[curr.getDay()];

          expandedRows.push({
            holiday_date: dateStr,
            weekday: wStr,
            emp_id: item.emp_id,
            user_name: item.user_name,
            team: item.team,
            category: item.category,
            holiday_days: 1, // 요구사항 27-2: 순수 숫자 1
            sub_holiday_date: (dayIdx === 0 && subDate !== '-') ? subDate : (totalDays > 1 ? '-' : subDate),
            sub_holiday_used: dayIdx === 0 ? subUsed : 0,
            project_no: item.project_no || '-',
            location: item.location || '-',
            reason: item.reason || '-',
            bonus_granted: item.bonus_granted || 0,
            status_text: isConf ? '확인완료' : '승인대기',
            confirmed_by: item.confirmed_by || '-',
            created_at: item.created_at || ''
          });

          curr.setDate(curr.getDate() + 1);
          dayIdx++;
        }
      });

      expandedRows.sort((a, b) => a.holiday_date.localeCompare(b.holiday_date) || a.team.localeCompare(b.team) || a.user_name.localeCompare(b.user_name));

      const sheet1Rows = expandedRows.map((r, idx) => ({
        "순번": idx + 1,
        "휴일날짜": r.holiday_date,
        "요일": r.weekday,
        "사번": r.emp_id,
        "성명": r.user_name,
        "소속팀": r.team,
        "특근분류": r.category,
        "휴일일수": r.holiday_days, // 순수 숫자
        "대체휴일 사용일": r.sub_holiday_date,
        "프로젝트 번호": r.project_no,
        "근무 장소": r.location,
        "특근 사유": r.reason,
        "보너스 부여": (r.bonus_granted === 1 ? "부여(O)" : "미부여(-)"),
        "확인(승인)": r.status_text,
        "확인자": r.confirmed_by,
        "신청일시": r.created_at
      }));

      // 2. 개인별 휴일합산 정산 시트 (요구사항 27-2, 27-3)
      // 0. 대체휴무 등록 건들에서 사원별 출장기간 수집
      const userTrips = {};
      targetList.forEach(item => {
        const cat = (item.category || '').trim();
        const ts = (item.trip_start_date || '').trim();
        const te = (item.trip_end_date || '').trim();
        const empId = item.emp_id;
        if ((cat === '대체휴무' || cat === '대체휴일') && ts && te && empId) {
          if (!userTrips[empId]) userTrips[empId] = [];
          userTrips[empId].push({ start: ts, end: te });
        }
      });

      const userMap = {};
      targetList.forEach(item => {
        let days = 1;
        try {
          const d1 = new Date(item.start_date);
          const d2 = new Date(item.end_date || item.start_date);
          days = Math.max(1, Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1);
        } catch (e) {}

        const empId = item.emp_id;
        if (!userMap[empId]) {
          userMap[empId] = {
            emp_id: empId,
            name: item.user_name,
            team: item.team,
            sub_work_days: 0,
            legal_holiday_days: 0,
            normal_holiday_days: 0,
            sub_holiday_days: 0,
            pre_deduct_count: 0,
            trip_pre_deduct_count: 0,
            total_days: 0,
            sub_holiday_used: 0,
            bonus_count: 0,
            records_count: 0
          };
        }
        const u = userMap[empId];
        u.records_count += 1;
        u.sub_holiday_used += Number(item.sub_holiday_used || 0);

        if (item.category === '대체근무') u.sub_work_days += days;
        else if (item.category === '법정휴일') u.legal_holiday_days += days;
        else if (item.category === '일반휴일') u.normal_holiday_days += days;
        else if (item.category === '대체휴무' || item.category === '대체휴일') u.sub_holiday_days += days;

        // 요구사항 1-1: 총 특근일수 = 대체근무 + 법정휴일 + 일반휴일 (대체휴무 제외)
        u.total_days = u.sub_work_days + u.legal_holiday_days + u.normal_holiday_days;

        // 요구사항 1-3: 보너스 건수 집계
        if (item.bonus_granted === 1) {
          u.bonus_count += 1;
        }

        if (item.is_pre_deduct === 1) {
          u.pre_deduct_count += 1;
          const sStr = item.start_date;
          const eStr = item.end_date || item.start_date;
          const trips = userTrips[empId] || [];
          for (const tr of trips) {
            if (!(eStr < tr.start || sStr > tr.end)) {
              u.trip_pre_deduct_count += 1;
              break;
            }
          }
        }
      });

      const userList = Object.values(userMap).sort((a, b) => a.team.localeCompare(b.team) || a.name.localeCompare(b.name));
      const sheet2Rows = userList.map((u, idx) => {
        const totalSub = u.sub_holiday_used + u.sub_holiday_days;
        // 최종 실특근 = 일반특근 - 사전차감 - (대체휴무 - 대체휴무시 작성한 출장기간 이내의 사전차감)
        const actualOvertime = Math.max(0, Math.round((u.normal_holiday_days - u.pre_deduct_count - (totalSub - u.trip_pre_deduct_count)) * 10) / 10);
        // 사전차감 잔여 = 총 사전차감 - 출장내 사전차감
        const preRemain = Math.max(0, u.pre_deduct_count - u.trip_pre_deduct_count);
        // 최종 실특근일 + 보너스 합산
        const actualOvertimeWithBonus = Math.max(0, Math.round((actualOvertime + (u.bonus_count || 0)) * 10) / 10);
        return {
          "순번": idx + 1,
          "사원번호": u.emp_id,
          "성명": u.name,
          "소속팀": u.team,
          "대체근무 일수": u.sub_work_days,
          "법정휴일 일수": u.legal_holiday_days,
          "일반휴일 일수": u.normal_holiday_days,
          "대체휴무 일수": totalSub,
          "총 특근일수": u.total_days,
          "총 사전차감": u.pre_deduct_count,
          "사전차감 잔여수": preRemain,
          "★ 최종 실특근일": actualOvertime,
          "보너스 부여 (건)": u.bonus_count,
          "★ 최종 실특근일+보너스": actualOvertimeWithBonus,
          "신청건수": u.records_count
        };
      });

      // 개인별 합산 총계 행
      if (sheet2Rows.length > 0) {
        sheet2Rows.push({
          "순번": "합계",
          "사원번호": "-",
          "성명": `${sheet2Rows.length}명`,
          "소속팀": "-",
          "대체근무 일수": sheet2Rows.reduce((a, b) => a + (b["대체근무 일수"] || 0), 0),
          "법정휴일 일수": sheet2Rows.reduce((a, b) => a + (b["법정휴일 일수"] || 0), 0),
          "일반휴일 일수": sheet2Rows.reduce((a, b) => a + (b["일반휴일 일수"] || 0), 0),
          "대체휴무 일수": Math.round(sheet2Rows.reduce((a, b) => a + (b["대체휴무 일수"] || 0), 0) * 10) / 10,
          "총 특근일수": sheet2Rows.reduce((a, b) => a + (b["총 특근일수"] || 0), 0),
          "총 사전차감": sheet2Rows.reduce((a, b) => a + (b["총 사전차감"] || 0), 0),
          "사전차감 잔여수": sheet2Rows.reduce((a, b) => a + (b["사전차감 잔여수"] || 0), 0),
          "★ 최종 실특근일": Math.round(sheet2Rows.reduce((a, b) => a + (b["★ 최종 실특근일"] || 0), 0) * 10) / 10,
          "보너스 부여 (건)": sheet2Rows.reduce((a, b) => a + (b["보너스 부여 (건)"] || 0), 0),
          "★ 최종 실특근일+보너스": Math.round(sheet2Rows.reduce((a, b) => a + (b["★ 최종 실특근일+보너스"] || 0), 0) * 10) / 10,
          "신청건수": sheet2Rows.reduce((a, b) => a + (b["신청건수"] || 0), 0)
        });
      }

      // 3. 특근신청 전체원장 시트
      const sheet3Rows = targetList.map((item, idx) => {
        let days = 1;
        try {
          const d1 = new Date(item.start_date);
          const d2 = new Date(item.end_date || item.start_date);
          days = Math.max(1, Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1);
        } catch (e) {}

        return {
          "순번": idx + 1,
          "사번": item.emp_id,
          "성명": item.user_name,
          "소속팀": item.team,
          "특근분류": item.category,
          "시작일": item.start_date,
          "종료일": item.end_date || item.start_date,
          "일수": days,
          "대체휴무 사용일": item.sub_holiday_date || '-',
          "대체휴무 사용일수": Number(item.sub_holiday_used || 0),
          "사전차감": item.is_pre_deduct === 1 ? 'O' : '-',
          "출장시작일": item.trip_start_date || '-',
          "출장종료일": item.trip_end_date || '-',
          "프로젝트 번호": item.project_no || '-',
          "근무 장소": item.location || '-',
          "특근 사유": item.reason || '-',
          "보너스 부여": (item.bonus_granted === 1 ? "부여(O)" : "미부여(-)"),
          "확인(승인)": item.is_confirmed === 1 ? '확인완료' : '승인대기',
          "확인자": item.confirmed_by || '-',
          "신청일시": item.created_at || ''
        };
      });

      const workbook = XLSX.utils.book_new();

      const ws1 = XLSX.utils.json_to_sheet(sheet1Rows);
      ws1['!cols'] = [
        { wch: 6 }, { wch: 12 }, { wch: 6 }, { wch: 10 }, { wch: 10 },
        { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 14 },
        { wch: 16 }, { wch: 16 }, { wch: 26 }, { wch: 10 }, { wch: 14 }, { wch: 18 }
      ];
      XLSX.utils.book_append_sheet(workbook, ws1, "휴일일자별_특근현황");

      const ws2 = XLSX.utils.json_to_sheet(sheet2Rows);
      ws2['!cols'] = [
        { wch: 6 },  // 순번
        { wch: 12 }, // 사원번호
        { wch: 10 }, // 성명
        { wch: 14 }, // 소속팀
        { wch: 14 }, // 대체근무 일수
        { wch: 14 }, // 법정휴일 일수
        { wch: 14 }, // 일반휴일 일수
        { wch: 14 }, // 대체휴무 일수
        { wch: 14 }, // 총 특근일수
        { wch: 14 }, // 총 사전차감
        { wch: 16 }, // 사전차감 잔여수
        { wch: 16 }, // ★ 최종 실특근일
        { wch: 16 }, // 보너스 부여 (건)
        { wch: 22 }, // ★ 최종 실특근일+보너스
        { wch: 10 }  // 신청건수
      ];
      XLSX.utils.book_append_sheet(workbook, ws2, "개인별_휴일합산_정산표");

      const ws3 = XLSX.utils.json_to_sheet(sheet3Rows);
      ws3['!cols'] = [
        { wch: 6 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 12 },
        { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 14 }, { wch: 14 },
        { wch: 16 }, { wch: 16 }, { wch: 26 }, { wch: 10 }, { wch: 14 }, { wch: 18 }
      ];
      XLSX.utils.book_append_sheet(workbook, ws3, "특근신청_원장");

      XLSX.writeFile(workbook, `특근현황_일자별및개인별정산_${getTodayStr()}.xlsx`);
      showToast('엑셀 다운로드가 완료되었습니다!');
      return;
    } catch (e) {
      console.error('Export error:', e);
      showToast('엑셀 생성 중 오류가 발생했습니다.', 'error');
    }
  } else {
    showToast('엑셀 처리 라이브러리를 불러올 수 없습니다.', 'error');
  }
});

// ===== 12. 팀원 관리 & 요구사항 12: 엑셀 내보내기/가져오기 & 소팅 =====
async function loadAdminUserTable() {
  const tbody = document.getElementById('adminUserTbody');
  tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding: 2rem;">팀원 목록 로딩 중...</td></tr>';

  try {
    const adminParam = currentUser ? `?admin_emp_id=${encodeURIComponent(currentUser.emp_id)}` : '';
    const res = await fetch('/api/users' + adminParam);
    const data = await res.json();
    if (!res.ok) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:var(--danger);">조회 실패</td></tr>';
      return;
    }

    allUsersCache = data.users || [];
    renderAdminUserRows();
  } catch (err) {
    console.error(err);
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:var(--danger);">통신 오류</td></tr>';
  }
}

function renderAdminUserRows() {
  const tbody = document.getElementById('adminUserTbody');
  tbody.innerHTML = '';

  let list = [...allUsersCache];

  // 팀원 소팅
  list.sort((a, b) => {
    let vA = a[userSortCol] ?? '';
    let vB = b[userSortCol] ?? '';
    if (typeof vA === 'string') {
      return userSortAsc ? vA.localeCompare(vB) : vB.localeCompare(vA);
    }
    return userSortAsc ? (vA > vB ? 1 : -1) : (vA < vB ? 1 : -1);
  });

  const isCurrentUserSuper = currentUser && (currentUser.is_super === 1 || currentUser.emp_id.toLowerCase() === 'ps37082');

  list.forEach(u => {
    const tr = document.createElement('tr');
    const isAdmin = u.is_admin === 1 || u.is_super === 1;
    const isSuperUser = u.emp_id.toLowerCase() === 'ps37082' || u.is_super === 1;
    const canToggle = isCurrentUserSuper && !isSuperUser;
    const toggleTitle = !isCurrentUserSuper ? '관리자 지정 권한은 총괄관리자만 가능합니다' : (isSuperUser ? '총괄관리자 권한은 고정입니다' : '관리자 권한 토글');

    // v1.41: 3단계 권한 선택 UI
    const canEditRole = isCurrentUserSuper && u.emp_id.toLowerCase() !== 'ps37082' && u.emp_id !== currentUser.emp_id;
    let roleLevel = 0; // 0=팀원, 1=팀관리자, 2=슈퍼관리자
    if (u.is_super === 1) roleLevel = 2;
    else if (u.is_admin === 1) roleLevel = 1;

    let roleSelectHtml = '';
    if (canEditRole) {
      roleSelectHtml = `
        <select class="role-level-select" data-empid="${u.emp_id}" data-name="${escapeHtml(u.name)}" style="font-size:0.8rem; padding:3px 6px; border-radius:6px; border:1px solid var(--border-color); background:var(--bg-input); color:var(--text-main); cursor:pointer;">
          <option value="0" ${roleLevel===0?'selected':''}>👤 팀원</option>
          <option value="1" ${roleLevel===1?'selected':''}>🛡️ 팀관리자</option>
          <option value="2" ${roleLevel===2?'selected':''}>👑 슈퍼관리자</option>
        </select>`;
    } else {
      if (roleLevel === 2) {
        roleSelectHtml = '<span class="badge" style="background:#fbbf24; color:#78350f; font-weight:700; font-size:0.72rem;">👑 슈퍼관리자</span>';
      } else if (roleLevel === 1) {
        roleSelectHtml = '<span class="badge" style="background:#0284c7; color:#fff; font-weight:700; font-size:0.72rem;">🛡️ 팀관리자</span>';
      } else {
        roleSelectHtml = '<span class="badge" style="background:#e2e8f0; color:#475569; font-size:0.72rem;">👤 팀원</span>';
      }
    }

    tr.innerHTML = `
      <td><b>${u.emp_id}</b></td>
      <td>${escapeHtml(u.name)}</td>
      <td>${escapeHtml(u.team)}</td>
      <td>${escapeHtml(u.position || '팀원')}</td>
      <td>
        ${roleSelectHtml}
      </td>
      <td>
        <div style="display: flex; gap: 4px; align-items: center; flex-wrap: wrap;">
          <button type="button" class="btn btn-secondary btn-sm" onclick="showUserStatsPopup('${u.emp_id}', '${escapeHtml(u.name)}')" style="padding: 3px 8px; font-size: 0.78rem;">📊 통계</button>
          <button type="button" class="btn btn-secondary btn-sm edit-user-btn" onclick="openUserEditModal('${u.emp_id}')" data-empid="${u.emp_id}" style="padding: 3px 8px; font-size: 0.78rem;">✏️ 수정</button>
          <button type="button" class="btn btn-danger btn-sm delete-user-btn" onclick="handleDeleteUser('${u.emp_id}', '${escapeHtml(u.name)}')" data-empid="${u.emp_id}" data-name="${escapeHtml(u.name)}" ${isSuperUser ? 'disabled title="총괄관리자는 삭제할 수 없습니다"' : ''} style="padding: 3px 8px; font-size: 0.78rem;">삭제</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // v1.41: 3단계 권한 드롭다운 이벤트 바인딩
  tbody.querySelectorAll('.role-level-select').forEach(sel => {
    sel.addEventListener('change', async () => {
      const empId = sel.getAttribute('data-empid');
      const name = sel.getAttribute('data-name');
      const newLevel = parseInt(sel.value, 10);
      let confirmMsg = '';
      if (newLevel === 2) confirmMsg = `'${name}'(${empId}) 님을 슈퍼관리자로 승격하시겠습니까?\n슈퍼관리자는 전체 부서 특근 승인, 인원 관리 및 시스템 총괄 제어가 가능합니다.`;
      else if (newLevel === 1) confirmMsg = `'${name}'(${empId}) 님을 팀관리자로 지정하시겠습니까?`;
      else confirmMsg = `'${name}'(${empId}) 님을 일반 팀원으로 변경하시겠습니까?`;
      if (!confirm(confirmMsg)) { await loadAdminUserTable(); return; }

      if (newLevel === 2) {
        await updateUserSuperRole(empId, 1);
      } else if (newLevel === 1) {
        // 팀관리자: is_super=0, is_admin=1
        await fetch(`/api/users/${empId}`, { method: 'PUT', headers: {'Content-Type':'application/json'}, body: JSON.stringify({is_super: 0, is_admin: 1, admin_emp_id: currentUser.emp_id}) });
        showToast(`'${name}' 님이 팀관리자로 지정되었습니다.`);
        await loadAdminUserTable();
      } else {
        // 팀원: is_super=0, is_admin=0
        await fetch(`/api/users/${empId}`, { method: 'PUT', headers: {'Content-Type':'application/json'}, body: JSON.stringify({is_super: 0, is_admin: 0, admin_emp_id: currentUser.emp_id}) });
        showToast(`'${name}' 님이 일반 팀원으로 변경되었습니다.`);
        await loadAdminUserTable();
      }
    });
  });

  // 팀원 수정 / 삭제 버튼 이벤트 바인딩
  tbody.querySelectorAll('.edit-user-btn').forEach(btn => {
    btn.addEventListener('click', () => { openUserEditModal(btn.getAttribute('data-empid')); });
  });
  tbody.querySelectorAll('.delete-user-btn').forEach(btn => {
    btn.addEventListener('click', async () => { await handleDeleteUser(btn.getAttribute('data-empid'), btn.getAttribute('data-name')); });
  });
}

// 요구사항 1: 팀원 정보 수정 모달 열기
function openUserEditModal(empId) {
  const user = allUsersCache.find(u => u.emp_id === empId);
  if (!user) {
    showToast('회원 정보를 찾을 수 없습니다.', 'error');
    return;
  }

  const isCurrentUserSuper = currentUser && (currentUser.is_super === 1 || currentUser.emp_id.toLowerCase() === 'ps37082');
  const isSuperUser = user.emp_id.toLowerCase() === 'ps37082' || user.is_super === 1;

  document.getElementById('editUserEmpIdHidden').value = user.emp_id;
  document.getElementById('editUserEmpId').value = user.emp_id;
  document.getElementById('editUserName').value = user.name;

  // 소속팀 셀렉트 옵션 채우기 (teamsCache 안전 참조)
  const teamSelect = document.getElementById('editUserTeam');
  teamSelect.innerHTML = '';
  const currentTeams = (typeof teamsCache !== 'undefined' && Array.isArray(teamsCache)) ? teamsCache : [];

  currentTeams.forEach(tName => {
    const opt = document.createElement('option');
    opt.value = tName;
    opt.textContent = tName;
    if (tName === user.team) opt.selected = true;
    teamSelect.appendChild(opt);
  });
  if (user.team && !currentTeams.includes(user.team)) {
    const opt = document.createElement('option');
    opt.value = user.team;
    opt.textContent = user.team;
    opt.selected = true;
    teamSelect.appendChild(opt);
  }

  // 팀 이동 권한 제어 (슈퍼관리자만 가능)
  if (!isCurrentUserSuper) {
    teamSelect.disabled = true;
    document.getElementById('editUserTeamHelp').textContent = '팀원의 소속팀 이동은 슈퍼관리자만 가능합니다 (현재 수정 불가).';
  } else {
    teamSelect.disabled = false;
    document.getElementById('editUserTeamHelp').textContent = '이동시킬 새로운 소속팀을 선택하세요.';
  }

  document.getElementById('editUserPosition').value = user.position || '팀원';

  // 권한 단계 셀렉트박스 설정 (요구사항 1: 팀원-팀관리자-슈퍼관리자 단계적 설정)
  const roleSelect = document.getElementById('editUserRoleSelect');
  let currentLevel = 0;
  if (user.is_super === 1) currentLevel = 2;
  else if (user.is_admin === 1) currentLevel = 1;
  if (roleSelect) roleSelect.value = String(currentLevel);

  if (!isCurrentUserSuper || isSuperUser || user.emp_id === currentUser.emp_id) {
    if (roleSelect) roleSelect.disabled = true;
    document.getElementById('editUserRoleHelp').textContent = isSuperUser ? '총괄 슈퍼관리자의 권한은 고정입니다.' : '권한 단계는 슈퍼관리자만 변경할 수 있습니다.';
  } else {
    if (roleSelect) roleSelect.disabled = false;
    document.getElementById('editUserRoleHelp').textContent = '팀원, 팀관리자, 슈퍼관리자 중 원하는 권한 단계를 선택하세요.';
  }

  openModal('userEditModal');
}
window.openUserEditModal = openUserEditModal;

// 요구사항 1: 팀원 정보 수정 폼 submit 핸들러
const editUserForm = document.getElementById('editUserForm');
if (editUserForm) {
  editUserForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const empId = document.getElementById('editUserEmpIdHidden').value;
    const name = document.getElementById('editUserName').value.trim();
    const team = document.getElementById('editUserTeam').value;
    const position = document.getElementById('editUserPosition').value.trim();
    const roleVal = parseInt(document.getElementById('editUserRoleSelect')?.value || '0', 10);
    const isSuper = roleVal === 2 ? 1 : 0;
    const isAdmin = roleVal >= 1 ? 1 : 0;

    if (!name) {
      showToast('성명을 입력해주세요.', 'error');
      return;
    }
    if (!team) {
      showToast('소속팀을 선택해주세요.', 'error');
      return;
    }

    try {
      const res = await fetch(`/api/users/${encodeURIComponent(empId)}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name,
          team,
          position,
          is_admin: isAdmin,
          is_super: isSuper,
          admin_emp_id: currentUser ? currentUser.emp_id : ''
        })
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || '수정에 실패했습니다.', 'error');
        return;
      }

      showToast(`'${name}' 팀원의 정보가 성공적으로 수정되었습니다.`);
      closeModal('userEditModal');
      await loadAdminUserTable();
      await loadAdminData(); // 특근 내역의 성명/팀도 동기화되었으므로 갱신
      await loadTeams();
    } catch (err) {
      console.error(err);
      showToast('서버 통신 오류', 'error');
    }
  });
}

// 팀원 테이블 헤더 소팅 바인딩
document.querySelectorAll('#adminUserView th.sortable').forEach(th => {
  th.addEventListener('click', () => {
    const col = th.getAttribute('data-sort');
    if (userSortCol === col) {
      userSortAsc = !userSortAsc;
    } else {
      userSortCol = col;
      userSortAsc = true;
    }
    document.querySelectorAll('#adminUserView th.sortable .sort-icon').forEach(icon => icon.textContent = '↕');
    th.querySelector('.sort-icon').textContent = userSortAsc ? '▲' : '▼';
    renderAdminUserRows();
  });
});

async function updateUserAdminRole(empId, isAdmin) {
  try {
    const res = await fetch(`/api/users/${empId}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        is_admin: isAdmin,
        admin_emp_id: currentUser ? currentUser.emp_id : ''
      })
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '권한 변경 실패', 'error');
      loadAdminUserTable();
      return;
    }
    showToast(`권한이 변경되었습니다: ${isAdmin ? '관리자 권한 부여' : '일반 권한으로 변경'}`);
    loadAdminUserTable();
  } catch (err) {
    console.error(err);
    showToast('통신 오류', 'error');
  }
}

async function updateUserSuperRole(empId, isSuper) {
  try {
    const payload = isSuper === 1
      ? { is_super: 1, is_admin: 1, admin_emp_id: currentUser ? currentUser.emp_id : '' }
      : { is_super: 0, is_admin: 1, admin_emp_id: currentUser ? currentUser.emp_id : '' }; // 하야 시 팀관리자로
    const res = await fetch(`/api/users/${empId}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '권한 변경 실패', 'error');
      loadAdminUserTable();
      return;
    }
    const actionText = isSuper === 1 ? '슈퍼관리자로 승격되었습니다.' : '팀관리자로 변경되었습니다.';
    showToast(`'${data.user ? data.user.name : empId}' 님이 ${actionText}`);
    loadAdminUserTable();
  } catch (err) {
    console.error(err);
    showToast('통신 오류가 발생했습니다.', 'error');
  }
}

async function handleDeleteUser(empId, uName = '') {
  const nameStr = uName ? `'${uName}(${empId})'` : `사번 ${empId}`;
  if (!confirm(`${nameStr} 팀원을 정말 삭제하시겠습니까?\n(해당 팀원의 특근 내역 데이터도 함께 정리됩니다.)`)) return;
  try {
    const adminParam = currentUser ? `?admin_emp_id=${encodeURIComponent(currentUser.emp_id)}` : '';
    const res = await fetch(`/api/users/${empId}${adminParam}`, { method: 'DELETE' });
    const data = await res.json();
    if (!res.ok) {
      showToast(data.detail || '삭제 실패', 'error');
      return;
    }
    showToast(data.message || '회원이 삭제되었습니다.');
    await loadAdminUserTable();
    await loadTeams();
    renderTeamManageList();
    if (typeof loadAdminData === 'function') await loadAdminData();
  } catch (err) {
    console.error(err);
    showToast('통신 오류', 'error');
  }
}
window.handleDeleteUser = handleDeleteUser;

// 요구사항 12: 팀원 명부 엑셀 내보내기
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
});

// 신규 팀원 추가 모달 열기 (관리자용)
document.getElementById('addNewUserModalBtn').addEventListener('click', async () => {
  isAddingMemberFromAdmin = true;
  const modal = document.getElementById('registerModal');
  const title = modal.querySelector('.modal-title');
  if (title) title.textContent = '👥 신규 팀원 추가 (관리자)';
  const desc = modal.querySelector('.modal-body > p');
  if (desc) desc.textContent = '신규 팀원의 정보를 입력하시면 팀원 명부에 즉시 등록됩니다.';
  const submitBtn = modal.querySelector('button[type="submit"]');
  if (submitBtn) submitBtn.textContent = '+ 팀원 추가하기';

  const regEmpId = document.getElementById('regEmpId');
  regEmpId.value = '';
  regEmpId.removeAttribute('readonly');
  regEmpId.placeholder = '사원번호 입력 (예: 2024002)';
  document.getElementById('regName').value = '';
  await loadTeams();
  document.getElementById('regTeam').value = '';
  document.getElementById('regPosition').value = '팀원';
  openModal('registerModal');
});

// 요구사항 25: 관리자 전용 소속팀 관리 모달 열기 및 팀 추가 폼
window.openTeamManageModal = async function() {
  openModal('teamManageModal');
  renderTeamManageList();
  try {
    await loadTeams();
    renderTeamManageList();
  } catch (err) {
    console.error('Failed to load teams:', err);
  }
};

const manageTeamsBtnAction = document.getElementById('manageTeamsBtn');
if (manageTeamsBtnAction) {
  manageTeamsBtnAction.addEventListener('click', window.openTeamManageModal);
}
const manageTeamsTopBtnAction = document.getElementById('manageTeamsTopBtn');
if (manageTeamsTopBtnAction) {
  manageTeamsTopBtnAction.addEventListener('click', window.openTeamManageModal);
}

const addTeamForm = document.getElementById('addTeamForm');
if (addTeamForm) {
  addTeamForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const input = document.getElementById('newTeamNameInput');
    const name = input.value.trim();
    if (!name) return;
    if (!currentUser || (!currentUser.is_admin && !currentUser.is_super)) {
      showToast('관리자 권한이 필요합니다.', 'error');
      return;
    }
    try {
      const res = await fetch('/api/teams', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, admin_emp_id: currentUser.emp_id })
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || '소속팀 추가 실패', 'error');
        return;
      }
      input.value = '';
      showToast(data.message);
      await loadTeams();
      renderTeamManageList();
      if (currentMode === 'admin') {
        await loadAdminUserTable();
      }
    } catch (err) {
      console.error(err);
      showToast('통신 오류', 'error');
    }
  });
}


// 모바일 접속 QR 모달 열기
document.getElementById('qrModalOpenBtn').addEventListener('click', async () => {
  try {
    const res = await fetch('/api/system/info');
    const data = await res.json();
    if (res.ok) {
      document.getElementById('qrCodeImg').src = data.qr_code_base64;
      document.getElementById('primaryUrlText').textContent = data.primary_url;
    }
  } catch (err) {
    console.error(err);
  }
  openModal('qrModal');
});

// ===== 14. 요구사항 16, 17: 팀별/인원별 실특근 정산 및 엑셀 내보내기 =====
function initSummaryDateFilter() {
  const startInput = document.getElementById('summaryStartDate');
  const endInput = document.getElementById('summaryEndDate');
  if (startInput && endInput && !startInput.value) {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    startInput.value = `${y}-${m}-01`;
    const lastDay = new Date(y, now.getMonth() + 1, 0).getDate();
    endInput.value = `${y}-${m}-${String(lastDay).padStart(2, '0')}`;
  }

  // 소속팀 셀렉트 옵션 동기화
  const teamSel = document.getElementById('summaryTeamFilter');
  if (teamSel && allUsersCache.length > 0) {
    const curVal = teamSel.value;
    const teams = Array.from(new Set(allUsersCache.map(u => u.team).filter(Boolean))).sort();
    teamSel.innerHTML = '<option value="">전체 소속팀</option>';
    teams.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t;
      teamSel.appendChild(opt);
    });
    teamSel.value = curVal;
  }
}

async function loadSettlementSummary() {
  const startInput = document.getElementById('summaryStartDate');
  const endInput = document.getElementById('summaryEndDate');
  const teamInput = document.getElementById('summaryTeamFilter');

  const startDate = startInput ? startInput.value : '';
  const endDate = endInput ? endInput.value : '';
  const team = teamInput ? teamInput.value : '';

  const params = [];
  if (startDate) params.push(`start_date=${encodeURIComponent(startDate)}`);
  if (endDate) params.push(`end_date=${encodeURIComponent(endDate)}`);
  if (team) params.push(`team=${encodeURIComponent(team)}`);
  if (currentUser) params.push(`admin_emp_id=${encodeURIComponent(currentUser.emp_id)}`);

  let url = '/api/overtimes/summary';
  if (params.length > 0) url += '?' + params.join('&');

  try {
    const res = await fetch(url);
    const data = await res.json();
    if (!res.ok) {
      showToast('정산 데이터 조회 실패', 'error');
      return;
    }

    lastFetchedSummary = data;
    renderSummaryTeamCards(data.team_summary || []);
    renderSummaryUserTable(data.user_summary || []);
  } catch (err) {
    console.error(err);
    showToast('정산 통신 오류 발생', 'error');
  }
}

function renderSummaryTeamCards(teamList) {
  const container = document.getElementById('summaryTeamCards');
  if (!container) return;
  container.innerHTML = '';

  if (!teamList || teamList.length === 0) {
    container.innerHTML = '<div style="grid-column: 1 / -1; color: var(--text-muted); text-align: center; padding: 1.5rem; background:#fff; border-radius:8px; border:1px dashed var(--border-color);">해당 기간의 팀별 특근 실적이 없습니다.</div>';
    return;
  }

  teamList.forEach(tm => {
    const card = document.createElement('div');
    card.className = 'stat-box';
    card.style.borderLeft = '5px solid var(--primary)';
    card.style.background = '#ffffff';
    card.innerHTML = `
      <div style="display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 0.5rem;">
        <span style="font-size: 1.05rem; font-weight: 800; color: var(--text-main);">🏢 ${escapeHtml(tm.team)}</span>
        <span style="font-size: 0.8rem; color: var(--text-muted);">팀원: <b>${tm.member_count}명</b></span>
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 0.35rem; font-size: 0.78rem; color: var(--text-muted); margin-bottom: 0.6rem; background: var(--bg-subtle); padding: 0.6rem; border-radius: 8px;">
        <div>총 신청일수: <b>${tm.total_days}일</b></div>
        <div>제외(대체/법정): <b style="color:var(--text-light);">${tm.excluded_days}일</b></div>
        <div>인정(일반휴일): <b style="color:var(--primary);">${tm.overtime_days}일</b></div>
        <div>대체휴일 사용: <b style="color:var(--warning-text);">${tm.sub_holiday_used}일</b></div>
      </div>
      <div style="display: flex; justify-content: space-between; align-items: center; border-top: 1px dashed var(--border-color); padding-top: 0.5rem;">
        <span style="font-size: 0.84rem; font-weight: 700; color: #166534;">★ 팀 최종 실특근일</span>
        <span style="font-size: 1.35rem; font-weight: 900; color: var(--success);">${tm.actual_overtime_days}일</span>
      </div>
    `;
    container.appendChild(card);
  });
}

function renderSummaryUserTable(userList) {
  const tbody = document.getElementById('summaryUserTbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!userList || userList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align: center; color: var(--text-muted); padding: 2.5rem 1rem;">선택하신 조건에 부합하는 특근 정산 데이터가 없습니다.</td></tr>`;
    return;
  }

  // 정렬 수행
  const sorted = [...userList].sort((a, b) => {
    let va = a[summarySortCol];
    let vb = b[summarySortCol];
    if (typeof va === 'string') {
      return summarySortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
    }
    va = Number(va || 0);
    vb = Number(vb || 0);
    return summarySortAsc ? va - vb : vb - va;
  });

  sorted.forEach(u => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><b>${escapeHtml(u.emp_id)}</b></td>
      <td>${escapeHtml(u.name)}</td>
      <td><span class="status-badge" style="background:#f1f5f9; color:#334155; font-weight:600;">${escapeHtml(u.team)}</span></td>
      <td>${u.total_days}일 <span style="font-size:0.75rem; color:var(--text-light);">(${u.records_count}건)</span></td>
      <td style="color:var(--text-light);">${u.excluded_sub_days}일</td>
      <td style="color:var(--text-light);">${u.excluded_legal_days}일</td>
      <td style="font-weight:700; color:var(--primary);">${u.overtime_days}일</td>
      <td style="font-weight:700; color:var(--warning-text);">${u.sub_holiday_used}일</td>
      <td style="font-weight:900; color:var(--success); background:#ecfdf5; font-size:0.98rem;">★ ${u.actual_overtime_days}일</td>
    `;
    tbody.appendChild(tr);
  });
}

// 정산 테이블 소팅 이벤트 바인딩
document.querySelectorAll('#summaryDataTable th.sortable').forEach(th => {
  th.addEventListener('click', () => {
    const col = th.getAttribute('data-sort');
    if (summarySortCol === col) {
      summarySortAsc = !summarySortAsc;
    } else {
      summarySortCol = col;
      summarySortAsc = true;
    }
    document.querySelectorAll('#summaryDataTable th.sortable .sort-icon').forEach(icon => icon.textContent = '↕');
    th.querySelector('.sort-icon').textContent = summarySortAsc ? '▲' : '▼';
    renderSummaryUserTable(lastFetchedSummary.user_summary || []);
  });
});

const summarySearchBtn = document.getElementById('summarySearchBtn');
if (summarySearchBtn) {
  summarySearchBtn.addEventListener('click', loadSettlementSummary);
}

const summaryResetBtn = document.getElementById('summaryResetBtn');
if (summaryResetBtn) {
  summaryResetBtn.addEventListener('click', () => {
    const sInput = document.getElementById('summaryStartDate');
    const eInput = document.getElementById('summaryEndDate');
    const tInput = document.getElementById('summaryTeamFilter');
    if (sInput) sInput.value = '';
    if (eInput) eInput.value = '';
    if (tInput) tInput.value = '';
    loadSettlementSummary();
  });
}

// 실특근 정산표 엑셀 내보내기 - openpyxl 서버 엔진 전용
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
    const tFilter = document.getElementById('summaryTeamFilter')?.value || '';

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
}

// ===== 15. 외부망 모바일 접속 QR 갱신 (요구사항 7) =====
const btnUpdateCustomQr = document.getElementById('btnUpdateCustomQr');
if (btnUpdateCustomQr) {
  btnUpdateCustomQr.addEventListener('click', async () => {
    const input = document.getElementById('customExternalUrlInput');
    const customUrl = input ? input.value.trim() : '';
    if (!customUrl) {
      showToast('외부 접속 URL을 입력해주세요.', 'warning');
      return;
    }
    try {
      const res = await fetch(`/api/system/info?custom_url=${encodeURIComponent(customUrl)}`);
      const data = await res.json();
      if (res.ok) {
        document.getElementById('qrCodeImg').src = data.qr_code_base64;
        document.getElementById('primaryUrlText').textContent = data.primary_url;
        showToast('외부 접속용 QR 코드가 갱신되었습니다!');
      }
    } catch (e) {
      showToast('QR 갱신 실패', 'error');
    }
  });
}

// ===== 16. 웹상 데이터 스냅샷 저장 및 복원(열기) (요구사항 1) =====
async function saveWebBackup() {
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
}

async function loadBackupList() {
  const tbody = document.getElementById('backupListTbody');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color:var(--text-muted); padding:1rem;">스냅샷 목록 조회 중...</td></tr>';

  try {
    const res = await fetch('/api/backup/list');
    const data = await res.json();
    if (!res.ok || !data.backups || data.backups.length === 0) {
      tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color:var(--text-muted); padding:1.5rem;">저장된 웹 스냅샷이 없습니다.</td></tr>';
      return;
    }

    tbody.innerHTML = '';
    data.backups.forEach(b => {
      const tr = document.createElement('tr');
      const counts = b.counts || {};
      const countStr = `팀 ${counts.teams || 0} / 사원 ${counts.users || 0} / 특근 ${counts.overtimes || 0}`;

      tr.innerHTML = `
        <td><b>${escapeHtml(b.name)}</b></td>
        <td><small style="color:var(--text-muted);">${escapeHtml(b.created_at || '')}</small></td>
        <td><span style="font-size:0.78rem; background:var(--bg-subtle); padding:2px 6px; border-radius:4px;">${countStr}</span></td>
        <td style="text-align:center;">
          <button type="button" class="btn btn-warning btn-xs btn-restore-backup" data-filename="${escapeHtml(b.filename)}" data-name="${escapeHtml(b.name)}">
            📂 열기(복원)
          </button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    tbody.querySelectorAll('.btn-restore-backup').forEach(btn => {
      btn.addEventListener('click', async () => {
        const filename = btn.getAttribute('data-filename');
        const name = btn.getAttribute('data-name');
        if (!confirm(`'${name}' 스냅샷을 열어 현재 데이터베이스를 복원하시겠습니까?\n\n* 현재 데이터가 해당 스냅샷 시점으로 덮어쓰기됩니다.`)) {
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
    });
  } catch (err) {
    console.error(err);
    tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color:var(--danger); padding:1rem;">목록 조회 오류</td></tr>';
  }
}

const backupSaveBtn = document.getElementById('backupSaveBtn');
if (backupSaveBtn) {
  backupSaveBtn.addEventListener('click', async () => {
    openModal('backupModal');
    await loadBackupList();
  });
}

const backupLoadBtn = document.getElementById('backupLoadBtn');
if (backupLoadBtn) {
  backupLoadBtn.addEventListener('click', async () => {
    openModal('backupModal');
    await loadBackupList();
  });
}

const btnExecuteWebBackupSave = document.getElementById('btnExecuteWebBackupSave');
if (btnExecuteWebBackupSave) {
  btnExecuteWebBackupSave.addEventListener('click', saveWebBackup);
}

const btnRefreshBackupList = document.getElementById('btnRefreshBackupList');
if (btnRefreshBackupList) {
  btnRefreshBackupList.addEventListener('click', loadBackupList);
}

// ===== 17. 초기 로드 및 모바일 세션 자동 복원 (요구사항 1, 6) =====
loadTeams();

try {
  const savedSessionStr = localStorage.getItem('overtime_session');
  if (savedSessionStr) {
    const savedUser = JSON.parse(savedSessionStr);
    if (savedUser && savedUser.emp_id) {
      // 서버에서 여전히 유효한 계정인지 확인
      fetch('/api/users/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ emp_id: savedUser.emp_id })
      }).then(res => res.json()).then(data => {
        if (data && data.exists) {
          setUserSession(data.user);
        } else {
          localStorage.removeItem('overtime_session');
        }
      }).catch(() => {
        setUserSession(savedUser);
      });
    }
  }
} catch (e) {
  console.warn('세션 복구 실패:', e);
}

// ===== 18. 슈퍼관리자 전용 보안 감사 및 접속 로그 (신규 요구사항 5) =====

function initAccessLogsFilter() {
  const startInput = document.getElementById('accessFilterStartDate');
  const endInput = document.getElementById('accessFilterEndDate');
  if (startInput && endInput && !startInput.value) {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    startInput.value = `${y}-${m}-01`;
    endInput.value = getTodayStr();
  }
}

async function loadAccessLogs() {
  if (!currentUser) return;
  const tbody = document.getElementById('accessLogsTbody');
  if (tbody) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center; color:var(--text-muted); padding:2rem;">보안 감사 로그를 조회 중입니다...</td></tr>';
  }

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

  try {
    const res = await fetch(`/api/admin/access-logs?${params.join('&')}`);
    const data = await res.json();
    if (!res.ok) {
      if (tbody) tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; color:var(--danger); padding:2rem;">${escapeHtml(data.detail || '조회 권한이 없거나 실패했습니다.')}</td></tr>`;
      return;
    }

    // 1. 상단 통계 카드 갱신
    if (data.stats) {
      const s = data.stats;
      const elTotal = document.getElementById('statAccessTotal');
      const elToday = document.getElementById('statAccessTodayTotal');
      const elFail = document.getElementById('statAccessTodayFailed');
      const elReg = document.getElementById('statAccessTodayRegistered');
      if (elTotal) elTotal.textContent = Number(s.total_logs || 0).toLocaleString();
      if (elToday) elToday.textContent = Number(s.today_total || 0).toLocaleString();
      if (elFail) elFail.textContent = Number(s.today_failed || 0).toLocaleString();
      if (elReg) elReg.textContent = Number(s.today_registered || 0).toLocaleString();
    }

    // 2. 테이블 렌더링
    renderAccessLogs(data.logs || []);
  } catch (err) {
    console.error('Failed to load access logs:', err);
    if (tbody) tbody.innerHTML = '<tr><td colspan="9" style="text-align:center; color:var(--danger); padding:2rem;">통신 오류가 발생했습니다.</td></tr>';
  }
}

function renderAccessLogs(logs) {
  const tbody = document.getElementById('accessLogsTbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!logs || logs.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center; color:var(--text-muted); padding:2.5rem;">조회된 보안 감사 로그가 없습니다.</td></tr>';
    return;
  }

  logs.forEach(item => {
    const tr = document.createElement('tr');
    
    // 상태 배지
    const isSuccess = item.status === 'SUCCESS';
    const statusBadge = isSuccess 
      ? '<span class="badge" style="background:#dcfce7; color:#166534; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:9999px;">✓ 성공</span>'
      : '<span class="badge" style="background:#fee2e2; color:#991B1B; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:9999px;">⚠️ 실패/경고</span>';

    // 액션 구분 배지
    let actionBadge = `<span class="badge" style="background:#f1f5f9; color:#475569; font-size:0.75rem;">${escapeHtml(item.action_type)}</span>`;
    if (item.action_type === 'LOGIN') {
      actionBadge = '<span class="badge" style="background:#e0f2fe; color:#0369a1; font-weight:700; font-size:0.75rem; padding:2px 8px; border-radius:4px;">🔑 로그인</span>';
    } else if (item.action_type === 'REGISTER') {
      actionBadge = '<span class="badge" style="background:#f3e8ff; color:#7e22ce; font-weight:700; font-size:0.75rem; padding:2px 8px; border-radius:4px;">🎉 신규등록</span>';
    } else if (item.action_type === 'LOGOUT') {
      actionBadge = '<span class="badge" style="background:#f1f5f9; color:#64748b; font-size:0.75rem; padding:2px 8px; border-radius:4px;">🚪 로그아웃</span>';
    }

    tr.innerHTML = `
      <td style="text-align:center; color:var(--text-muted); font-size:0.8rem;">${item.id}</td>
      <td style="text-align:center; font-size:0.82rem; white-space:nowrap;">${escapeHtml(item.created_at || '')}</td>
      <td style="text-align:center; font-weight:700;">${escapeHtml(item.emp_id || '-')}</td>
      <td style="text-align:center;">${escapeHtml(item.user_name || '-')}</td>
      <td style="text-align:center;">${actionBadge}</td>
      <td style="text-align:center;">${statusBadge}</td>
      <td style="text-align:center; font-family:monospace; font-size:0.8rem; color:#2563eb;">${escapeHtml(item.ip_address || '-')}</td>
      <td style="font-size:0.84rem; max-width:260px; word-break:break-all;">${escapeHtml(item.details || '-')}</td>
      <td style="font-size:0.75rem; color:var(--text-muted); max-width:220px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${escapeHtml(item.user_agent || '')}">
        ${escapeHtml(item.user_agent || '-')}
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// 엑셀 내보내기 이벤트 바인딩
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
}

// 검색 및 초기화 버튼 바인딩
const accessSearchBtn = document.getElementById('accessSearchBtn');
if (accessSearchBtn) {
  accessSearchBtn.addEventListener('click', loadAccessLogs);
}

const refreshAccessLogsBtn = document.getElementById('refreshAccessLogsBtn');
if (refreshAccessLogsBtn) {
  refreshAccessLogsBtn.addEventListener('click', () => {
    loadAccessLogs();
    showToast('보안 감사 로그를 새로고침했습니다.');
  });
}

const accessResetBtn = document.getElementById('accessResetBtn');
if (accessResetBtn) {
  accessResetBtn.addEventListener('click', () => {
    const sInput = document.getElementById('accessFilterStartDate');
    const eInput = document.getElementById('accessFilterEndDate');
    const aSelect = document.getElementById('accessFilterAction');
    const stSelect = document.getElementById('accessFilterStatus');
    const searchInput = document.getElementById('accessFilterSearch');
    if (sInput) sInput.value = '';
    if (eInput) eInput.value = '';
    if (aSelect) aSelect.value = '';
    if (stSelect) stSelect.value = '';
    if (searchInput) searchInput.value = '';
    loadAccessLogs();
  });
}

// 엔터 키 검색 지원
const accessFilterSearchInput = document.getElementById('accessFilterSearch');
if (accessFilterSearchInput) {
  accessFilterSearchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      loadAccessLogs();
    }
  });
}

// ===== v1.42: 신청자 상세 특근 통계 팝업 (요구사항 7) =====
async function showUserStatsPopup(empId, name) {
  try {
    const res = await fetch(`/api/users/${encodeURIComponent(empId)}/overtime-stats`);
    if (!res.ok) { showToast('통계 조회 실패', 'error'); return; }
    const data = await res.json();
    
    const recentMonths = data.recent_months || [];
    const quarters = data.quarters || [];
    const halves = data.halves || [];
    const byYear = data.by_year || {};

    const tHeader = `<tr style="background:var(--bg-subtle); font-size:0.78rem;">
      <th style="padding:6px 10px; text-align:left; border-bottom:1px solid var(--border-color); white-space:nowrap;">구분</th>
      <th style="padding:6px 6px; text-align:center; border-bottom:1px solid var(--border-color); white-space:nowrap;">일반휴일</th>
      <th style="padding:6px 6px; text-align:center; border-bottom:1px solid var(--border-color); white-space:nowrap;">대체근무</th>
      <th style="padding:6px 6px; text-align:center; border-bottom:1px solid var(--border-color); white-space:nowrap;">법정휴일</th>
      <th style="padding:6px 6px; text-align:center; border-bottom:1px solid var(--border-color); color:#065f46; font-weight:700; white-space:nowrap;">대체휴무</th>
      <th style="padding:6px 8px; text-align:center; border-bottom:1px solid var(--border-color); background:#ecfdf5; color:#047857; font-weight:800; white-space:nowrap;" title="일반특근 - 사전차감 - (대체휴무 - 출장내차감)">★실특근</th>
      <th style="padding:6px 6px; text-align:center; border-bottom:1px solid var(--border-color); white-space:nowrap;">총사전차감</th>
      <th style="padding:6px 6px; text-align:center; border-bottom:1px solid var(--border-color); white-space:nowrap;">출장내차감</th>
      <th style="padding:6px 8px; text-align:center; border-bottom:1px solid var(--border-color); background:#f0f9ff; color:#0369a1; font-weight:800; white-space:nowrap;" title="총사전차감 - 출장내차감">★차감잔여</th>
      <th style="padding:6px 6px; text-align:center; border-bottom:1px solid var(--border-color); white-space:nowrap;">신청건수</th>
    </tr>`;

    const mkRow = (label, b) => {
      const normalH = b['일반휴일'] || 0;
      const subWork = b['대체근무'] || 0;
      const legalH = b['법정휴일'] || 0;
      const subRest = b['대체휴무'] || 0;
      const totalPre = b['사전차감'] || 0;
      const tripPre = b['출장내사전차감'] || 0;
      const actualOt = b['최종실특근'] !== undefined ? b['최종실특근'] : Math.max(0, Math.round((normalH - totalPre - (subRest - tripPre)) * 10) / 10);
      const remainPre = b['사전차감잔여'] !== undefined ? b['사전차감잔여'] : Math.max(0, totalPre - tripPre);
      const recCount = b['total_records'] || 0;

      return `<tr>
        <td style="font-weight:700; color:var(--primary); padding:6px 10px; border-bottom:1px solid var(--border-color); white-space:nowrap;">${escapeHtml(label || '-')}</td>
        <td style="text-align:center; padding:6px 6px; border-bottom:1px solid var(--border-color);">${normalH}일</td>
        <td style="text-align:center; padding:6px 6px; border-bottom:1px solid var(--border-color);">${subWork}일</td>
        <td style="text-align:center; padding:6px 6px; border-bottom:1px solid var(--border-color);">${legalH}일</td>
        <td style="text-align:center; padding:6px 6px; border-bottom:1px solid var(--border-color); font-weight:700; color:#065f46;">${subRest}일</td>
        <td style="text-align:center; padding:6px 8px; border-bottom:1px solid var(--border-color); background:#ecfdf5; font-weight:800; color:#047857;">${actualOt}일</td>
        <td style="text-align:center; padding:6px 6px; border-bottom:1px solid var(--border-color);">${totalPre}회</td>
        <td style="text-align:center; padding:6px 6px; border-bottom:1px solid var(--border-color);">${tripPre}회</td>
        <td style="text-align:center; padding:6px 8px; border-bottom:1px solid var(--border-color); background:#f0f9ff; font-weight:800; color:#0369a1;">${remainPre}회</td>
        <td style="text-align:center; padding:6px 6px; border-bottom:1px solid var(--border-color); color:var(--text-muted);">${recCount}건</td>
      </tr>`;
    };

    const toRows = (obj) => {
      if (!obj) return '';
      if (Array.isArray(obj)) {
        return obj.map(item => mkRow(item.label || item.name, item)).join('');
      }
      return Object.entries(obj).map(([k, v]) => mkRow(v.label || k, v)).join('');
    };

    const monthRows = toRows(recentMonths) || '<tr><td colspan="10" style="text-align:center; padding:1rem; color:var(--text-muted);">데이터 없음</td></tr>';
    const quarterRows = toRows(quarters) || '<tr><td colspan="10" style="text-align:center; padding:1rem; color:var(--text-muted);">데이터 없음</td></tr>';
    const halfRows = toRows(halves) || '<tr><td colspan="10" style="text-align:center; padding:1rem; color:var(--text-muted);">데이터 없음</td></tr>';
    const yearRows = (byYear && Object.keys(byYear).length > 0)
      ? Object.entries(byYear).map(([yr, v]) => mkRow(yr.endsWith('년') ? yr : `${yr}년`, v)).join('')
      : '<tr><td colspan="10" style="text-align:center; padding:1rem; color:var(--text-muted);">데이터 없음</td></tr>';

    const existingOverlay = document.getElementById('userStatsOverlay');
    if (existingOverlay) existingOverlay.remove();

    document.body.insertAdjacentHTML('beforeend', `
      <div id="userStatsOverlay" style="position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.55); z-index:9999; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(2px);" onclick="if(event.target===this)this.remove()">
        <div style="background:var(--bg-card); border-radius:16px; padding:1.5rem 1.75rem; max-width:840px; width:95%; max-height:88vh; overflow-y:auto; box-shadow:0 20px 60px rgba(0,0,0,0.35); border:1.5px solid var(--border-color);">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:1.2rem; padding-bottom:0.75rem; border-bottom:1px solid var(--border-color);">
            <div>
              <h3 style="margin:0; font-size:1.15rem; color:var(--text-main); display:flex; align-items:center; gap:8px;">
                📊 <b>${escapeHtml(name)}</b> <span style="font-size:0.85rem; color:var(--text-muted); font-weight:normal;">(${escapeHtml(empId)} / ${escapeHtml(data.team||'')})</span>
              </h3>
              <p style="font-size:0.78rem; color:var(--text-muted); margin:4px 0 0 0;">
                ★ 실특근 = 일반특근 - 사전차감 - (대체휴무 - 출장내차감)  |  ★ 차감잔여 = 총사전차감 - 출장내차감
              </p>
            </div>
            <button onclick="document.getElementById('userStatsOverlay').remove()" style="border:none; background:none; cursor:pointer; font-size:1.5rem; color:var(--text-muted); line-height:1;">✕</button>
          </div>

          <div style="font-size:0.88rem; font-weight:700; color:var(--primary); margin-bottom:0.4rem; display:flex; align-items:center; gap:6px;">
            🗓️ 최근 3개월 현황 (당월 및 직전 2개월)
          </div>
          <div style="overflow-x:auto; margin-bottom:1.3rem;">
            <table style="width:100%; border-collapse:collapse; font-size:0.82rem; border:1px solid var(--border-color); border-radius:8px; overflow:hidden;">${tHeader}${monthRows}</table>
          </div>

          <div style="font-size:0.88rem; font-weight:700; color:#4338ca; margin-bottom:0.4rem; display:flex; align-items:center; gap:6px;">
            📊 1~4분기별 현황
          </div>
          <div style="overflow-x:auto; margin-bottom:1.3rem;">
            <table style="width:100%; border-collapse:collapse; font-size:0.82rem; border:1px solid var(--border-color); border-radius:8px; overflow:hidden;">${tHeader}${quarterRows}</table>
          </div>

          <div style="font-size:0.88rem; font-weight:700; color:#0d9488; margin-bottom:0.4rem; display:flex; align-items:center; gap:6px;">
            🌗 상반기 / 하반기 현황
          </div>
          <div style="overflow-x:auto; margin-bottom:1.3rem;">
            <table style="width:100%; border-collapse:collapse; font-size:0.82rem; border:1px solid var(--border-color); border-radius:8px; overflow:hidden;">${tHeader}${halfRows}</table>
          </div>

          <div style="font-size:0.88rem; font-weight:700; color:var(--text-main); margin-bottom:0.4rem; display:flex; align-items:center; gap:6px;">
            📆 연도별 누적 현황
          </div>
          <div style="overflow-x:auto; margin-bottom:1.2rem;">
            <table style="width:100%; border-collapse:collapse; font-size:0.82rem; border:1px solid var(--border-color); border-radius:8px; overflow:hidden;">${tHeader}${yearRows}</table>
          </div>

          <div style="text-align:right; margin-top:1.2rem; padding-top:0.75rem; border-top:1px solid var(--border-color);">
            <button onclick="document.getElementById('userStatsOverlay').remove()" class="btn btn-secondary btn-sm" style="padding:6px 16px;">닫기</button>
          </div>
        </div>
      </div>`);
  } catch (err) {
    console.error(err);
    showToast('통계 조회 중 오류 발생', 'error');
  }
}

// ===== v1.42: 항목별 색상 사용자 설정 (요구사항 4) =====
const DEFAULT_CHIP_COLORS = {
  normal: '#b45309',
  subwork: '#0369a1',
  legal: '#b91c1c',
  subrest: '#059669',
  prededuct: '#0284c7',
  bonus: '#7c3aed'
};

function hexToRgba(hex, alpha) {
  if (!hex || hex.length < 7) return hex;
  const r = parseInt(hex.slice(1, 3), 16) || 0;
  const g = parseInt(hex.slice(3, 5), 16) || 0;
  const b = parseInt(hex.slice(5, 7), 16) || 0;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function applyCustomColors(colors) {
  const c = Object.assign({}, DEFAULT_CHIP_COLORS, colors || {});
  let styleEl = document.getElementById('custom-chip-colors-style');
  if (!styleEl) {
    styleEl = document.createElement('style');
    styleEl.id = 'custom-chip-colors-style';
    document.head.appendChild(styleEl);
  }

  styleEl.innerHTML = `
    :root {
      --chip-normal-color: ${c.normal};
      --chip-normal-bg: ${hexToRgba(c.normal, 0.14)};
      --chip-normal-text: ${c.normal};
      --chip-normal-border: ${hexToRgba(c.normal, 0.38)};

      --chip-subwork-color: ${c.subwork};
      --chip-subwork-bg: ${hexToRgba(c.subwork, 0.14)};
      --chip-subwork-text: ${c.subwork};
      --chip-subwork-border: ${hexToRgba(c.subwork, 0.38)};

      --chip-legal-color: ${c.legal};
      --chip-legal-bg: ${hexToRgba(c.legal, 0.14)};
      --chip-legal-text: ${c.legal};
      --chip-legal-border: ${hexToRgba(c.legal, 0.38)};

      --chip-subrest-color: ${c.subrest};
      --chip-subrest-bg: ${hexToRgba(c.subrest, 0.15)};
      --chip-subrest-text: ${c.subrest};
      --chip-subrest-border: ${hexToRgba(c.subrest, 0.38)};

      --chip-prededuct-color: ${c.prededuct};
      --chip-prededuct-bg: ${hexToRgba(c.prededuct, 0.14)};
      --chip-prededuct-text: ${c.prededuct};
      --chip-prededuct-border: ${c.prededuct};

      --chip-bonus-color: ${c.bonus};
      --chip-bonus-bg: ${hexToRgba(c.bonus, 0.14)};
      --chip-bonus-text: ${c.bonus};
      --chip-bonus-border: ${hexToRgba(c.bonus, 0.38)};
    }
  `;

  // 범례 배지 즉시 스타일 반영
  const legNormal = document.getElementById('legendChipNormal');
  if (legNormal) {
    legNormal.style.background = hexToRgba(c.normal, 0.14);
    legNormal.style.color = c.normal;
    legNormal.style.border = `1px solid ${hexToRgba(c.normal, 0.38)}`;
  }
  const legSubWork = document.getElementById('legendChipSubWork');
  if (legSubWork) {
    legSubWork.style.background = hexToRgba(c.subwork, 0.14);
    legSubWork.style.color = c.subwork;
    legSubWork.style.border = `1px solid ${hexToRgba(c.subwork, 0.38)}`;
  }
  const legLegal = document.getElementById('legendChipLegal');
  if (legLegal) {
    legLegal.style.background = hexToRgba(c.legal, 0.14);
    legLegal.style.color = c.legal;
    legLegal.style.border = `1px solid ${hexToRgba(c.legal, 0.38)}`;
  }
  const legSubRest = document.getElementById('legendChipSubRest');
  if (legSubRest) {
    legSubRest.style.background = hexToRgba(c.subrest, 0.15);
    legSubRest.style.color = c.subrest;
    legSubRest.style.border = `1px solid ${hexToRgba(c.subrest, 0.38)}`;
  }
  const legPre = document.getElementById('legendChipPreDeduct');
  if (legPre) {
    legPre.style.background = hexToRgba(c.prededuct, 0.14);
    legPre.style.color = c.prededuct;
    legPre.style.border = `1px solid ${hexToRgba(c.prededuct, 0.38)}`;
  }
  const legBonus = document.getElementById('legendChipBonus');
  if (legBonus) {
    legBonus.style.background = hexToRgba(c.bonus, 0.14);
    legBonus.style.color = c.bonus;
    legBonus.style.border = `1px solid ${hexToRgba(c.bonus, 0.38)}`;
  }
}

function initColorSettings() {
  let savedColors = null;
  try {
    const raw = localStorage.getItem('customChipColors');
    if (raw) savedColors = JSON.parse(raw);
  } catch (e) {
    console.warn('[customChipColors load failed]', e);
  }

  applyCustomColors(savedColors);

  const openBtn = document.getElementById('openColorSettingBtn');
  if (openBtn) {
    openBtn.addEventListener('click', () => {
      let current = DEFAULT_CHIP_COLORS;
      try {
        const raw = localStorage.getItem('customChipColors');
        if (raw) current = Object.assign({}, DEFAULT_CHIP_COLORS, JSON.parse(raw));
      } catch (e) {}

      const pNormal = document.getElementById('colorPickerNormal');
      const pSubWork = document.getElementById('colorPickerSubWork');
      const pLegal = document.getElementById('colorPickerLegal');
      const pSubRest = document.getElementById('colorPickerSubRest');
      const pPre = document.getElementById('colorPickerPreDeduct');
      const pBonus = document.getElementById('colorPickerBonus');

      if (pNormal) pNormal.value = current.normal || DEFAULT_CHIP_COLORS.normal;
      if (pSubWork) pSubWork.value = current.subwork || DEFAULT_CHIP_COLORS.subwork;
      if (pLegal) pLegal.value = current.legal || DEFAULT_CHIP_COLORS.legal;
      if (pSubRest) pSubRest.value = current.subrest || DEFAULT_CHIP_COLORS.subrest;
      if (pPre) pPre.value = current.prededuct || DEFAULT_CHIP_COLORS.prededuct;
      if (pBonus) pBonus.value = current.bonus || DEFAULT_CHIP_COLORS.bonus;

      openModal('colorSettingModal');
    });
  }

  const saveBtn = document.getElementById('saveColorsBtn');
  if (saveBtn) {
    saveBtn.addEventListener('click', () => {
      const colors = {
        normal: document.getElementById('colorPickerNormal')?.value || DEFAULT_CHIP_COLORS.normal,
        subwork: document.getElementById('colorPickerSubWork')?.value || DEFAULT_CHIP_COLORS.subwork,
        legal: document.getElementById('colorPickerLegal')?.value || DEFAULT_CHIP_COLORS.legal,
        subrest: document.getElementById('colorPickerSubRest')?.value || DEFAULT_CHIP_COLORS.subrest,
        prededuct: document.getElementById('colorPickerPreDeduct')?.value || DEFAULT_CHIP_COLORS.prededuct,
        bonus: document.getElementById('colorPickerBonus')?.value || DEFAULT_CHIP_COLORS.bonus
      };

      try {
        localStorage.setItem('customChipColors', JSON.stringify(colors));
      } catch (e) {}

      applyCustomColors(colors);
      closeModal('colorSettingModal');
      showToast('🎨 항목별 맞춤 색상이 저장 및 적용되었습니다!');
    });
  }

  const resetBtn = document.getElementById('resetColorsBtn');
  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      try {
        localStorage.removeItem('customChipColors');
      } catch (e) {}

      const pNormal = document.getElementById('colorPickerNormal');
      const pSubWork = document.getElementById('colorPickerSubWork');
      const pLegal = document.getElementById('colorPickerLegal');
      const pSubRest = document.getElementById('colorPickerSubRest');
      const pPre = document.getElementById('colorPickerPreDeduct');
      const pBonus = document.getElementById('colorPickerBonus');

      if (pNormal) pNormal.value = DEFAULT_CHIP_COLORS.normal;
      if (pSubWork) pSubWork.value = DEFAULT_CHIP_COLORS.subwork;
      if (pLegal) pLegal.value = DEFAULT_CHIP_COLORS.legal;
      if (pSubRest) pSubRest.value = DEFAULT_CHIP_COLORS.subrest;
      if (pPre) pPre.value = DEFAULT_CHIP_COLORS.prededuct;
      if (pBonus) pBonus.value = DEFAULT_CHIP_COLORS.bonus;

      applyCustomColors(DEFAULT_CHIP_COLORS);
      closeModal('colorSettingModal');
      showToast('색상 설정이 기본값으로 초기화되었습니다.');
    });
  }
}
initColorSettings();
window.showUserStatsPopup = showUserStatsPopup;

// ===== v1.41: user_preferences 자동완성 =====
async function loadUserPreferences() {
  if (!currentUser) return;
  try {
    const res = await fetch(`/api/users/${encodeURIComponent(currentUser.emp_id)}/preferences`);
    if (!res.ok) return;
    const data = await res.json();
    if (!data.exists) return;
    const prefs = data.preferences || {};
    const pNoEl = document.getElementById('projectNo');
    const pLocEl = document.getElementById('location');
    if (pNoEl && !pNoEl.value && prefs.last_project_no) pNoEl.value = prefs.last_project_no;
    if (pLocEl && !pLocEl.value && prefs.last_location) pLocEl.value = prefs.last_location;
  } catch (e) {
    console.warn('[preferences load]', e);
  }
}

// 특근 신청 버튼 클릭 시 선호 정보 자동완성
const applyOvertimeBtn = document.getElementById('applyOvertimeBtn');
if (applyOvertimeBtn) {
  applyOvertimeBtn.addEventListener('click', () => {
    setTimeout(loadUserPreferences, 50);
  });
}
