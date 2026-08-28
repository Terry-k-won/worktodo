/**
 * Life & Task Dashboard Application Logic
 * Integrates LocalStorage for instant UI response and Google Apps Script Web App for cloud database sync.
 */

// Global Configuration & Constants
const STORAGE_KEY_ITEMS = 'life_dashboard_items_v2';
const STORAGE_KEY_URL = 'life_dashboard_script_url';

// Default Google Apps Script URL (Embedded User's Spreadsheet Web App)
const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyOJh2aHm1jfjHs3YZUbZFGK7gkl-cMqKJVMxL_dm4zBAl36eAegtLCFWnq_XW2KYHl0w/exec';
let googleScriptUrl = localStorage.getItem(STORAGE_KEY_URL) || DEFAULT_SCRIPT_URL;
if (!localStorage.getItem(STORAGE_KEY_URL)) {
  localStorage.setItem(STORAGE_KEY_URL, DEFAULT_SCRIPT_URL);
}

// App State
let appData = {
  items: [] // List of item objects
};

let currentFilter = 'all';
let currentArchiveTab = 'completed'; // 'completed' or 'deleted'

// Category Name Mapping
const CATEGORY_NAMES = {
  today: '오늘 할 일',
  urgent: '당장 쌓인 과제',
  longterm: '장기 목표',
  skills: '개발할 능력',
  weakness: '나의 부족함'
};

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
  loadData();
  setupEventListeners();
  renderAll();

  // If Script URL is configured, auto sync on launch
  if (googleScriptUrl) {
    document.getElementById('apps-script-url').value = googleScriptUrl;
    syncWithGoogleSheet();
  } else {
    updateSyncBadge('local', '로컬 저장소 모드');
  }
});

/* ==========================================================================
   LocalStorage & Data Management
   ========================================================================== */

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_ITEMS);
    if (raw) {
      appData.items = JSON.parse(raw);
    } else {
      // Demo initial data if fresh start
      appData.items = getInitialDemoData();
      saveData();
    }
  } catch (err) {
    console.error('Failed to load data from LocalStorage:', err);
    appData.items = [];
  }
}

function saveData() {
  try {
    localStorage.setItem(STORAGE_KEY_ITEMS, JSON.stringify(appData.items));
  } catch (err) {
    console.error('Failed to save data to LocalStorage:', err);
  }
}

function getPriorityValue(prio) {
  if (!prio) return 2;
  const p = String(prio).toLowerCase().trim();
  if (p === 'high' || p === '높음' || p.includes('높음') || p === '1') return 1;
  if (p === 'medium' || p === '보통' || p.includes('보통') || p === '2') return 2;
  if (p === 'low' || p === '낮음' || p.includes('낮음') || p === '3') return 3;
  return 2;
}

function getInitialDemoData() {
  const now = new Date().toISOString();
  const todayStr = new Date().toISOString().split('T')[0];

  return [
    // 1. 오늘 할 일 (Today)
    {
      id: 'demo_today_1',
      category: 'today',
      content: '🔥 긴급 업무 서류 검토 및 최종 제출',
      priority: 'high',
      dueDate: todayStr,
      note: '우선순위 높음 (최상단 노출)',
      status: 'active',
      createdAt: now,
      updatedAt: now
    },
    {
      id: 'demo_today_2',
      category: 'today',
      content: '⚡ 주간 프로젝트 일정 팀원 공유',
      priority: 'medium',
      dueDate: todayStr,
      note: '우선순위 보통',
      status: 'active',
      createdAt: now,
      updatedAt: now
    },
    {
      id: 'demo_today_3',
      category: 'today',
      content: '🌱 책상 정돈 및 데스크탑 파일 정리',
      priority: 'low',
      dueDate: todayStr,
      note: '우선순위 낮음',
      status: 'active',
      createdAt: now,
      updatedAt: now
    },

    // 2. 당장 내게 쌓여있는 과제 (Urgent)
    {
      id: 'demo_urgent_1',
      category: 'urgent',
      content: '🔥 구글 앱스 스크립트 DB 연동 및 배포 테스트',
      priority: 'high',
      dueDate: '',
      note: '시급한 핵심 개발 과제',
      status: 'active',
      createdAt: now,
      updatedAt: now
    },
    {
      id: 'demo_urgent_2',
      category: 'urgent',
      content: '⚡ 모바일 웹 UI 반응형 레이아웃 점검',
      priority: 'medium',
      dueDate: '',
      note: '우선순위 보통',
      status: 'active',
      createdAt: now,
      updatedAt: now
    },

    // 3. 장기 목표 (Longterm)
    {
      id: 'demo_long_1',
      category: 'longterm',
      content: '🔥 매일 1시간 개인 역량 강화 공부 습관화',
      priority: 'high',
      dueDate: '',
      note: '장기 성장 핵심 목표',
      status: 'active',
      createdAt: now,
      updatedAt: now
    },
    {
      id: 'demo_long_2',
      category: 'longterm',
      content: '🌱 블로그 기술 포스팅 월 2회 작성',
      priority: 'low',
      dueDate: '',
      note: '우선순위 낮음',
      status: 'active',
      createdAt: now,
      updatedAt: now
    },

    // 4. 개발할 능력 (Skills)
    {
      id: 'demo_skills_1',
      category: 'skills',
      content: '🔥 모바일 프론트엔드 UI & UX 디자인 능력',
      priority: 'high',
      dueDate: '',
      note: '핵심 역량',
      status: 'active',
      createdAt: now,
      updatedAt: now
    },

    // 5. 나의 부족함 (Weakness)
    {
      id: 'demo_weak_1',
      category: 'weakness',
      content: '🔥 미루는 습관 줄이고 실시간 피드백 반영하기',
      priority: 'high',
      dueDate: '',
      note: '최우선 개선 과제',
      status: 'active',
      createdAt: now,
      updatedAt: now
    }
  ];
}

/* ==========================================================================
   Event Listeners Setup
   ========================================================================== */

function setupEventListeners() {
  // Navigation Filter Tabs
  document.querySelectorAll('.category-nav .nav-tab').forEach(tab => {
    tab.addEventListener('click', (e) => {
      document.querySelectorAll('.category-nav .nav-tab').forEach(t => t.classList.remove('active'));
      const target = e.currentTarget.getAttribute('data-target');
      e.currentTarget.classList.add('active');
      filterDashboardView(target);
    });
  });

  // Action Buttons in Header
  document.getElementById('btn-sync').addEventListener('click', () => {
    if (!googleScriptUrl) {
      openSettingsModal();
      showToast('구글 앱스 스크립트 URL을 먼저 등록해주세요.', 'warning');
      return;
    }
    syncWithGoogleSheet(true);
  });

  document.getElementById('btn-archive').addEventListener('click', openArchiveModal);
  document.getElementById('btn-settings').addEventListener('click', openSettingsModal);
  document.getElementById('fab-add').addEventListener('click', () => openAddModal('today'));

  // Modals Close Buttons
  document.getElementById('btn-close-item-modal').addEventListener('click', closeAddModal);
  document.getElementById('btn-cancel-item').addEventListener('click', closeAddModal);

  document.getElementById('btn-close-archive-modal').addEventListener('click', closeArchiveModal);
  document.getElementById('btn-close-archive-bottom').addEventListener('click', closeArchiveModal);

  document.getElementById('btn-close-settings-modal').addEventListener('click', closeSettingsModal);
  document.getElementById('btn-cancel-settings').addEventListener('click', closeSettingsModal);

  // Auto-fill today's date when 'today' category is selected in modal
  document.getElementById('item-category').addEventListener('change', (e) => {
    if (e.target.value === 'today') {
      const dateInput = document.getElementById('item-duedate');
      if (!dateInput.value) {
        dateInput.value = getTodayDateString();
      }
    }
  });

  // Close modals on overlay click
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) {
        overlay.classList.remove('active');
      }
    });
  });
}

// Priority Rank Mapping for Sorting
const PRIORITY_ORDER = { high: 1, medium: 2, low: 3 };

// Category Icons & Meta
const CATEGORY_META = {
  today: { name: '오늘 할 일', icon: 'ri-sun-fill', iconClass: 'today-icon' },
  urgent: { name: '당장 쌓인 과제', icon: 'ri-fire-fill', iconClass: 'urgent-icon' },
  longterm: { name: '장기 목표', icon: 'ri-flag-fill', iconClass: 'longterm-icon' },
  skills: { name: '개발할 능력', icon: 'ri-tools-fill', iconClass: 'skills-icon' },
  weakness: { name: '나의 부족함', icon: 'ri-shield-flash-fill', iconClass: 'weakness-icon' }
};

/* ==========================================================================
   UI Rendering
   ========================================================================== */

function renderAll() {
  renderCategoryLists();
  renderArchiveList();
  updateBadgesAndCounts();
  renderRightSidebar();
  filterDashboardView(currentFilter);
}

function renderCategoryLists() {
  const categories = ['today', 'urgent', 'longterm', 'skills', 'weakness'];

  categories.forEach(cat => {
    const listEl = document.getElementById(`list-${cat}`);
    const emptyEl = document.getElementById(`empty-${cat}`);
    
    // Filter active items for this category & SORT BY PRIORITY (High=1 -> Medium=2 -> Low=3)
    const items = appData.items
      .filter(item => item.category === cat && item.status === 'active')
      .sort((a, b) => {
        const prioA = getPriorityValue(a.priority);
        const prioB = getPriorityValue(b.priority);
        if (prioA !== prioB) return prioA - prioB; // 1 (High) comes before 2 (Medium) comes before 3 (Low)
        
        // Secondary sort: Due Date (Earliest first)
        if (a.dueDate && b.dueDate) return a.dueDate.localeCompare(b.dueDate);
        if (a.dueDate) return -1;
        if (b.dueDate) return 1;

        // Tertiary sort: Latest Created
        return new Date(b.createdAt) - new Date(a.createdAt);
      });

    listEl.innerHTML = '';

    if (items.length === 0) {
      emptyEl.style.display = 'flex';
    } else {
      emptyEl.style.display = 'none';
      items.forEach(item => {
        const itemEl = createCategoryItemElement(item);
        listEl.appendChild(itemEl);
      });
    }
  });
}

function renderRightSidebar() {
  const sidebarListEl = document.getElementById('sidebar-tab-list');
  if (!sidebarListEl) return;

  const categories = ['today', 'urgent', 'longterm', 'skills', 'weakness'];
  
  // Render sidebar list items for categories other than 'all'
  sidebarListEl.innerHTML = '';
  
  categories.forEach(cat => {
    const meta = CATEGORY_META[cat];
    const count = appData.items.filter(i => i.category === cat && i.status === 'active').length;
    const isCurrent = (cat === currentFilter);

    const li = document.createElement('li');
    li.className = `sidebar-tab-item ${isCurrent ? 'active' : ''}`;
    li.onclick = () => switchCategoryView(cat);

    li.innerHTML = `
      <div class="sidebar-tab-title">
        <span class="category-icon ${meta.iconClass}"><i class="${meta.icon}"></i></span>
        <span>${meta.name}</span>
      </div>
      <span class="item-count">${count}</span>
    `;

    sidebarListEl.appendChild(li);
  });
}

function filterDashboardView(target) {
  currentFilter = target;
  const wrapper = document.getElementById('dashboard-wrapper');
  const cards = document.querySelectorAll('.category-card');

  // Update top navigation active class
  document.querySelectorAll('.category-nav .nav-tab').forEach(tab => {
    const t = tab.getAttribute('data-target');
    tab.classList.toggle('active', t === target);
  });

  if (target === 'all') {
    wrapper.className = 'dashboard-wrapper mode-all';
    cards.forEach(card => {
      card.classList.remove('focused');
      card.style.display = 'flex';
    });
  } else {
    wrapper.className = 'dashboard-wrapper mode-focused';
    cards.forEach(card => {
      const cat = card.getAttribute('data-category');
      if (cat === target) {
        card.classList.add('focused');
        card.style.display = 'flex';
      } else {
        card.classList.remove('focused');
        card.style.display = 'none';
      }
    });
  }

  renderRightSidebar();
}

function switchCategoryView(target) {
  filterDashboardView(target);
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function createCategoryItemElement(item) {
  const li = document.createElement('li');
  li.className = 'task-item';
  li.id = `item-element-${item.id}`;

  const priorityVal = getPriorityValue(item.priority);
  const priorityLabel = { 1: '🔥 높음', 2: '⚡ 보통', 3: '🌱 낮음' }[priorityVal] || '⚡ 보통';
  const priorityClass = { 1: 'prio-high', 2: 'prio-medium', 3: 'prio-low' }[priorityVal] || 'prio-medium';

  const dueDateHtml = item.dueDate 
    ? `<span class="due-date"><i class="ri-calendar-line"></i> ${item.dueDate}</span>` 
    : '';

  const noteHtml = item.note 
    ? `<div class="task-note">${escapeHtml(item.note)}</div>` 
    : '';

  li.innerHTML = `
    <div class="task-checkbox-wrapper" onclick="toggleItemComplete('${item.id}', event)" title="완료 체크">
      <div class="task-checkbox">
        <i class="ri-check-line"></i>
      </div>
    </div>
    <div class="task-content-area" onclick="openEditModal('${item.id}')">
      <div class="task-main-text">${escapeHtml(item.content)}</div>
      <div class="task-meta">
        <span class="prio-badge ${priorityClass}">${priorityLabel}</span>
        ${dueDateHtml}
      </div>
      ${noteHtml}
    </div>
    <div class="task-actions">
      <button class="item-action-btn" onclick="openEditModal('${item.id}')" title="수정">
        <i class="ri-edit-line"></i>
      </button>
      <button class="item-action-btn delete" onclick="deleteItemToTrash('${item.id}', event)" title="삭제 (휴지통으로)">
        <i class="ri-delete-bin-line"></i>
      </button>
    </div>
  `;

  return li;
}

function renderArchiveList() {
  const archiveListEl = document.getElementById('archive-list');
  const emptyArchiveEl = document.getElementById('empty-archive');

  const filteredItems = appData.items.filter(item => item.status === currentArchiveTab);
  archiveListEl.innerHTML = '';

  if (filteredItems.length === 0) {
    emptyArchiveEl.style.display = 'flex';
  } else {
    emptyArchiveEl.style.display = 'none';
    filteredItems.forEach(item => {
      const li = document.createElement('li');
      li.className = `archive-item ${item.status}`;
      
      const catName = CATEGORY_NAMES[item.category] || item.category;

      li.innerHTML = `
        <div class="archive-item-info">
          <div class="archive-item-title">${escapeHtml(item.content)}</div>
          <div class="archive-meta">
            [${catName}] ${item.status === 'completed' ? '완료됨' : '삭제됨'} • ${formatDate(item.updatedAt || item.createdAt)}
          </div>
        </div>
        <div class="archive-actions">
          <button class="btn-restore" onclick="restoreItem('${item.id}')" title="원래 카테고리로 복원">
            <i class="ri-restart-line"></i> 복원
          </button>
          <button class="btn-perm-delete" onclick="permanentlyDeleteItem('${item.id}')" title="영구 삭제">
            <i class="ri-close-circle-line"></i> 삭제
          </button>
        </div>
      `;
      archiveListEl.appendChild(li);
    });
  }
}

function updateBadgesAndCounts() {
  const categories = ['today', 'urgent', 'longterm', 'skills', 'weakness'];
  
  categories.forEach(cat => {
    const count = appData.items.filter(i => i.category === cat && i.status === 'active').length;
    document.getElementById(`count-${cat}`).textContent = count;
  });

  const completedCount = appData.items.filter(i => i.status === 'completed').length;
  const deletedCount = appData.items.filter(i => i.status === 'deleted').length;
  const totalArchiveCount = completedCount + deletedCount;

  document.getElementById('archive-count-badge').textContent = totalArchiveCount;
  document.getElementById('archive-completed-count').textContent = completedCount;
  document.getElementById('archive-deleted-count').textContent = deletedCount;
}

function filterDashboardView(target) {
  currentFilter = target;
  const cards = document.querySelectorAll('.category-card');

  cards.forEach(card => {
    const cat = card.getAttribute('data-category');
    if (target === 'all' || target === cat) {
      card.style.display = 'flex';
    } else {
      card.style.display = 'none';
    }
  });
}

/* ==========================================================================
   CRUD Actions
   ========================================================================== */

function handleItemFormSubmit(e) {
  e.preventDefault();
  const itemId = document.getElementById('item-id').value;
  const category = document.getElementById('item-category').value;
  const content = document.getElementById('item-content').value.trim();
  const priority = document.getElementById('item-priority').value;
  const dueDate = document.getElementById('item-duedate').value;
  const note = document.getElementById('item-note').value.trim();

  if (!content) return;

  const now = new Date().toISOString();

  if (itemId) {
    // Update existing
    const itemIndex = appData.items.findIndex(i => i.id === itemId);
    if (itemIndex !== -1) {
      appData.items[itemIndex] = {
        ...appData.items[itemIndex],
        category,
        content,
        priority,
        dueDate,
        note,
        updatedAt: now
      };
      showToast('항목이 수정되었습니다.', 'info');
    }
  } else {
    // Create new item
    const newItem = {
      id: 'item_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
      category,
      content,
      priority,
      dueDate,
      note,
      status: 'active',
      createdAt: now,
      updatedAt: now
    };
    appData.items.unshift(newItem);
    showToast('새 항목이 등록되었습니다.', 'success');
  }

  saveData();
  renderAll();
  closeAddModal();

  // Background Cloud Sync
  if (googleScriptUrl) {
    syncWithGoogleSheet();
  }
}

function toggleItemComplete(id, event) {
  if (event) event.stopPropagation();

  const itemIndex = appData.items.findIndex(i => i.id === id);
  if (itemIndex === -1) return;

  const itemEl = document.getElementById(`item-element-${id}`);
  
  // Apply completion animation
  if (itemEl) {
    itemEl.classList.add('completing');
  }

  setTimeout(() => {
    appData.items[itemIndex].status = 'completed';
    appData.items[itemIndex].updatedAt = new Date().toISOString();
    
    saveData();
    renderAll();
    showToast('항목이 완료되어 보관함으로 이동했습니다.', 'success', () => restoreItem(id));

    if (googleScriptUrl) {
      syncWithGoogleSheet();
    }
  }, 350);
}

function deleteItemToTrash(id, event) {
  if (event) event.stopPropagation();

  const itemIndex = appData.items.findIndex(i => i.id === id);
  if (itemIndex === -1) return;

  const itemEl = document.getElementById(`item-element-${id}`);
  if (itemEl) {
    itemEl.classList.add('completing');
  }

  setTimeout(() => {
    appData.items[itemIndex].status = 'deleted';
    appData.items[itemIndex].updatedAt = new Date().toISOString();

    saveData();
    renderAll();
    showToast('항목이 휴지통으로 이동했습니다.', 'info', () => restoreItem(id));

    if (googleScriptUrl) {
      syncWithGoogleSheet();
    }
  }, 350);
}

function restoreItem(id) {
  const itemIndex = appData.items.findIndex(i => i.id === id);
  if (itemIndex === -1) return;

  appData.items[itemIndex].status = 'active';
  appData.items[itemIndex].updatedAt = new Date().toISOString();

  saveData();
  renderAll();
  showToast('항목이 원래 카테고리로 복원되었습니다.', 'success');

  if (googleScriptUrl) {
    syncWithGoogleSheet();
  }
}

function permanentlyDeleteItem(id) {
  appData.items = appData.items.filter(i => i.id !== id);

  saveData();
  renderAll();
  showToast('항목이 영구적으로 삭제되었습니다.', 'warning');

  if (googleScriptUrl) {
    syncWithGoogleSheet();
  }
}

function confirmClearArchive() {
  const targetStatus = currentArchiveTab;
  const count = appData.items.filter(i => i.status === targetStatus).length;
  
  if (count === 0) {
    showToast('비울 항목이 없습니다.', 'info');
    return;
  }

  if (confirm(`정말로 ${targetStatus === 'completed' ? '완료된 항목' : '휴지통의 항목'} ${count}개를 모두 영구 삭제하시겠습니까?`)) {
    appData.items = appData.items.filter(i => i.status !== targetStatus);
    saveData();
    renderAll();
    showToast('선택한 항목들이 모두 삭제되었습니다.', 'success');

    if (googleScriptUrl) {
      syncWithGoogleSheet();
    }
  }
}

/* Helper: Get Today's Date String in YYYY-MM-DD */
function getTodayDateString() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/* ==========================================================================
   Modal Dialog Controllers
   ========================================================================== */

function openAddModal(category = 'today') {
  document.getElementById('modal-item-title').textContent = '새 항목 추가';
  document.getElementById('item-id').value = '';
  document.getElementById('item-category').value = category;
  document.getElementById('item-content').value = '';
  document.getElementById('item-priority').value = 'medium';
  
  // Default date to today if category is 'today'
  if (category === 'today') {
    document.getElementById('item-duedate').value = getTodayDateString();
  } else {
    document.getElementById('item-duedate').value = '';
  }

  document.getElementById('item-note').value = '';

  document.getElementById('modal-item').classList.add('active');
  setTimeout(() => document.getElementById('item-content').focus(), 150);
}

function openEditModal(id) {
  const item = appData.items.find(i => i.id === id);
  if (!item) return;

  document.getElementById('modal-item-title').textContent = '항목 수정';
  document.getElementById('item-id').value = item.id;
  document.getElementById('item-category').value = item.category;
  document.getElementById('item-content').value = item.content;
  document.getElementById('item-priority').value = item.priority || 'medium';
  document.getElementById('item-duedate').value = item.dueDate || '';
  document.getElementById('item-note').value = item.note || '';

  document.getElementById('modal-item').classList.add('active');
}

function closeAddModal() {
  document.getElementById('modal-item').classList.remove('active');
}

function openArchiveModal() {
  renderArchiveList();
  document.getElementById('modal-archive').classList.add('active');
}

function closeArchiveModal() {
  document.getElementById('modal-archive').classList.remove('active');
}

function switchArchiveTab(tab) {
  currentArchiveTab = tab;
  document.getElementById('tab-archived-completed').classList.toggle('active', tab === 'completed');
  document.getElementById('tab-archived-deleted').classList.toggle('active', tab === 'deleted');
  renderArchiveList();
}

function openSettingsModal() {
  document.getElementById('apps-script-url').value = googleScriptUrl;
  document.getElementById('modal-settings').classList.add('active');
}

function closeSettingsModal() {
  document.getElementById('modal-settings').classList.remove('active');
}

function saveSettings() {
  const url = document.getElementById('apps-script-url').value.trim();
  googleScriptUrl = url;
  localStorage.setItem(STORAGE_KEY_URL, url);

  showToast('설정이 저장되었습니다.', 'success');
  closeSettingsModal();

  if (googleScriptUrl) {
    syncWithGoogleSheet(true);
  } else {
    updateSyncBadge('local', '로컬 저장소 모드');
  }
}

/* ==========================================================================
   Google Apps Script Sync Backend
   ========================================================================== */

async function syncWithGoogleSheet(manual = false) {
  if (!googleScriptUrl) return;

  updateSyncBadge('syncing', '동기화 중...');

  // 1. Primary Sync method: GET request with URL-encoded JSON payload
  // This bypasses browser CORS preflight blocks completely on all mobile and web devices!
  try {
    const encodedData = encodeURIComponent(JSON.stringify(appData.items));
    const syncUrl = `${googleScriptUrl}?action=sync&data=${encodedData}`;

    const response = await fetch(syncUrl);
    const result = await response.json();

    if (result && result.status === 'success') {
      if (result.items && Array.isArray(result.items)) {
        appData.items = result.items;
        saveData();
        renderAll();
      }
      updateSyncBadge('success', '동기화 완료');
      if (manual) showToast('구글 스프레드시트와 실시간 동기화 되었습니다!', 'success');
      return;
    }
  } catch (err) {
    console.warn('GET sync failed, trying POST fallback:', err);
  }

  // 2. Fallback POST sync
  try {
    const payload = JSON.stringify({
      action: 'sync',
      items: appData.items
    });

    const response = await fetch(googleScriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8',
      },
      body: payload
    });

    const result = await response.json();

    if (result && result.status === 'success') {
      if (result.items && Array.isArray(result.items)) {
        appData.items = result.items;
        saveData();
        renderAll();
      }
      updateSyncBadge('success', '동기화 완료');
      if (manual) showToast('구글 스프레드시트와 동기화 되었습니다!', 'success');
    } else {
      throw new Error(result.message || 'Server error');
    }
  } catch (fetchErr) {
    console.error('Cloud Sync failed:', fetchErr);
    updateSyncBadge('error', '동기화 실패');
    if (manual) showToast('스프레드시트 연동 실패. 로컬 데이터로 유지됩니다.', 'error');
  }
}

async function testGoogleConnection() {
  const url = document.getElementById('apps-script-url').value.trim();
  if (!url) {
    alert('테스트할 Web App URL을 입력해주세요.');
    return;
  }

  showToast('연결 상태를 확인하는 중...', 'info');

  try {
    const response = await fetch(`${url}?action=ping`);
    const data = await response.json();
    if (data && data.status === 'ok') {
      alert('✅ 구글 스프레드시트와 성공적으로 연결되었습니다!');
    } else {
      alert('⚠️ 응답을 받았으나 스프레드시트 설정 확인이 필요합니다.');
    }
  } catch (err) {
    alert('❌ 연결 실패! Web App URL 또는 배포 설정(액세스 권한: 모든 사용자)을 확인해주세요.\n\n오류: ' + err.message);
  }
}

function exportDataJSON() {
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(appData, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute("href", dataStr);
  downloadAnchor.setAttribute("download", `life_dashboard_backup_${new Date().toISOString().split('T')[0]}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
  showToast('데이터 백업 JSON 파일이 다운로드되었습니다.', 'success');
}

/* ==========================================================================
   Utility Helpers
   ========================================================================== */

function updateSyncBadge(status, text) {
  const badge = document.getElementById('sync-status');
  badge.className = `sync-badge ${status}`;
  
  let iconHtml = '<i class="ri-cloud-line"></i>';
  if (status === 'syncing') iconHtml = '<i class="ri-refresh-line spin-icon"></i>';
  if (status === 'success') iconHtml = '<i class="ri-cloud-line"></i>';
  if (status === 'error') iconHtml = '<i class="ri-cloud-off-line"></i>';
  if (status === 'local') iconHtml = '<i class="ri-hard-drive-2-line"></i>';

  badge.innerHTML = `${iconHtml} <span class="sync-text">${text}</span>`;
}

function showToast(message, type = 'info', undoAction = null) {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;

  let icon = 'ri-information-fill';
  if (type === 'success') icon = 'ri-checkbox-circle-fill';
  if (type === 'warning') icon = 'ri-alert-fill';
  if (type === 'error') icon = 'ri-error-warning-fill';

  let undoButtonHtml = '';
  if (undoAction) {
    undoButtonHtml = `<button style="margin-left: 10px; background: rgba(255,255,255,0.2); border: none; color: #fff; padding: 2px 8px; border-radius: 4px; font-size: 0.75rem; cursor: pointer;" id="btn-undo-toast">되돌리기</button>`;
  }

  toast.innerHTML = `<i class="${icon}"></i> <span>${escapeHtml(message)}</span> ${undoButtonHtml}`;
  container.appendChild(toast);

  if (undoAction) {
    const btnUndo = toast.querySelector('#btn-undo-toast');
    if (btnUndo) {
      btnUndo.addEventListener('click', () => {
        undoAction();
        toast.remove();
      });
    }
  }

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

function formatDate(isoStr) {
  if (!isoStr) return '';
  const d = new Date(isoStr);
  if (isNaN(d.getTime())) return isoStr;
  return `${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
}

function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
