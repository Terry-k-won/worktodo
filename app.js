/**
 * Life & Task Dashboard Application Logic
 * Integrates LocalStorage for instant UI response and Google Apps Script Web App for cloud database sync.
 */

// Global Configuration & Constants
const STORAGE_KEY_ITEMS = 'life_dashboard_items_v2';
const STORAGE_KEY_URL = 'life_dashboard_script_url';
const STORAGE_KEY_CLIENT_ID = 'google_calendar_client_id';
const STORAGE_KEY_AUTOCONNECT = 'gcal_auto_connect';
const STORAGE_KEY_GCAL_TOKEN = 'gcal_access_token';
const STORAGE_KEY_GCAL_EXPIRES = 'gcal_token_expires_at';

// Default Google Apps Script URL & Google Calendar Client ID
const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzA5TKY3pMGv7wjRW6sSopFDXOjEON731j3D25WJzrVwt1ax7P361FqqO2JrYz7mSoTRA/exec';
const DEFAULT_CLIENT_ID = '1078530860351-oh3i9sj5om2a7ka3fi4j2mlebmik76ho.apps.googleusercontent.com';

let googleScriptUrl = DEFAULT_SCRIPT_URL;
let googleClientId = localStorage.getItem(STORAGE_KEY_CLIENT_ID) || DEFAULT_CLIENT_ID;
localStorage.setItem(STORAGE_KEY_URL, DEFAULT_SCRIPT_URL);
localStorage.setItem(STORAGE_KEY_CLIENT_ID, googleClientId);

// App State
let appData = {
  items: [] // List of item objects
};

let currentFilter = 'all';
let isCalendarView = false;  // Separate flag to track calendar view state
let currentArchiveTab = 'completed'; // 'completed' or 'deleted'

// Calendar View State
let currentCalYear = new Date().getFullYear();
let currentCalMonth = new Date().getMonth(); // 0-indexed (0=Jan..11=Dec)

// Google Calendar API State
let gapiInited = false;
let gisInited = false;
let gcalTokenClient = null;
let gcalEvents = []; // Cached Google Calendar events for current month
const GCAL_SCOPE = 'https://www.googleapis.com/auth/calendar.events';
const GCAL_DISCOVERY = 'https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest';

// Google Calendar 11 Standard Color Map
const GCAL_COLOR_MAP = {
  '1':  { border: '#7986cb', bg: '#e8eaf6', text: '#283593' }, // Lavender
  '2':  { border: '#33b679', bg: '#e8f5e9', text: '#1b5e20' }, // Sage
  '3':  { border: '#8e24aa', bg: '#f3e5f5', text: '#4a148c' }, // Grape
  '4':  { border: '#e67c73', bg: '#fbe9e7', text: '#b71c1c' }, // Flamingo
  '5':  { border: '#f6bf26', bg: '#fffde7', text: '#f57f17' }, // Banana
  '6':  { border: '#f4511e', bg: '#fbe9e7', text: '#bf360c' }, // Tangerine
  '7':  { border: '#039be5', bg: '#e1f5fe', text: '#01579b' }, // Peacock
  '8':  { border: '#616161', bg: '#f5f5f5', text: '#212121' }, // Graphite
  '9':  { border: '#3f51b5', bg: '#e8f0fe', text: '#1a56db' }, // Blueberry (Default)
  '10': { border: '#0b8043', bg: '#e8f5e9', text: '#004d40' }, // Basil
  '11': { border: '#d50000', bg: '#ffebee', text: '#b71c1c' }, // Tomato
};

function getGCalEventColor(colorId) {
  return GCAL_COLOR_MAP[String(colorId)] || GCAL_COLOR_MAP['9'];
}

// Category Name Mapping
const CATEGORY_NAMES = {
  today: '오늘 할 일',
  urgent: '당장 쌓인 과제',
  longterm: '장기 목표',
  skills: '개발할 능력',
  weakness: '나의 부족함'
};

// Initialize Application
document.addEventListener('DOMContentLoaded', async () => {
  loadData();
  setupEventListeners();
  renderAll();

  // If Script URL is configured, PULL LATEST MASTER DB FROM GOOGLE SHEETS FIRST!
  if (googleScriptUrl) {
    document.getElementById('apps-script-url').value = googleScriptUrl;
    await fetchLatestFromGoogleSheet();
  } else {
    updateSyncBadge('local', '로컬 저장소 모드');
  }
});

// Auto pull latest data when switching back to this tab
window.addEventListener('focus', () => {
  if (googleScriptUrl) fetchLatestFromGoogleSheet();
});

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible' && googleScriptUrl) {
    fetchLatestFromGoogleSheet();
  }
});

/* ==========================================================================
   LocalStorage & Data Management
   ========================================================================== */

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_ITEMS);
    if (raw) {
      const parsed = JSON.parse(raw);
      appData.items = Array.isArray(parsed) ? parsed.filter(item => item && item.id && !String(item.id).startsWith('demo_')) : [];
    } else {
      appData.items = [];
      saveData();
    }
  } catch (err) {
    console.error('Failed to load data from LocalStorage:', err);
    appData.items = [];
  }
}

function saveData() {
  try {
    // Filter demo items before saving
    if (Array.isArray(appData.items)) {
      appData.items = appData.items.filter(item => item && item.id && !String(item.id).startsWith('demo_'));
    }
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
  return [];
}

/* ==========================================================================
   Google Calendar API Integration
   ========================================================================== */

// Called by gapi.js onload
function gapiLoaded() {
  gapi.load('client', async () => {
    try {
      // init with no args first, then explicitly load Calendar v3
      await gapi.client.init({});
      await gapi.client.load('https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest');
      gapiInited = true;
      console.log('[GCal] gapi + Calendar API loaded OK');
      maybeEnableGCalUI();
    } catch (err) {
      console.error('[GCal] gapi init error:', err);
    }
  });
}

// Called by GIS (accounts.google.com/gsi/client) onload
function gisLoaded() {
  try {
    gcalTokenClient = google.accounts.oauth2.initTokenClient({
      client_id: DEFAULT_CLIENT_ID,
      scope: GCAL_SCOPE,
      callback: '',
    });
    gisInited = true;
    console.log('[GCal] GIS token client initialized OK');
    maybeEnableGCalUI();
  } catch (err) {
    console.error('[GCal] GIS init error:', err);
  }
}

// Helper to save token in LocalStorage
function saveGCalToken(token, expiresInSeconds = 3600) {
  const expiresAt = Date.now() + (parseInt(expiresInSeconds, 10) || 3600) * 1000;
  localStorage.setItem(STORAGE_KEY_GCAL_TOKEN, token);
  localStorage.setItem(STORAGE_KEY_GCAL_EXPIRES, expiresAt.toString());
  localStorage.setItem(STORAGE_KEY_AUTOCONNECT, 'true');
}

// Helper to clear token in LocalStorage
function clearGCalToken() {
  localStorage.removeItem(STORAGE_KEY_GCAL_TOKEN);
  localStorage.removeItem(STORAGE_KEY_GCAL_EXPIRES);
  localStorage.setItem(STORAGE_KEY_AUTOCONNECT, 'false');
  if (gapiInited) gapi.client.setToken('');
}

// Enable buttons once BOTH gapi and GIS are ready + AUTO RESTORE TOKEN IF VALID
function maybeEnableGCalUI() {
  if (!gapiInited || !gisInited) return;

  const savedToken = localStorage.getItem(STORAGE_KEY_GCAL_TOKEN);
  const expiresAt = parseInt(localStorage.getItem(STORAGE_KEY_GCAL_EXPIRES) || '0', 10);

  // 1. Instant Token Restore: If valid token exists in LocalStorage, reuse it immediately!
  if (savedToken && Date.now() < expiresAt - 30000) {
    gapi.client.setToken({ access_token: savedToken });
    console.log('[GCal] Instant auto-login: restored valid token from LocalStorage!');
    updateGCalUI();
    loadCalendarEventsForMonth();
    return;
  }

  // 2. Silent Refresh: If token expired but auto-connect enabled, attempt silent request
  if (localStorage.getItem(STORAGE_KEY_AUTOCONNECT) === 'true' && !isGCalSignedIn()) {
    console.log('[GCal] Token expired, attempting silent background refresh...');
    gcalTokenClient.callback = async (resp) => {
      if (!resp.error && resp.access_token) {
        console.log('[GCal] Silent token refresh success!');
        saveGCalToken(resp.access_token, resp.expires_in);
        gapi.client.setToken({ access_token: resp.access_token });
        updateGCalUI();
        await loadCalendarEventsForMonth();
      } else {
        console.warn('[GCal] Silent token refresh failed:', resp.error);
        clearGCalToken();
        updateGCalUI();
      }
    };
    try {
      gcalTokenClient.requestAccessToken({ prompt: '' });
    } catch (err) {
      console.error('[GCal] Error requesting silent token:', err);
      updateGCalUI();
    }
  } else {
    updateGCalUI();
  }
}

// Safe check: is user currently signed in with a valid token?
function isGCalSignedIn() {
  try {
    if (!gapiInited) return false;
    const token = gapi.client.getToken();
    return token !== null && !!token.access_token;
  } catch {
    return false;
  }
}

// Update the auth area UI based on current state
function updateGCalUI() {
  const statusEl = document.getElementById('gcal-status');
  const signinBtn = document.getElementById('gcal-signin-btn');
  const addEventBtn = document.getElementById('gcal-add-event-btn');
  const signoutBtn = document.getElementById('gcal-signout-btn');
  if (!statusEl || !signinBtn || !signoutBtn) return;

  if (!gapiInited || !gisInited) {
    statusEl.className = 'gcal-status-badge gcal-loading';
    statusEl.innerHTML = '<i class="ri-loader-4-line spin-icon"></i><span>초기화 중...</span>';
    signinBtn.style.display = 'none';
    if (addEventBtn) addEventBtn.style.display = 'none';
    signoutBtn.style.display = 'none';
    return;
  }

  if (isGCalSignedIn()) {
    statusEl.className = 'gcal-status-badge gcal-connected';
    statusEl.innerHTML = '<i class="ri-checkbox-circle-fill"></i><span>구글 캘린더 연동됨</span>';
    signinBtn.style.display = 'none';
    if (addEventBtn) addEventBtn.style.display = 'flex';
    signoutBtn.style.display = 'flex';
  } else {
    statusEl.className = 'gcal-status-badge gcal-disconnected';
    statusEl.innerHTML = '<i class="ri-calendar-close-line"></i><span>구글 캘린더 미연동</span>';
    signinBtn.style.display = 'flex';
    if (addEventBtn) addEventBtn.style.display = 'none';
    signoutBtn.style.display = 'none';
  }
}

// OAuth sign-in — opens popup, then loads events on success
function handleGoogleSignIn() {
  if (!gapiInited || !gisInited || !gcalTokenClient) {
    showToast('Google API가 아직 준비 중입니다. 잠시 후 다시 시도해주세요.', 'warning');
    return;
  }
  gcalTokenClient.callback = async (resp) => {
    if (resp.error) {
      console.error('[GCal] OAuth error:', resp);
      showToast('구글 로그인 실패: ' + resp.error, 'error');
      return;
    }
    if (resp.access_token) {
      saveGCalToken(resp.access_token, resp.expires_in);
      gapi.client.setToken({ access_token: resp.access_token });
      console.log('[GCal] OAuth success & token saved to LocalStorage');
      updateGCalUI();
      showToast('구글 캘린더 연동 완료! 이벤트를 불러옵니다...', 'success');
      await loadCalendarEventsForMonth();
    }
  };
  const existingToken = gapi.client.getToken();
  gcalTokenClient.requestAccessToken({ prompt: existingToken ? '' : 'consent' });
}

// OAuth sign-out — revoke token and clear events
function handleGoogleSignOut() {
  const token = gapi.client.getToken() || { access_token: localStorage.getItem(STORAGE_KEY_GCAL_TOKEN) };
  if (token && token.access_token) {
    try {
      google.accounts.oauth2.revoke(token.access_token, () => {
        console.log('[GCal] Token revoked');
      });
    } catch (e) {
      console.error('[GCal] Revoke error:', e);
    }
  }
  clearGCalToken();
  gcalEvents = [];
  updateGCalUI();
  renderCalendar();
  showToast('구글 캘린더 연동이 해제됐습니다.', 'info');
}

// Fetch Google Calendar events for the currently displayed month (including Public Holidays & Subscribed Calendars)
async function loadCalendarEventsForMonth() {
  if (!isGCalSignedIn()) {
    console.warn('[GCal] Not signed in, skipping event load');
    return;
  }

  const timeMin = new Date(currentCalYear, currentCalMonth, 1).toISOString();
  const timeMax = new Date(currentCalYear, currentCalMonth + 1, 0, 23, 59, 59).toISOString();

  try {
    console.log(`[GCal] Fetching all user & holiday calendars...`);
    let targetCalendars = ['primary', 'ko.south_korea#holiday@group.v.calendar.google.com', 'ko.south_korea.official#holiday@group.v.calendar.google.com'];

    // Query user's calendar list to get all active calendars (including holidays, shared calendars, etc.)
    try {
      const listResp = await gapi.client.calendar.calendarList.list();
      if (listResp.result && listResp.result.items) {
        const activeIds = listResp.result.items
          .filter(c => c.selected !== false)
          .map(c => c.id);
        if (activeIds.length > 0) {
          targetCalendars = Array.from(new Set([...targetCalendars, ...activeIds]));
        }
      }
    } catch (listErr) {
      console.warn('[GCal] Could not list secondary calendars, using fallback list:', listErr);
    }

    console.log('[GCal] Target calendars to load:', targetCalendars);

    let allFetchedEvents = [];

    // Fetch events for each calendar
    for (const calId of targetCalendars) {
      try {
        const resp = await gapi.client.calendar.events.list({
          calendarId: calId,
          timeMin,
          timeMax,
          showDeleted: false,
          singleEvents: true,
          maxResults: 250,
          orderBy: 'startTime',
        });

        const items = resp.result.items || [];
        const isHolidayCal = calId.includes('holiday');

        items.forEach(item => {
          if (isHolidayCal) {
            item.isHoliday = true;
          }
          allFetchedEvents.push(item);
        });
      } catch (err) {
        // Skip optional secondary calendars if unauthorized or missing
        if (calId === 'primary') {
          console.error('[GCal] Primary calendar fetch failed:', err);
        }
      }
    }

    // De-duplicate events by ID
    const eventMap = new Map();
    allFetchedEvents.forEach(e => {
      if (e && e.id) eventMap.set(e.id, e);
    });

    gcalEvents = Array.from(eventMap.values());
    console.log(`[GCal] Loaded total ${gcalEvents.length} events from ${targetCalendars.length} calendars`);
    renderCalendar();
  } catch (err) {
    console.error('[GCal] Event load error:', err);
    const msg = err.result?.error?.message || err.message || '알 수 없는 오류';
    showToast(`구글 캘린더 로드 실패: ${msg}`, 'error');
  }
}

// Create a new Google Calendar event directly
async function createGoogleCalendarEvent(eventData) {
  if (!isGCalSignedIn()) return null;
  try {
    const resp = await gapi.client.calendar.events.insert({
      calendarId: 'primary',
      resource: eventData,
    });
    console.log('[GCal] Event created:', resp.result);
    return resp.result;
  } catch (err) {
    console.error('[GCal] Event creation failed:', err);
    throw err;
  }
}

// Update an existing Google Calendar event
async function updateGoogleCalendarEvent(eventId, eventData) {
  if (!isGCalSignedIn() || !eventId) return null;
  try {
    const resp = await gapi.client.calendar.events.patch({
      calendarId: 'primary',
      eventId: eventId,
      resource: eventData,
    });
    console.log('[GCal] Event updated:', resp.result);
    return resp.result;
  } catch (err) {
    console.error('[GCal] Event update failed:', err);
    throw err;
  }
}

// Delete a Google Calendar event by ID
async function deleteGoogleCalendarEvent(eventId) {
  if (!isGCalSignedIn() || !eventId) return;
  try {
    await gapi.client.calendar.events.delete({
      calendarId: 'primary',
      eventId: eventId,
    });
    console.log('[GCal] Event deleted:', eventId);
  } catch (err) {
    if (err.status !== 410) { // 410 Gone = already deleted
      console.error('[GCal] Event deletion failed:', err);
      throw err;
    }
  }
}

/* ==========================================================================
   Event Listeners Setup
   ========================================================================== */

function setupEventListeners() {
  // Navigation Filter Tabs (only non-calendar tabs - calendar has its own onclick)
  document.querySelectorAll('.category-nav .nav-tab').forEach(tab => {
    tab.addEventListener('click', (e) => {
      const target = e.currentTarget.getAttribute('data-target');
      if (target === 'calendar') return; // handled by onclick on the button
      showDashboardView(target);
    });
  });

  // Action Buttons in Header
  document.getElementById('btn-sync').addEventListener('click', () => {
    if (!googleScriptUrl) {
      openSettingsModal();
      showToast('구글 앱스 스크립트 URL을 먼저 등록해주세요.', 'warning');
      return;
    }
    fetchLatestFromGoogleSheet(true);
  });

  document.getElementById('btn-archive').addEventListener('click', openArchiveModal);
  document.getElementById('btn-settings').addEventListener('click', openSettingsModal);
  document.getElementById('fab-add').addEventListener('click', () => {
    if (isCalendarView) {
      if (isGCalSignedIn()) {
        openGCalModal();
      } else {
        showToast('구글 캘린더를 연동하면 일정을 추가할 수 있습니다.', 'warning');
        handleGoogleSignIn();
      }
    } else {
      openAddModal('today');
    }
  });

  // Modals Close Buttons
  document.getElementById('btn-close-item-modal').addEventListener('click', closeAddModal);
  document.getElementById('btn-cancel-item').addEventListener('click', closeAddModal);

  document.getElementById('btn-close-archive-modal').addEventListener('click', closeArchiveModal);
  document.getElementById('btn-close-archive-bottom').addEventListener('click', closeArchiveModal);

  document.getElementById('btn-close-settings-modal').addEventListener('click', closeSettingsModal);
  document.getElementById('btn-cancel-settings').addEventListener('click', closeSettingsModal);

  // Calendar Controls Event Listeners
  const prevBtn = document.getElementById('cal-btn-prev');
  const nextBtn = document.getElementById('cal-btn-next');
  const todayBtn = document.getElementById('cal-btn-today');
  if (prevBtn) prevBtn.addEventListener('click', () => changeCalMonth(-1));
  if (nextBtn) nextBtn.addEventListener('click', () => changeCalMonth(1));
  if (todayBtn) todayBtn.addEventListener('click', goCalToday);

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
  if (isCalendarView) {
    renderCalendar(); // Just refresh calendar data, don't touch display
  } else {
    filterDashboardView(currentFilter);
  }
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
  isCalendarView = false;

  const wrapper = document.getElementById('dashboard-wrapper');
  const calContainer = document.getElementById('calendar-container');
  const mainCardsContainer = document.getElementById('main-cards-container');
  const categorySidebar = document.getElementById('category-sidebar');
  const cards = document.querySelectorAll('.category-card');

  // Hide calendar, show dashboard wrapper
  if (calContainer) calContainer.style.display = 'none';
  if (wrapper) wrapper.style.display = '';

  if (target === 'all') {
    wrapper.className = 'dashboard-wrapper mode-all';
    if (categorySidebar) categorySidebar.style.display = 'none';
    if (mainCardsContainer) mainCardsContainer.style.display = 'grid';
    cards.forEach(card => {
      card.classList.remove('focused');
      card.style.display = 'flex';
    });
  } else {
    wrapper.className = 'dashboard-wrapper mode-focused';
    if (categorySidebar) categorySidebar.style.display = 'flex';
    if (mainCardsContainer) mainCardsContainer.style.display = 'flex';
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

/* ==========================================================================
   Redesigned Navigation Dropdown & View Handlers
   ========================================================================== */

function toggleNavDropdown(e) {
  if (e) e.stopPropagation();
  const wrapper = document.getElementById('nav-dropdown-wrapper');
  if (wrapper) wrapper.classList.toggle('open');
}

function selectNavFilter(target, label, iconClass) {
  const wrapper = document.getElementById('nav-dropdown-wrapper');
  if (wrapper) wrapper.classList.remove('open');

  // Update dropdown button label & icon
  const labelEl = document.getElementById('nav-dropdown-label');
  const iconEl = document.getElementById('nav-dropdown-icon');
  const btnEl = document.getElementById('nav-dropdown-btn');
  const calBtn = document.getElementById('nav-calendar-btn');

  if (labelEl) labelEl.textContent = label;
  if (iconEl) iconEl.className = iconClass;
  if (btnEl) btnEl.classList.add('active');
  if (calBtn) calBtn.classList.remove('active');

  // Update active state in dropdown menu items
  document.querySelectorAll('.nav-dropdown-item').forEach(item => {
    item.classList.toggle('active', item.getAttribute('data-target') === target);
  });

  filterDashboardView(target);
}

function selectNavCalendar(btnEl) {
  const wrapper = document.getElementById('nav-dropdown-wrapper');
  if (wrapper) wrapper.classList.remove('open');

  showCalendarView(btnEl);
}

// Close navigation dropdown when clicking anywhere outside
document.addEventListener('click', (e) => {
  const wrapper = document.getElementById('nav-dropdown-wrapper');
  if (wrapper && !wrapper.contains(e.target)) {
    wrapper.classList.remove('open');
  }
});

// Dedicated Calendar Show/Hide
function showCalendarView(btnEl) {
  isCalendarView = true;
  currentFilter = 'calendar';

  const dropdownBtn = document.getElementById('nav-dropdown-btn');
  const calBtn = document.getElementById('nav-calendar-btn');
  if (dropdownBtn) dropdownBtn.classList.remove('active');
  if (calBtn) calBtn.classList.add('active');

  // Show calendar, hide dashboard
  const cal = document.getElementById('calendar-container');
  const dash = document.getElementById('dashboard-wrapper');
  if (cal) { cal.style.display = 'flex'; cal.style.flexDirection = 'column'; }
  if (dash) dash.style.display = 'none';

  renderCalendar();
}

/* ==========================================================================
   Calendar View Rendering & Controls
   ========================================================================== */

// Helper: Check if an event falls on a specific date (supports multi-day date ranges!)
function isEventOnDate(event, cellDateStr) {
  if (!event || !event.start) return false;

  const startStr = (event.start.date || event.start.dateTime || '').substring(0, 10);
  let endStr = (event.end?.date || event.end?.dateTime || '').substring(0, 10);

  if (!startStr) return false;
  if (!endStr) endStr = startStr;

  // Google Calendar API all-day events use exclusive end.date
  // e.g. Aug 29 ~ Aug 31 has start='2026-08-29', end='2026-09-01'
  if (event.start.date && event.end?.date && endStr > startStr) {
    return cellDateStr >= startStr && cellDateStr < endStr;
  }

  // Timed event or same-day event
  return cellDateStr >= startStr && cellDateStr <= endStr;
}

function openDayDetailForCell(dateStr) {
  if (!isGCalSignedIn()) {
    showToast('구글 캘린더를 연동하면 일정을 추가/관리할 수 있습니다.', 'warning');
    handleGoogleSignIn();
    return;
  }
  const dayEvents = gcalEvents.filter(e => isEventOnDate(e, dateStr));
  if (dayEvents.length > 0) {
    openDayDetailModal(dateStr, dayEvents);
  } else {
    openGCalModal(null, dateStr);
  }
}

function renderCalendar() {
  const titleEl = document.getElementById('cal-month-year-title');
  const daysBodyEl = document.getElementById('cal-days-body');
  if (!titleEl || !daysBodyEl) return;

  // Set Month Title
  titleEl.textContent = `${currentCalYear}년 ${currentCalMonth + 1}월`;

  daysBodyEl.innerHTML = '';

  const firstDayIndex = new Date(currentCalYear, currentCalMonth, 1).getDay(); // 0 = Sun
  const lastDate = new Date(currentCalYear, currentCalMonth + 1, 0).getDate();
  const prevLastDate = new Date(currentCalYear, currentCalMonth, 0).getDate();

  const todayStr = getTodayDateString();

  // 1. Previous Month Days Padding
  for (let x = firstDayIndex; x > 0; x--) {
    const prevDate = prevLastDate - x + 1;
    const cell = document.createElement('div');
    cell.className = 'cal-day-cell other-month';
    cell.innerHTML = `<span class="cal-day-num">${prevDate}</span>`;
    daysBodyEl.appendChild(cell);
  }

  // 2. Current Month Days
  for (let i = 1; i <= lastDate; i++) {
    const cellDateStr = `${currentCalYear}-${String(currentCalMonth + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
    const isToday = (cellDateStr === todayStr);

    const cell = document.createElement('div');
    cell.className = `cal-day-cell ${isToday ? 'is-today' : ''}`;
    cell.onclick = () => openDayDetailForCell(cellDateStr);

    let cellHtml = `<span class="cal-day-num">${i}</span>`;
    cellHtml += `<div class="cal-events-list">`;

    // Render Google Calendar & Holiday events using multi-day range checker
    gcalEvents.forEach(event => {
      if (!isEventOnDate(event, cellDateStr)) return;

      const title = escapeHtml(event.summary || '(제목 없음)');
      const isHoliday = event.isHoliday || (event.organizer?.email || '').includes('holiday') || (event.id || '').includes('holiday');

      if (isHoliday) {
        cellHtml += `
          <div class="cal-task-pill gcal-holiday" title="공휴일: ${title}" onclick="event.stopPropagation(); openDayDetailForCell('${cellDateStr}')">
            <i class="ri-flag-fill" style="font-size:0.75em; flex-shrink:0; color:#dc2626;"></i>
            <span>${title}</span>
          </div>
        `;
      } else {
        const colorStyle = getGCalEventColor(event.colorId);
        cellHtml += `
          <div class="cal-task-pill gcal-event" title="구글 캘린더: ${title}" style="border-left-color: ${colorStyle.border}; background: ${colorStyle.bg}; color: ${colorStyle.text};" onclick="event.stopPropagation(); openDayDetailForCell('${cellDateStr}')">
            <i class="ri-google-fill" style="font-size:0.75em; flex-shrink:0; color:${colorStyle.border};"></i>
            <span>${title}</span>
          </div>
        `;
      }
    });

    cellHtml += `</div>`;
    cell.innerHTML = cellHtml;
    daysBodyEl.appendChild(cell);
  }

  // 3. Next Month Days Padding to complete grid rows
  const totalCells = firstDayIndex + lastDate;
  const remainingCells = (7 - (totalCells % 7)) % 7;
  for (let j = 1; j <= remainingCells; j++) {
    const cell = document.createElement('div');
    cell.className = 'cal-day-cell other-month';
    cell.innerHTML = `<span class="cal-day-num">${j}</span>`;
    daysBodyEl.appendChild(cell);
  }
}

function changeCalMonth(offset) {
  currentCalMonth += offset;
  if (currentCalMonth > 11) {
    currentCalMonth = 0;
    currentCalYear += 1;
  } else if (currentCalMonth < 0) {
    currentCalMonth = 11;
    currentCalYear -= 1;
  }
  if (isGCalSignedIn()) {
    gcalEvents = [];
    loadCalendarEventsForMonth(); // async, will call renderCalendar when done
  } else {
    renderCalendar();
  }
}

function goCalToday() {
  const now = new Date();
  currentCalYear = now.getFullYear();
  currentCalMonth = now.getMonth();
  if (isGCalSignedIn()) {
    gcalEvents = [];
    loadCalendarEventsForMonth();
  } else {
    renderCalendar();
  }
}

function switchCategoryView(target) {
  const meta = CATEGORY_META[target];
  const label = meta ? meta.name : '전체보기';
  const icon = meta ? meta.icon : 'ri-dashboard-line';
  selectNavFilter(target, label, icon);
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
    ? `<span class="due-date"><i class="ri-calendar-line"></i> ${formatKoreanDateOnly(item.dueDate)}</span>` 
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
            [${catName}] ${item.status === 'completed' ? '완료됨' : '삭제됨'} • ${formatKoreanDateOnly(item.updatedAt || item.createdAt)}
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

  // Push updated state to Google Sheets master DB
  if (googleScriptUrl) {
    pushLocalToGoogleSheet();
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
      pushLocalToGoogleSheet();
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
      pushLocalToGoogleSheet();
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
    pushLocalToGoogleSheet();
  }
}

function permanentlyDeleteItem(id) {
  appData.items = appData.items.filter(i => i.id !== id);

  saveData();
  renderAll();
  showToast('항목이 영구적으로 삭제되었습니다.', 'warning');

  if (googleScriptUrl) {
    pushLocalToGoogleSheet();
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
      pushLocalToGoogleSheet();
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

function openAddModal(category = 'today', prefillDate = null) {
  document.getElementById('modal-item-title').textContent = '새 항목 추가';
  document.getElementById('item-id').value = '';
  document.getElementById('item-category').value = category;
  document.getElementById('item-content').value = '';
  document.getElementById('item-priority').value = 'medium';
  
  if (prefillDate) {
    document.getElementById('item-duedate').value = prefillDate;
  } else if (category === 'today') {
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
  const clientIdInput = document.getElementById('google-client-id');
  if (clientIdInput) clientIdInput.value = googleClientId;
  document.getElementById('modal-settings').classList.add('active');
}

function closeSettingsModal() {
  document.getElementById('modal-settings').classList.remove('active');
}

function saveSettings() {
  const url = document.getElementById('apps-script-url').value.trim();
  googleScriptUrl = url;
  localStorage.setItem(STORAGE_KEY_URL, url);

  const clientIdInput = document.getElementById('google-client-id');
  if (clientIdInput) {
    googleClientId = clientIdInput.value.trim();
    localStorage.setItem(STORAGE_KEY_CLIENT_ID, googleClientId);
  }

  showToast('설정이 저장되었습니다.', 'success');
  closeSettingsModal();

  if (googleScriptUrl) {
    fetchLatestFromGoogleSheet(true);
  } else {
    updateSyncBadge('local', '로컬 저장소 모드');
  }
}

/* ==========================================================================
   Google Apps Script Bidirectional Sync Engine
   ========================================================================== */

// 1. Pull Latest Data from Google Sheet Master DB
async function fetchLatestFromGoogleSheet(manual = false) {
  if (!googleScriptUrl) return false;

  updateSyncBadge('syncing', '클라우드 수신 중...');

  try {
    const response = await fetch(`${googleScriptUrl}?action=get&t=${Date.now()}`);
    const result = await response.json();

    if (result && result.status === 'success' && Array.isArray(result.items)) {
      if (result.items.length > 0) {
        // Master Sheet has data: Merge cloud items with local storage
        appData.items = mergeItems(result.items, appData.items);
        saveData();
        renderAll();
      } else if (appData.items.length > 0) {
        // Master Sheet is empty: push initial local items to sheet
        await pushLocalToGoogleSheet();
      }
      updateSyncBadge('success', '동기화 완료');
      if (manual) showToast('구글 스프레드시트 최신 데이터를 불러왔습니다!', 'success');
      return true;
    }
  } catch (err) {
    console.error('Fetch from Google Sheet failed:', err);
    updateSyncBadge('error', '동기화 실패');
    if (manual) showToast('스프레드시트 수신 실패. 로컬 데이터를 유지합니다.', 'error');
  }
  return false;
}

// 2. Push Local Items State to Google Sheet Master DB
async function pushLocalToGoogleSheet(manual = false) {
  if (!googleScriptUrl) return false;

  updateSyncBadge('syncing', '스프레드시트에 저장 중...');

  try {
    const encodedData = encodeURIComponent(JSON.stringify(appData.items));
    const syncUrl = `${googleScriptUrl}?action=sync&data=${encodedData}&t=${Date.now()}`;

    const response = await fetch(syncUrl);
    const result = await response.json();

    if (result && result.status === 'success') {
      if (result.items && Array.isArray(result.items) && result.items.length > 0) {
        appData.items = result.items;
        saveData();
        renderAll();
      }
      updateSyncBadge('success', '동기화 완료');
      if (manual) showToast('구글 스프레드시트에 저장되었습니다!', 'success');
      return true;
    }
  } catch (err) {
    console.warn('GET push failed, trying POST fallback:', err);
  }

  // Fallback POST push
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
      if (result.items && Array.isArray(result.items) && result.items.length > 0) {
        appData.items = result.items;
        saveData();
        renderAll();
      }
      updateSyncBadge('success', '동기화 완료');
      if (manual) showToast('구글 스프레드시트에 저장되었습니다!', 'success');
      return true;
    }
  } catch (fetchErr) {
    console.error('Cloud Push failed:', fetchErr);
    updateSyncBadge('error', '동기화 실패');
    if (manual) showToast('스프레드시트 저장 실패. 로컬 데이터로 유지됩니다.', 'error');
  }

  return false;
}

// 3. Intelligent Item Merger by ID & Timestamp
function mergeItems(cloudItems, localItems) {
  const itemMap = new Map();

  // Cloud items are SSOT (Single Source of Truth) from Google Sheets DB
  cloudItems.forEach(item => {
    if (item && item.id) {
      itemMap.set(item.id, item);
    }
  });

  // Merge local items if they are newer or not present in cloud
  localItems.forEach(item => {
    if (!item || !item.id) return;

    if (!itemMap.has(item.id)) {
      itemMap.set(item.id, item);
    } else {
      const existing = itemMap.get(item.id);
      const cloudTime = parseTimestampToMillis(existing.updatedAt || existing.createdAt);
      const localTime = parseTimestampToMillis(item.updatedAt || item.createdAt);

      if (localTime > cloudTime) {
        itemMap.set(item.id, item);
      }
    }
  });

  return Array.from(itemMap.values());
}

function parseTimestampToMillis(dateStr) {
  if (!dateStr) return 0;
  const str = String(dateStr).trim();
  
  // Standard JS Date parse
  let time = new Date(str).getTime();
  if (!isNaN(time)) return time;

  // Mobile Safari fallback: strip timezone parentheses
  const cleanStr = str.replace(/\(.*\)/, '').replace(/GMT.*$/, '').trim();
  time = new Date(cleanStr).getTime();
  if (!isNaN(time)) return time;

  // Regex parse for YYYY-MM-DD
  const match = str.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  if (match) {
    return new Date(match[1], match[2] - 1, match[3]).getTime();
  }

  return 0;
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

function formatKoreanDateOnly(dateStr) {
  if (!dateStr) return '';
  const str = String(dateStr).trim();
  
  // Format YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    const parts = str.split('-');
    return `${parts[0]}.${parts[1]}.${parts[2]}`;
  }

  try {
    const cleanStr = str.replace(/\(.*\)/, '').replace(/GMT.*$/, '').trim();
    const d = new Date(cleanStr);
    if (isNaN(d.getTime())) return str;
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}.${m}.${day}`;
  } catch (e) {
    return str;
  }
}

function escapeHtml(unsafe) {
  if (!unsafe) return '';
  return String(unsafe)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/* ==========================================================================
   Google Calendar Direct Modal & Event Management
   ========================================================================== */

function openGCalModal(eventObj = null, defaultDateStr = null) {
  const modal = document.getElementById('modal-gcal-event');
  const titleEl = document.getElementById('modal-gcal-title');
  const idInput = document.getElementById('gcal-event-id');
  const summaryInput = document.getElementById('gcal-event-summary');
  const startDateInput = document.getElementById('gcal-event-startdate');
  const endDateInput = document.getElementById('gcal-event-enddate');
  const descInput = document.getElementById('gcal-event-description');
  const deleteBtn = document.getElementById('btn-delete-gcal-event');

  if (!modal) return;

  if (eventObj) {
    // Edit existing Google Calendar event
    titleEl.textContent = '구글 캘린더 일정 수정';
    idInput.value = eventObj.id || '';
    summaryInput.value = eventObj.summary || '';
    
    const startDate = eventObj.start?.date || (eventObj.start?.dateTime || '').substring(0, 10);
    const endDate = eventObj.end?.date || (eventObj.end?.dateTime || '').substring(0, 10);
    
    startDateInput.value = startDate;
    endDateInput.value = endDate;
    descInput.value = eventObj.description || '';

    const colorVal = String(eventObj.colorId || '9');
    const colorRadio = document.querySelector(`input[name="gcal-color"][value="${colorVal}"]`);
    if (colorRadio) colorRadio.checked = true;
    
    deleteBtn.style.display = 'flex';
  } else {
    // Create new Google Calendar event
    titleEl.textContent = '구글 캘린더 일정 추가';
    idInput.value = '';
    summaryInput.value = '';
    startDateInput.value = defaultDateStr || getTodayDateString();
    endDateInput.value = '';
    descInput.value = '';

    const defaultRadio = document.querySelector('input[name="gcal-color"][value="9"]');
    if (defaultRadio) defaultRadio.checked = true;
    
    deleteBtn.style.display = 'none';
  }

  modal.classList.add('active');
}

function closeGCalModal() {
  const modal = document.getElementById('modal-gcal-event');
  if (modal) modal.classList.remove('active');
}

function handleGCalPillClick(eventId) {
  const eventObj = gcalEvents.find(e => e.id === eventId);
  if (eventObj) {
    openGCalModal(eventObj);
  } else {
    showToast('선택한 구글 캘린더 이벤트 정보를 찾을 수 없습니다.', 'warning');
  }
}

async function handleGCalFormSubmit(e) {
  e.preventDefault();

  if (!isGCalSignedIn()) {
    showToast('구글 캘린더가 연동되어 있지 않습니다.', 'error');
    return;
  }

  const eventId = document.getElementById('gcal-event-id').value;
  const summary = document.getElementById('gcal-event-summary').value.trim();
  const startDate = document.getElementById('gcal-event-startdate').value;
  const endDate = document.getElementById('gcal-event-enddate').value || startDate;
  const description = document.getElementById('gcal-event-description').value.trim();
  
  const selectedColorEl = document.querySelector('input[name="gcal-color"]:checked');
  const colorId = selectedColorEl ? selectedColorEl.value : '9';

  if (!summary || !startDate) return;

  const eventResource = {
    summary,
    description,
    start: { date: startDate },
    end: { date: endDate },
    colorId: colorId,
  };

  try {
    if (eventId) {
      await updateGoogleCalendarEvent(eventId, eventResource);
      showToast('구글 캘린더 일정이 수정되었습니다.', 'success');
    } else {
      await createGoogleCalendarEvent(eventResource);
      showToast('구글 캘린더에 새 일정이 추가되었습니다.', 'success');
    }
    closeGCalModal();
    await loadCalendarEventsForMonth();
  } catch (err) {
    showToast('구글 캘린더 일정 저장 실패: ' + (err.result?.error?.message || err.message || '오류 발생'), 'error');
  }
}

async function handleGCalDeleteClick() {
  const eventId = document.getElementById('gcal-event-id').value;
  if (!eventId) return;

  if (!confirm('이 구글 캘린더 일정을 정말 삭제하시겠습니까?')) return;

  try {
    await deleteGoogleCalendarEvent(eventId);
    showToast('구글 캘린더 일정이 삭제되었습니다.', 'info');
    closeGCalModal();
    await loadCalendarEventsForMonth();
  } catch (err) {
    showToast('구글 캘린더 일정 삭제 실패: ' + (err.result?.error?.message || err.message || '오류 발생'), 'error');
  }
}

/* ==========================================================================
   Google Calendar Day Detail Agenda Modal
   ========================================================================== */

let currentSelectedDayStr = null;

function openDayDetailModal(dateStr, events = []) {
  currentSelectedDayStr = dateStr;
  const modal = document.getElementById('modal-gcal-day-detail');
  const titleEl = document.getElementById('modal-day-detail-title');
  const subtitleEl = document.getElementById('modal-day-detail-subtitle');
  const listEl = document.getElementById('day-detail-events-list');

  if (!modal || !titleEl || !listEl) return;

  const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
  const d = new Date(dateStr + 'T00:00:00');
  const dayName = isNaN(d.getTime()) ? '' : ` (${dayNames[d.getDay()]})`;

  const parts = dateStr.split('-');
  titleEl.textContent = `${parts[0]}년 ${parseInt(parts[1], 10)}월 ${parseInt(parts[2], 10)}일${dayName} 일정`;
  if (subtitleEl) subtitleEl.textContent = `등록된 일정 ${events.length}개`;

  listEl.innerHTML = '';

  events.forEach(event => {
    const card = document.createElement('div');
    card.className = 'day-detail-event-card';
    
    const colorStyle = getGCalEventColor(event.colorId);
    card.style.borderLeftColor = colorStyle.border;
    
    const summary = escapeHtml(event.summary || '(제목 없음)');
    const desc = escapeHtml(event.description || '');
    const descHtml = desc ? `<div class="day-detail-event-desc">${desc}</div>` : '';
    const eventIdEscaped = escapeHtml(event.id);

    let timeInfo = '';
    if (event.start?.dateTime && event.end?.dateTime) {
      const startTime = event.start.dateTime.substring(11, 16);
      const endTime = event.end.dateTime.substring(11, 16);
      timeInfo = `<div class="day-detail-event-time"><i class="ri-time-line"></i> ${startTime} ~ ${endTime}</div>`;
    } else {
      timeInfo = `<div class="day-detail-event-time"><i class="ri-calendar-line"></i> 하루 종일</div>`;
    }

    card.innerHTML = `
      <div class="day-detail-event-info">
        <div class="day-detail-event-title">${summary}</div>
        ${timeInfo}
        ${descHtml}
      </div>
      <div class="day-detail-event-actions">
        <button class="btn-event-action edit" onclick="handleDayDetailEdit('${eventIdEscaped}')" title="수정">
          <i class="ri-edit-line"></i> 수정
        </button>
        <button class="btn-event-action delete" onclick="handleDayDetailDelete('${eventIdEscaped}')" title="삭제">
          <i class="ri-delete-bin-line"></i> 삭제
        </button>
      </div>
    `;
    listEl.appendChild(card);
  });

  modal.classList.add('active');
}

function closeDayDetailModal() {
  const modal = document.getElementById('modal-gcal-day-detail');
  if (modal) modal.classList.remove('active');
}

function handleAddEventOnDayClick() {
  closeDayDetailModal();
  openGCalModal(null, currentSelectedDayStr || getTodayDateString());
}

function handleDayDetailEdit(eventId) {
  closeDayDetailModal();
  const eventObj = gcalEvents.find(e => e.id === eventId);
  if (eventObj) {
    openGCalModal(eventObj);
  }
}

async function handleDayDetailDelete(eventId) {
  const eventObj = gcalEvents.find(e => e.id === eventId);
  const title = eventObj?.summary || '일정';
  if (!confirm(`'${title}' 일정을 구글 캘린더에서 정말 삭제하시겠습니까?`)) return;

  try {
    await deleteGoogleCalendarEvent(eventId);
    showToast('구글 캘린더 일정이 삭제되었습니다.', 'info');
    await loadCalendarEventsForMonth();
    
    // Refresh day detail modal or close if empty
    const remainingEvents = gcalEvents.filter(e => {
      const d = e.start?.date || (e.start?.dateTime || '').substring(0, 10);
      return d === currentSelectedDayStr;
    });

    if (remainingEvents.length > 0) {
      openDayDetailModal(currentSelectedDayStr, remainingEvents);
    } else {
      closeDayDetailModal();
    }
  } catch (err) {
    showToast('구글 캘린더 일정 삭제 실패: ' + (err.result?.error?.message || err.message || '오류 발생'), 'error');
  }
}

/* ==========================================================================
   Dr.NAEO PDF Report Generator
   ========================================================================== */

function openPdfReportModal() {
  const modal = document.getElementById('modal-pdf-report');
  if (!modal) return;

  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const cycleInput = document.getElementById('report-cycle-name');
  if (cycleInput && !cycleInput.value) {
    cycleInput.value = `${year}년 ${month}월 4주차 마케팅 리포트`;
  }

  loadAppDataToReportTable();
  modal.classList.add('active');
}

function closePdfReportModal() {
  const modal = document.getElementById('modal-pdf-report');
  if (modal) modal.classList.remove('active');
}

function loadAppDataToReportTable() {
  const tbody = document.getElementById('report-table-body');
  if (!tbody) return;

  tbody.innerHTML = '';

  const items = appData.items.filter(i => i && i.status !== 'deleted');

  if (items.length === 0) {
    addReportTableRow({ keyword: '강남성형외과', type: '블로그포스팅', date: getTodayDateString(), url: 'https://blog.naver.com/sample', result: '스마트블럭 1위' });
    addReportTableRow({ keyword: '지방흡입후기', type: '지식iN답변', date: getTodayDateString(), url: 'https://kin.naver.com/sample', result: '상위 노출' });
  } else {
    items.forEach(item => {
      const categoryName = CATEGORY_NAMES[item.category] || item.category || '기타';
      const date = item.dueDate || (item.createdAt ? item.createdAt.substring(0, 10) : getTodayDateString());
      const resultText = item.note || (item.status === 'completed' ? '발행 완료 (상위 노출)' : '발행 및 검수 완료');
      
      addReportTableRow({
        keyword: item.content || '주요 마케팅 키워드',
        type: categoryName,
        date: date,
        url: item.link || '',
        result: resultText
      });
    });
  }

  updateReportItemCount();
}

function addReportTableRow(data = {}) {
  const tbody = document.getElementById('report-table-body');
  if (!tbody) return;

  const tr = document.createElement('tr');
  tr.style.borderBottom = '1px solid #e2e8f0';
  
  tr.innerHTML = `
    <td style="padding: 6px 8px;"><input type="text" class="form-control report-kw" value="${escapeHtml(data.keyword || '')}" placeholder="키워드 입력" style="font-size:0.82rem; padding:6px 8px;" /></td>
    <td style="padding: 6px 8px;"><input type="text" class="form-control report-type" value="${escapeHtml(data.type || '블로그포스팅')}" placeholder="종류 (예: 블로그)" style="font-size:0.82rem; padding:6px 8px;" /></td>
    <td style="padding: 6px 8px;"><input type="date" class="form-control report-date" value="${data.date || getTodayDateString()}" style="font-size:0.82rem; padding:6px 8px;" /></td>
    <td style="padding: 6px 8px;"><input type="url" class="form-control report-url" value="${escapeHtml(data.url || '')}" placeholder="https://..." style="font-size:0.82rem; padding:6px 8px;" /></td>
    <td style="padding: 6px 8px;"><input type="text" class="form-control report-result" value="${escapeHtml(data.result || '상위 노출')}" placeholder="성과 / 순위" style="font-size:0.82rem; padding:6px 8px;" /></td>
    <td style="padding: 6px 4px; text-align: center;">
      <button type="button" onclick="this.closest('tr').remove(); updateReportItemCount();" style="background:none; border:none; color:#dc2626; cursor:pointer;" title="삭제"><i class="ri-delete-bin-line"></i></button>
    </td>
  `;

  tbody.appendChild(tr);
  updateReportItemCount();
}

function updateReportItemCount() {
  const rows = document.querySelectorAll('#report-table-body tr');
  const countEl = document.getElementById('report-item-count');
  if (countEl) countEl.textContent = rows.length.toString();
}

function getReportData() {
  const hospitalName = document.getElementById('report-hospital-name')?.value.trim() || '닥터내오 의원';
  const cycleName = document.getElementById('report-cycle-name')?.value.trim() || '2026년 8월 마케팅 리포트';
  const targetQuota = parseInt(document.getElementById('report-target-quota')?.value || '0', 10);
  const summaryText = document.getElementById('report-summary-text')?.value.trim() || '주요 키워드별 상위 노출 달성 및 목표 발행량 정상 완료.';

  const rows = document.querySelectorAll('#report-table-body tr');
  const tableData = [];

  rows.forEach((tr, index) => {
    const kw = tr.querySelector('.report-kw')?.value.trim() || '';
    const type = tr.querySelector('.report-type')?.value.trim() || '';
    const date = tr.querySelector('.report-date')?.value || '';
    const url = tr.querySelector('.report-url')?.value.trim() || '';
    const result = tr.querySelector('.report-result')?.value.trim() || '';

    if (kw || url) {
      tableData.push({ idx: index + 1, kw, type, date, url, result });
    }
  });

  const publishedCount = tableData.length;
  const achievementRate = targetQuota > 0 ? Math.round((publishedCount / targetQuota) * 100) : 100;

  const keywordSummaryList = tableData.map(item => {
    return `<span style="display:inline-block; background:#e8f5e9; color:#0e8a57; font-weight:700; padding:4px 10px; border-radius:12px; font-size:12px; border:1px solid #a3e635; margin:3px;">🏷️ ${escapeHtml(item.kw)} (${escapeHtml(item.result || '노출완료')})</span>`;
  }).join(' ');

  return {
    hospitalName,
    cycleName,
    targetQuota,
    publishedCount,
    achievementRate,
    summaryText,
    keywordSummaryList,
    tableData,
    createdDate: new Date().toLocaleDateString('ko-KR')
  };
}

function buildReportHtml(data) {
  const tableRowsHtml = data.tableData.map(row => `
    <tr style="border-bottom: 1px solid #e2e8f0;">
      <td style="padding: 10px 12px; text-align: center; font-weight: 700; color: #64748b; font-size: 12px;">${row.idx}</td>
      <td style="padding: 10px 12px; font-weight: 700; color: #0f172a; font-size: 13px;">${escapeHtml(row.kw)}</td>
      <td style="padding: 10px 12px; font-size: 12px; color: #334155;"><span style="background:#f1f5f9; padding:3px 8px; border-radius:6px; font-weight:600;">${escapeHtml(row.type)}</span></td>
      <td style="padding: 10px 12px; font-size: 12px; color: #475569;">${escapeHtml(row.date)}</td>
      <td style="padding: 10px 12px; font-size: 12px; word-break: break-all;">
        ${row.url ? `<a href="${escapeHtml(row.url)}" target="_blank" style="color:#0284c7; text-decoration:none; font-weight:600;">${escapeHtml(row.url.substring(0, 35))}${row.url.length > 35 ? '...' : ''}</a>` : '-'}
      </td>
      <td style="padding: 10px 12px; font-size: 12px;">
        <span style="background:#dcfce7; color:#15803d; font-weight:700; padding:3px 8px; border-radius:6px; border:1px solid #86efac;">${escapeHtml(row.result || '상위 노출')}</span>
      </td>
    </tr>
  `).join('');

  return `
    <div id="drnaeo-pdf-document" style="font-family:'Pretendard Variable',Pretendard,sans-serif; width: 100%; max-width: 800px; margin: 0 auto; background: #ffffff; padding: 40px; color: #1e293b; box-sizing: border-box;">
      
      <!-- Report Header -->
      <div style="display: flex; align-items: center; justify-content: space-between; border-bottom: 3px solid #0E8A57; padding-bottom: 20px; margin-bottom: 28px;">
        <div style="display: flex; align-items: center; gap: 12px;">
          <svg viewBox="0 0 24 24" width="44" height="44" style="flex-shrink:0;">
            <defs>
              <path id="tri-rep" d="M12 2.2c1.5 0 2.7.8 3.5 2.2l4.8 8.3c.8 1.4.9 2.9.1 4.2-.8 1.3-2.1 2-3.7 2H7.3c-1.6 0-2.9-.7-3.7-2-.8-1.3-.7-2.8.1-4.2l4.8-8.3C9.3 3 10.5 2.2 12 2.2z"/>
              <clipPath id="tc-rep"><use href="#tri-rep"/></clipPath>
              <radialGradient id="wg1-rep" cx="30%" cy="24%" r="75%"><stop offset="0%" stop-color="#37D67F"/><stop offset="100%" stop-color="#37D67F" stop-opacity="0"/></radialGradient>
              <radialGradient id="wg2-rep" cx="76%" cy="30%" r="80%"><stop offset="0%" stop-color="#0FB864"/><stop offset="100%" stop-color="#0FB864" stop-opacity="0"/></radialGradient>
              <radialGradient id="wg3-rep" cx="28%" cy="80%" r="85%"><stop offset="0%" stop-color="#0E8A57"/><stop offset="100%" stop-color="#0E8A57" stop-opacity="0"/></radialGradient>
              <radialGradient id="wg4-rep" cx="80%" cy="78%" r="85%"><stop offset="0%" stop-color="#0B6E43"/><stop offset="100%" stop-color="#0B6E43" stop-opacity="0"/></radialGradient>
              <g id="lg-medigo-rep">
                <g clip-path="url(#tc-rep)">
                  <rect width="24" height="24" fill="#12A35C"/>
                  <rect width="24" height="24" fill="url(#wg4-rep)"/>
                  <rect width="24" height="24" fill="url(#wg3-rep)"/>
                  <rect width="24" height="24" fill="url(#wg2-rep)"/>
                  <rect width="24" height="24" fill="url(#wg1-rep)"/>
                  <circle cx="8" cy="9" r="4.5" fill="#37D67F" opacity=".28"/>
                  <circle cx="16.5" cy="15" r="5" fill="#0B6E43" opacity=".3"/>
                  <circle cx="13" cy="7.5" r="3" fill="#5FE49A" opacity=".22"/>
                </g>
              </g>
            </defs>
            <use href="#lg-medigo-rep"/>
          </svg>
          <div>
            <div style="font-size: 22px; font-weight: 800; color: #0b132a; letter-spacing: -0.5px;">Dr.<span style="color:#0E8A57;">NAEO</span></div>
            <div style="font-size: 11px; color: #64748b; font-weight: 600;">의사가 직접 쓰는 네이버 마케팅</div>
          </div>
        </div>
        <div style="text-align: right;">
          <div style="font-size: 20px; font-weight: 800; color: #0E8A57;">마케팅 성과보고서</div>
          <div style="font-size: 12px; color: #64748b; margin-top: 4px; font-weight: 600;">작성일자: ${data.createdDate}</div>
        </div>
      </div>

      <!-- Report Metadata Overview Cards -->
      <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 20px; margin-bottom: 28px;">
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 16px;">
          <div style="background: #ffffff; padding: 14px; border-radius: 10px; border: 1px solid #cbd5e1;">
            <div style="font-size: 11px; color: #64748b; font-weight: 700; margin-bottom: 4px;">🏥 대상 병원</div>
            <div style="font-size: 15px; font-weight: 800; color: #0f172a;">${escapeHtml(data.hospitalName)}</div>
          </div>
          <div style="background: #ffffff; padding: 14px; border-radius: 10px; border: 1px solid #cbd5e1;">
            <div style="font-size: 11px; color: #64748b; font-weight: 700; margin-bottom: 4px;">📅 마케팅 주기</div>
            <div style="font-size: 14px; font-weight: 800; color: #0f172a;">${escapeHtml(data.cycleName)}</div>
          </div>
          <div style="background: #ffffff; padding: 14px; border-radius: 10px; border: 1px solid #cbd5e1;">
            <div style="font-size: 11px; color: #64748b; font-weight: 700; margin-bottom: 4px;">🎯 발행 현황 (할당량)</div>
            <div style="font-size: 15px; font-weight: 800; color: #0E8A57;">${data.publishedCount} / ${data.targetQuota}건</div>
          </div>
          <div style="background: #ffffff; padding: 14px; border-radius: 10px; border: 1px solid #cbd5e1;">
            <div style="font-size: 11px; color: #64748b; font-weight: 700; margin-bottom: 4px;">📊 목표 달성률</div>
            <div style="font-size: 15px; font-weight: 800; color: #0284c7;">${data.achievementRate}% 달성</div>
          </div>
        </div>

        <!-- Summary & Keyword Badges -->
        <div style="background: #ffffff; padding: 14px; border-radius: 10px; border: 1px solid #cbd5e1;">
          <div style="font-size: 12px; font-weight: 800; color: #0f172a; margin-bottom: 6px;">💡 이번 주기 종합 성과 요약</div>
          <div style="font-size: 13px; color: #334155; line-height: 1.6; margin-bottom: 10px;">${escapeHtml(data.summaryText)}</div>
          <div style="font-size: 11px; font-weight: 700; color: #64748b; margin-bottom: 4px;">🏷️ 주요 키워드별 성과 목록</div>
          <div>${data.keywordSummaryList}</div>
        </div>
      </div>

      <!-- Detailed Table -->
      <div style="margin-bottom: 28px;">
        <div style="font-size: 16px; font-weight: 800; color: #0f172a; margin-bottom: 12px; display: flex; align-items: center; gap: 6px;">
          <span>📋 세부 발행 및 성과 내역</span>
          <span style="font-size: 12px; color: #64748b; font-weight: 600;">(총 ${data.tableData.length}건)</span>
        </div>

        <table style="width: 100%; border-collapse: collapse; border: 1px solid #cbd5e1; border-radius: 8px; overflow: hidden;">
          <thead>
            <tr style="background: #0e1015; color: #ffffff; font-size: 12px;">
              <th style="padding: 10px 12px; width: 6%; text-align: center;">NO</th>
              <th style="padding: 10px 12px; width: 26%; text-align: left;">키워드 (Keyword)</th>
              <th style="padding: 10px 12px; width: 16%; text-align: left;">항목종류</th>
              <th style="padding: 10px 12px; width: 14%; text-align: left;">발행일</th>
              <th style="padding: 10px 12px; width: 23%; text-align: left;">URL</th>
              <th style="padding: 10px 12px; width: 15%; text-align: left;">성과 / 순위</th>
            </tr>
          </thead>
          <tbody>
            ${tableRowsHtml}
          </tbody>
        </table>
      </div>

      <!-- Report Footer -->
      <div style="border-top: 1px solid #e2e8f0; padding-top: 16px; display: flex; align-items: center; justify-content: space-between; font-size: 11px; color: #94a3b8;">
        <div><b>Dr.NAEO</b> — 의사가 직접 쓰는 네이버 마케팅</div>
        <div>본 보고서는 Dr.NAEO 자동화 리포트 시스템에 의해 생성되었습니다.</div>
      </div>
    </div>
  `;
}

function previewPdfReport() {
  const data = getReportData();
  const reportHtml = buildReportHtml(data);
  
  const win = window.open('', '_blank');
  win.document.write(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Dr.NAEO 성과보고서 미리보기 - ${escapeHtml(data.hospitalName)}</title>
      <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css">
      <style>
        body { background: #64748b; margin: 0; padding: 40px 10px; font-family: 'Pretendard Variable', sans-serif; }
        @media print {
          body { background: #fff; padding: 0; }
          #btn-print-bar { display: none; }
        }
      </style>
    </head>
    <body>
      <div id="btn-print-bar" style="max-width: 800px; margin: 0 auto 20px; display: flex; justify-content: space-between; align-items: center; background: #ffffff; padding: 12px 20px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
        <div style="font-weight: 700; color: #0f172a;">📄 Dr.NAEO 마케팅 성과보고서 미리보기</div>
        <button onclick="window.print()" style="background: #0E8A57; color: #fff; border: none; padding: 8px 16px; border-radius: 6px; font-weight: 700; cursor: pointer;">🖨️ PDF로 인쇄 / 저장</button>
      </div>
      ${reportHtml}
    </body>
    </html>
  `);
  win.document.close();
}

async function generateAndDownloadPdfReport() {
  const data = getReportData();
  const reportHtml = buildReportHtml(data);

  showToast('Dr.NAEO PDF 보고서를 생성 중입니다...', 'info');

  const element = document.createElement('div');
  element.innerHTML = reportHtml;
  document.body.appendChild(element);

  const filename = `Dr.NAEO_성과보고서_${data.hospitalName.replace(/\s+/g, '_')}_${new Date().toISOString().substring(0, 10)}.pdf`;

  if (typeof html2pdf !== 'undefined') {
    const opt = {
      margin:       [10, 10, 10, 10],
      filename:     filename,
      image:        { type: 'jpeg', quality: 0.98 },
      html2canvas:  { scale: 2, useCORS: true, logging: false },
      jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    try {
      await html2pdf().set(opt).from(element.firstElementChild).save();
      showToast(`'${filename}' PDF 보고서 다운로드 완료!`, 'success');
    } catch (err) {
      console.error('html2pdf error:', err);
      previewPdfReport();
    } finally {
      element.remove();
    }
  } else {
    element.remove();
    previewPdfReport();
  }
}


