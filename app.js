/**
 * Unified Student Calendar — Live Edition
 * Activities: Due Dates, Clubs, Mentorship, Email Dates, Sports
 * Integrations: Canvas, Virtual Virginia, Gmail, Outlook, Teams
 */

const CATEGORIES = {
  duedates: { id: 'duedates', name: 'Due Dates', defaultColor: '#6366f1' },
  clubs: { id: 'clubs', name: 'Clubs', defaultColor: '#8b5cf6' },
  mentorship: { id: 'mentorship', name: 'Mentorship', defaultColor: '#3b82f6' },
  email: { id: 'email', name: 'Email Dates', defaultColor: '#a855f7' },
  sports: { id: 'sports', name: 'Sports', defaultColor: '#6366f1' }
};

const INTEGRATIONS = {
  canvas: {
    id: 'canvas',
    name: 'Canvas',
    sub: 'LMS API / iCal Feed',
    field1Label: 'Canvas Domain (e.g. school.instructure.com)',
    field2Label: 'Canvas API Access Token (optional if .ics set)',
    icsLabel: 'Canvas Calendar Feed (.ics URL)'
  },
  canvas2: {
    id: 'canvas2',
    name: 'Canvas (2)',
    sub: 'Second Canvas LMS / iCal Feed',
    field1Label: 'Canvas Domain (e.g. college.instructure.com)',
    field2Label: 'Canvas API Access Token (optional if .ics set)',
    icsLabel: 'Canvas (2) Calendar Feed (.ics URL)'
  },
  vva: {
    id: 'vva',
    name: 'Virtual Virginia',
    sub: 'VVA LMS API / iCal Feed',
    field1Label: 'VVA Domain (virtualvirginia.instructure.com)',
    field2Label: 'VVA API Access Token (optional if .ics set)',
    icsLabel: 'Virtual Virginia Calendar Feed (.ics URL)'
  },
  gmail: {
    id: 'gmail',
    name: 'Gmail',
    sub: 'Google API / IMAP / iCal',
    field1Label: 'Gmail Address',
    field2Label: 'Google OAuth Token or App Password (IMAP)',
    icsLabel: 'Google Calendar Secret .ics URL (optional)'
  },
  outlook: {
    id: 'outlook',
    name: 'Outlook',
    sub: 'Microsoft Graph / IMAP / iCal',
    field1Label: 'School Outlook Email',
    field2Label: 'MS Graph Bearer Token or App Password',
    icsLabel: 'Outlook Published Calendar .ics URL (optional)'
  },
  teams: {
    id: 'teams',
    name: 'Teams',
    sub: 'MS Graph Meetings / Exchange iCal',
    field1Label: 'Teams Account Email',
    field2Label: 'Microsoft Graph Bearer Token (OnlineMeetings)',
    icsLabel: 'Teams / Exchange Calendar .ics URL (optional)'
  }
};

function offsetDateStr(baseDate, dayOffset) {
  const d = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate() + dayOffset);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

const todayDate = new Date();
const todayStr = offsetDateStr(todayDate, 0);

const state = {
  studentName: '',
  currentView: 'month',
  viewDate: new Date(todayDate.getFullYear(), todayDate.getMonth(), todayDate.getDate()),
  selectedDateStr: todayStr,
  searchQuery: '',
  activeCategories: new Set(Object.keys(CATEGORIES)),
  activeIntegrations: new Set(Object.keys(INTEGRATIONS)),
  categoryColors: {
    duedates: '#6366f1',
    clubs: '#8b5cf6',
    mentorship: '#3b82f6',
    email: '#a855f7',
    sports: '#6366f1'
  },
  integrationsAuth: {
    canvas: { connected: false, domain: '', username: '', token: '', icsUrl: '' },
    canvas2: { connected: false, domain: '', username: '', token: '', icsUrl: '' },
    vva: { connected: false, domain: 'virtualvirginia.instructure.com', username: '', token: '', icsUrl: '' },
    gmail: { connected: false, domain: 'imap.gmail.com', username: '', password: '', token: '', icsUrl: '' },
    outlook: { connected: false, domain: 'outlook.office365.com', username: '', password: '', token: '', icsUrl: '' },
    teams: { connected: false, domain: 'graph.microsoft.com', username: '', password: '', token: '', icsUrl: '' }
  },
  syncStatuses: {},
  events: []
};

function getEventColor(ev) {
  return ev.customColor || state.categoryColors[ev.category] || '#6366f1';
}

function loadPersistedState() {
  try {
    // Clear legacy demo keys that had fake events
    localStorage.removeItem('nexus_unified_calendar_v1');
    localStorage.removeItem('nexus_minimal_calendar_v2');

    // Migrate credentials (if any) from v3 without importing fake ev-1..ev-11 items
    const v3Raw = localStorage.getItem('nexus_minimal_calendar_v3');
    if (v3Raw && !localStorage.getItem('nexus_live_calendar_v4')) {
      const v3 = JSON.parse(v3Raw);
      if (v3.studentName) state.studentName = v3.studentName;
      if (v3.integrationsAuth) {
        Object.keys(INTEGRATIONS).forEach(k => {
          if (v3.integrationsAuth[k]) {
            const c = v3.integrationsAuth[k];
            const hasRealCreds = Boolean((c.token && c.token.trim()) || (c.icsUrl && c.icsUrl.trim()) || (c.password && c.password.trim()));
            state.integrationsAuth[k] = {
              ...state.integrationsAuth[k],
              ...c,
              connected: hasRealCreds ? Boolean(c.connected) : false
            };
          }
        });
      }
      if (Array.isArray(v3.events)) {
        // Keep only non-mock events
        state.events = v3.events.filter(e => !/^ev-\d+$/.test(e.id));
      }
      localStorage.removeItem('nexus_minimal_calendar_v3');
      savePersistedState();
      return Boolean(state.studentName);
    }

    const raw = localStorage.getItem('nexus_live_calendar_v4');
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    if (parsed.studentName) state.studentName = parsed.studentName;
    if (parsed.categoryColors) {
      state.categoryColors = { ...state.categoryColors, ...parsed.categoryColors };
    }
    if (parsed.integrationsAuth) {
      Object.keys(INTEGRATIONS).forEach(k => {
        if (parsed.integrationsAuth[k]) {
          state.integrationsAuth[k] = { ...state.integrationsAuth[k], ...parsed.integrationsAuth[k] };
        }
      });
    }
    if (Array.isArray(parsed.events)) {
      state.events = parsed.events.filter(e => !/^ev-\d+$/.test(e.id));
    }
    if (!parsed.pastCompletedMigrated) {
      state.events = state.events.map(e => (e.date < todayStr ? { ...e, completed: true } : e));
      savePersistedState();
    }
    return Boolean(parsed.studentName);
  } catch (e) {
    return false;
  }
}

function savePersistedState() {
  try {
    localStorage.setItem('nexus_live_calendar_v4', JSON.stringify({
      studentName: state.studentName,
      categoryColors: state.categoryColors,
      integrationsAuth: state.integrationsAuth,
      events: state.events,
      pastCompletedMigrated: true
    }));
  } catch (e) {}
}

// =========================================================
// LIVE BACKEND SYNC (/api/sync)
// =========================================================
async function triggerLiveBackendSync() {
  const banner = document.getElementById('syncStatusBanner');
  const sideStatus = document.getElementById('sidebarSyncStatus');

  if (banner) {
    banner.style.display = 'block';
    banner.textContent = 'Syncing with connected services...';
  }
  if (sideStatus) {
    sideStatus.textContent = 'Syncing...';
  }

  try {
    const resp = await fetch('/api/sync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ integrations: state.integrationsAuth })
    });
    const data = await resp.json();

    if (Array.isArray(data.events)) {
      // Preserve completion status & custom colors of previously synced items, replace stale synced items
      const prevMap = new Map(state.events.map(e => [e.id, e]));
      const manualEvents = state.events.filter(e => e.id.startsWith('manual-') || e.id.startsWith('ics-local-'));

      const mergedSynced = data.events.map(ev => {
        const prev = prevMap.get(ev.id);
        return prev
          ? {
              ...ev,
              completed: prev.completed,
              customColor: prev.customColor,
              category: prev.categoryOverridden ? prev.category : ev.category,
              categoryOverridden: prev.categoryOverridden
            }
          : ev;
      });

      state.events = [...manualEvents, ...mergedSynced];
      savePersistedState();

      // Ensure connected integrations are visible in filter
      Object.entries(state.integrationsAuth).forEach(([id, auth]) => {
        if (auth.connected) state.activeIntegrations.add(id);
      });

      // If current view month has no events but synced events exist, jump to the closest event date
      if (mergedSynced.length > 0) {
        const ymPrefix = `${state.viewDate.getFullYear()}-${String(state.viewDate.getMonth() + 1).padStart(2, '0')}`;
        const hasInCurrentMonth = state.events.some(e => e.date.startsWith(ymPrefix));
        if (!hasInCurrentMonth) {
          const sortedDates = mergedSynced.map(e => e.date).sort();
          const upcomingOrLatest = sortedDates.find(d => d >= todayStr) || sortedDates[sortedDates.length - 1];
          if (upcomingOrLatest) {
            const [y, m, d] = upcomingOrLatest.split('-').map(Number);
            state.viewDate = new Date(y, m - 1, d);
            state.selectedDateStr = upcomingOrLatest;
          }
        }
      }
    }

    state.syncStatuses = data.statuses || {};
    const liveTotal = Array.isArray(data.events) ? data.events.length : 0;

    const statusLines = Object.entries(state.syncStatuses)
      .filter(([, st]) => !st.skipped)
      .map(([k, st]) => {
        const name = INTEGRATIONS[k]?.name || k;
        return st.ok ? `${name}: ${st.count} synced` : `${name}: ${st.message || 'error'}`;
      });

    if (banner) {
      banner.style.display = 'block';
      banner.textContent = statusLines.length > 0
        ? `Synced ${liveTotal} total — ${statusLines.join(' | ')}`
        : 'No integrations enabled yet. Click "Use Login" on a service above and enter your credentials.';
    }
    if (sideStatus) {
      const firstErr = Object.entries(state.syncStatuses).find(([, st]) => !st.skipped && !st.ok);
      if (firstErr) {
        sideStatus.textContent = `${INTEGRATIONS[firstErr[0]]?.name}: ${firstErr[1].message}`;
      } else {
        sideStatus.textContent = `Synced ${liveTotal} live items`;
      }
    }

    renderAll();
  } catch (err) {
    if (banner) {
      banner.style.display = 'block';
      banner.textContent = 'Server unreachable. Open http://localhost:8080 with server.py running.';
    }
    if (sideStatus) {
      sideStatus.textContent = 'Run python3 server.py';
    }
  }
}

// =========================================================
// ONBOARDING / LOGIN
// =========================================================
function renderOnboardingIntegrations() {
  const list = document.getElementById('onboardingIntegrationsList');
  if (!list) return;

  list.innerHTML = Object.values(INTEGRATIONS).map(intg => {
    const auth = state.integrationsAuth[intg.id] || { connected: false };
    const isLms = intg.id === 'canvas' || intg.id === 'canvas2' || intg.id === 'vva';
    const field1Val = isLms ? (auth.domain || '') : (auth.username || '');
    const field2Val = auth.token || auth.password || '';
    const icsVal = auth.icsUrl || '';

    return `
      <div class="integration-row ${auth.connected ? 'connected' : ''}">
        <div class="integration-row-top">
          <div class="integration-label">
            <span class="circle-dot"></span>
            <span>${intg.name}</span>
            <span class="integration-sub">${intg.sub}</span>
          </div>
          <button type="button" class="btn-flat" data-toggle-int="${intg.id}">
            ${auth.connected ? 'Connected' : 'Use Login'}
          </button>
        </div>
        <div class="integration-row-inputs">
          <div class="form-row-2">
            <input
              type="text"
              class="minimal-input"
              data-int-f1="${intg.id}"
              placeholder="${intg.field1Label}"
              value="${field1Val}"
            />
            <input
              type="password"
              class="minimal-input"
              data-int-f2="${intg.id}"
              placeholder="${intg.field2Label}"
              value="${field2Val}"
            />
          </div>
          <input
            type="text"
            class="minimal-input"
            data-int-ics="${intg.id}"
            placeholder="${intg.icsLabel}"
            value="${icsVal}"
          />
        </div>
      </div>
    `;
  }).join('');

  list.querySelectorAll('[data-toggle-int]').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.getAttribute('data-toggle-int');
      const auth = state.integrationsAuth[id];
      auth.connected = !auth.connected;
      if (auth.connected) state.activeIntegrations.add(id);
      savePersistedState();
      renderOnboardingIntegrations();
      updateOnboardingSummaryCount();
    });
  });

  list.querySelectorAll('[data-int-f1]').forEach(input => {
    input.addEventListener('input', (e) => {
      const id = input.getAttribute('data-int-f1');
      const val = e.target.value.trim();
      if (id === 'canvas' || id === 'canvas2' || id === 'vva') {
        state.integrationsAuth[id].domain = val;
      } else {
        state.integrationsAuth[id].username = val;
      }
      if (val) {
        state.integrationsAuth[id].connected = true;
        state.activeIntegrations.add(id);
        updateOnboardingSummaryCount();
      }
      savePersistedState();
    });
  });

  list.querySelectorAll('[data-int-f2]').forEach(input => {
    input.addEventListener('input', (e) => {
      const id = input.getAttribute('data-int-f2');
      const val = e.target.value.trim();
      state.integrationsAuth[id].token = val;
      state.integrationsAuth[id].password = val;
      if (val) {
        state.integrationsAuth[id].connected = true;
        state.activeIntegrations.add(id);
        updateOnboardingSummaryCount();
      }
      savePersistedState();
    });
  });

  list.querySelectorAll('[data-int-ics]').forEach(input => {
    input.addEventListener('input', (e) => {
      const id = input.getAttribute('data-int-ics');
      const val = e.target.value.trim();
      state.integrationsAuth[id].icsUrl = val;
      if (val) {
        state.integrationsAuth[id].connected = true;
        state.activeIntegrations.add(id);
        updateOnboardingSummaryCount();
      }
      savePersistedState();
    });
  });
}

function updateOnboardingSummaryCount() {
  const count = Object.values(state.integrationsAuth).filter(a => a.connected).length;
  const el = document.getElementById('connectedCountSummary');
  if (el) el.textContent = `${count} of 6 logins active`;
}

function openOnboardingModal() {
  const overlay = document.getElementById('onboardingOverlay');
  const nameInput = document.getElementById('studentNameInput');
  if (nameInput) nameInput.value = state.studentName || '';
  renderOnboardingIntegrations();
  updateOnboardingSummaryCount();
  overlay.classList.remove('hidden');
  setTimeout(() => nameInput && nameInput.focus(), 50);
}

function closeOnboardingAndLaunch() {
  const nameInput = document.getElementById('studentNameInput');
  const enteredName = (nameInput?.value || '').trim();
  if (!enteredName) {
    nameInput.style.borderColor = '#6366f1';
    nameInput.placeholder = 'Please enter your name...';
    nameInput.focus();
    return;
  }

  state.studentName = enteredName;
  Object.entries(state.integrationsAuth).forEach(([id, auth]) => {
    if (auth.connected) state.activeIntegrations.add(id);
  });

  savePersistedState();
  document.getElementById('onboardingOverlay').classList.add('hidden');
  renderAll();

  const anyConnected = Object.values(state.integrationsAuth).some(a => a.connected);
  if (anyConnected) {
    triggerLiveBackendSync();
  }
}

// =========================================================
// FILTERING & FORMATTERS
// =========================================================
function getFilteredEvents() {
  const q = state.searchQuery.toLowerCase().trim();
  return state.events.filter(ev => {
    if (!state.activeCategories.has(ev.category)) return false;
    if (!state.activeIntegrations.has(ev.integration)) return false;
    if (q) {
      const hay = `${ev.title} ${ev.location || ''} ${CATEGORIES[ev.category]?.name || ''} ${INTEGRATIONS[ev.integration]?.name || ''}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  }).sort((a, b) => `${a.date}T${a.time}`.localeCompare(`${b.date}T${b.time}`));
}

function formatReadableDate(dateStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  return new Date(y, m - 1, d).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

function formatWeekday(dateStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  return new Date(y, m - 1, d).toLocaleDateString('en-US', { weekday: 'long' }).toUpperCase();
}

// =========================================================
// SIDEBARS (With Custom Activity Color Pickers)
// =========================================================
function renderLeftSidebar() {
  document.getElementById('studentNameTop').textContent = state.studentName || 'Student';

  const catList = document.getElementById('categoryFilterList');
  catList.innerHTML = Object.values(CATEGORIES).map(cat => {
    const count = state.events.filter(e => e.category === cat.id).length;
    const active = state.activeCategories.has(cat.id);
    const color = state.categoryColors[cat.id] || cat.defaultColor;

    return `
      <div class="filter-item ${active ? '' : 'inactive'}">
        <div class="filter-left" data-cat-filter="${cat.id}">
          <span class="node-circle" style="border-color:${color}; ${active ? `background:${color};` : ''}"></span>
          <span>${cat.name}</span>
        </div>
        <div class="filter-right">
          <label class="color-picker-wrap" style="background:${color};" title="Customize ${cat.name} color">
            <input type="color" class="color-picker-input" data-cat-color="${cat.id}" value="${color}" />
          </label>
          <span class="filter-count">${count}</span>
        </div>
      </div>
    `;
  }).join('');

  catList.querySelectorAll('[data-cat-filter]').forEach(el => {
    el.addEventListener('click', () => {
      const id = el.getAttribute('data-cat-filter');
      if (state.activeCategories.has(id)) state.activeCategories.delete(id);
      else state.activeCategories.add(id);
      renderAll();
    });
  });

  catList.querySelectorAll('[data-cat-color]').forEach(input => {
    input.addEventListener('input', (e) => {
      const id = input.getAttribute('data-cat-color');
      state.categoryColors[id] = e.target.value;
      savePersistedState();
      renderAll();
    });
  });

  const intList = document.getElementById('integrationFilterList');
  intList.innerHTML = Object.values(INTEGRATIONS).map(intg => {
    const auth = state.integrationsAuth[intg.id];
    const connected = auth && auth.connected;
    const active = state.activeIntegrations.has(intg.id);
    const count = state.events.filter(e => e.integration === intg.id).length;
    return `
      <div class="filter-item ${active && connected ? '' : 'inactive'}">
        <div class="filter-left" data-int-filter="${intg.id}">
          <span class="node-circle" style="${connected && active ? 'background:var(--accent);' : ''}"></span>
          <span>${intg.name}</span>
        </div>
        <span class="filter-count">${connected ? count : 'off'}</span>
      </div>
    `;
  }).join('');

  intList.querySelectorAll('[data-int-filter]').forEach(el => {
    el.addEventListener('click', () => {
      const id = el.getAttribute('data-int-filter');
      if (!state.integrationsAuth[id].connected) {
        openOnboardingModal();
        return;
      }
      if (state.activeIntegrations.has(id)) state.activeIntegrations.delete(id);
      else state.activeIntegrations.add(id);
      renderAll();
    });
  });
}

// =========================================================
// VIEWS
// =========================================================
function renderMonthView(container, events) {
  const year = state.viewDate.getFullYear();
  const month = state.viewDate.getMonth();
  const startDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const prevMonthDays = new Date(year, month, 0).getDate();

  const cells = [];
  for (let i = startDay - 1; i >= 0; i--) {
    const d = new Date(year, month - 1, prevMonthDays - i);
    cells.push({ dateStr: offsetDateStr(d, 0), dayNum: d.getDate(), outside: true });
  }
  for (let d = 1; d <= daysInMonth; d++) {
    const dt = new Date(year, month, d);
    cells.push({ dateStr: offsetDateStr(dt, 0), dayNum: d, outside: false });
  }
  const totalSlots = cells.length > 35 ? 42 : 35;
  let nextDay = 1;
  while (cells.length < totalSlots) {
    const dt = new Date(year, month + 1, nextDay++);
    cells.push({ dateStr: offsetDateStr(dt, 0), dayNum: dt.getDate(), outside: true });
  }

  const weekdays = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  container.innerHTML = `
    <div class="month-grid">
      <div class="weekday-row">
        ${weekdays.map(w => `<div class="weekday-cell">${w}</div>`).join('')}
      </div>
      <div class="days-matrix">
        ${cells.map(cell => {
          const dayEvents = events.filter(e => e.date === cell.dateStr);
          const visible = dayEvents.slice(0, 3);
          const more = dayEvents.length - visible.length;
          return `
            <div class="day-cell ${cell.outside ? 'outside' : ''} ${cell.dateStr === todayStr ? 'today' : ''} ${cell.dateStr === state.selectedDateStr ? 'selected' : ''}" data-select-date="${cell.dateStr}">
              <div class="day-cell-top">
                <span class="day-num">${cell.dayNum}</span>
              </div>
              <div class="day-events-stack">
                ${visible.map(ev => {
                  const col = getEventColor(ev);
                  return `<div class="day-event-line ${ev.completed ? 'done' : ''}" style="border-left-color:${col};">${ev.time} ${ev.title}</div>`;
                }).join('')}
                ${more > 0 ? `<div class="day-more">+${more} more</div>` : ''}
              </div>
            </div>
          `;
        }).join('')}
      </div>
    </div>
  `;

  container.querySelectorAll('[data-select-date]').forEach(cell => {
    cell.addEventListener('click', () => {
      state.selectedDateStr = cell.getAttribute('data-select-date');
      renderAll();
    });
  });
}

function renderWeekView(container, events) {
  const curr = new Date(state.viewDate);
  const sunday = new Date(curr.getFullYear(), curr.getMonth(), curr.getDate() - curr.getDay());
  const weekDays = [];
  for (let i = 0; i < 7; i++) {
    const dt = new Date(sunday.getFullYear(), sunday.getMonth(), sunday.getDate() + i);
    weekDays.push({
      dateStr: offsetDateStr(dt, 0),
      label: dt.toLocaleDateString('en-US', { weekday: 'short' }) + ' ' + dt.getDate()
    });
  }

  const hours = [8, 10, 12, 14, 16, 18, 20, 22];

  let html = `<div class="week-table">`;
  html += `<div class="week-head-cell"></div>`;
  weekDays.forEach(wd => {
    html += `<div class="week-head-cell ${wd.dateStr === todayStr ? 'today' : ''}" data-select-date="${wd.dateStr}">
      <span class="meta-mono">${wd.label}</span>
    </div>`;
  });

  hours.forEach((h, idx) => {
    const nextH = hours[idx + 1] || 24;
    html += `<div class="time-cell">${String(h).padStart(2, '0')}:00</div>`;
    weekDays.forEach(wd => {
      const slotEvents = events.filter(e => {
        if (e.date !== wd.dateStr) return false;
        const hr = parseInt((e.time || '12:00').split(':')[0], 10);
        return idx === 0 ? hr < nextH : (hr >= h && hr < nextH);
      });
      html += `<div class="week-slot-cell" data-select-date="${wd.dateStr}">
        ${slotEvents.map(ev => {
          const col = getEventColor(ev);
          return `<div class="week-item" style="border-left-color:${col};">${ev.time} ${ev.title}</div>`;
        }).join('')}
      </div>`;
    });
  });

  html += `</div>`;
  container.innerHTML = html;

  container.querySelectorAll('[data-select-date]').forEach(el => {
    el.addEventListener('click', () => {
      state.selectedDateStr = el.getAttribute('data-select-date');
      renderRightInspector();
    });
  });
}

function renderTodoView(container, events) {
  const due = events.filter(e => !e.completed && e.category === 'duedates');
  const other = events.filter(e => !e.completed && e.category !== 'duedates');
  const done = events.filter(e => e.completed);

  const renderCol = (title, items) => `
    <div class="todo-col">
      <div class="todo-col-title">
        <span>${title}</span>
        <span>${items.length}</span>
      </div>
      ${items.length === 0 ? `<div class="meta-mono" style="padding:12px 0;">No activities.</div>` : ''}
      ${items.map(ev => {
        const col = getEventColor(ev);
        return `
          <div class="todo-row" style="border-left-color:${col};" data-select-date="${ev.date}">
            <button type="button" class="circle-check ${ev.completed ? 'checked' : ''}" style="border-color:${col}; ${ev.completed ? `background:${col};` : ''}" data-toggle-id="${ev.id}"></button>
            <div style="min-width:0; flex:1;">
              <div style="${ev.completed ? 'text-decoration:line-through;color:var(--muted);' : ''}">${ev.title}</div>
              <div class="meta-mono">${ev.date} · ${ev.time} · ${INTEGRATIONS[ev.integration]?.name || ''}</div>
            </div>
          </div>
        `;
      }).join('')}
    </div>
  `;

  container.innerHTML = `
    <div class="todo-columns">
      ${renderCol('Due Dates', due)}
      ${renderCol('Clubs, Mentorship & Sports', other)}
      ${renderCol('Completed', done)}
    </div>
  `;

  container.querySelectorAll('[data-toggle-id]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleEventComplete(btn.getAttribute('data-toggle-id'));
    });
  });
}

function renderOrbitView(container, events) {
  const [y, m, d] = state.selectedDateStr.split('-').map(Number);
  const baseDate = new Date(y, m - 1, d);

  const upcoming = events.map(ev => {
    const [ey, em, ed] = ev.date.split('-').map(Number);
    const diffDays = Math.round((new Date(ey, em - 1, ed) - baseDate) / 86400000);
    return { ...ev, diffDays };
  }).filter(ev => ev.diffDays >= 0 && ev.diffDays <= 7);

  const nodesSvg = upcoming.map((ev, i) => {
    const r = ev.diffDays === 0 ? 70 : (ev.diffDays <= 3 ? 130 : 190);
    const [hh, mm] = (ev.time || '12:00').split(':').map(Number);
    const angle = (((hh * 60 + mm) / 1440) * 360 + i * 25 - 90) * (Math.PI / 180);
    const cx = 220 + r * Math.cos(angle);
    const cy = 220 + r * Math.sin(angle);
    const col = getEventColor(ev);
    return `
      <g data-orbit-id="${ev.id}" style="cursor:pointer;">
        <line x1="220" y1="220" x2="${cx.toFixed(1)}" y2="${cy.toFixed(1)}" stroke="#1e2233" stroke-width="1" />
        <circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}" r="6" fill="#090a0f" stroke="${col}" stroke-width="2" />
        <text x="${(cx + 10).toFixed(1)}" y="${(cy + 3).toFixed(1)}" fill="#eceff4" font-family="JetBrains Mono" font-size="10">
          ${ev.title.slice(0, 16)}
        </text>
      </g>
    `;
  }).join('');

  container.innerHTML = `
    <div class="orbit-layout">
      <div class="orbit-frame">
        <svg class="orbit-svg" viewBox="0 0 440 440">
          <line x1="220" y1="20" x2="220" y2="420" stroke="#1e2233" stroke-width="1" />
          <line x1="20" y1="220" x2="420" y2="220" stroke="#1e2233" stroke-width="1" />
          <circle cx="220" cy="220" r="70" fill="none" stroke="#1e2233" stroke-width="1" />
          <circle cx="220" cy="220" r="130" fill="none" stroke="#1e2233" stroke-width="1" />
          <circle cx="220" cy="220" r="190" fill="none" stroke="#1e2233" stroke-width="1" />
          ${nodesSvg}
          <circle cx="220" cy="220" r="8" fill="#6366f1" />
        </svg>
      </div>
      <div class="inspector-list">
        <div class="sidebar-label"><span>7-Day Orbit (${upcoming.length})</span></div>
        ${upcoming.length === 0 ? `<div class="meta-mono" style="padding:12px 0;">No activities in next 7 days.</div>` : ''}
        ${upcoming.map(ev => {
          const col = getEventColor(ev);
          return `
            <div class="inspector-item" style="border-left-color:${col}; cursor:pointer;" data-orbit-id="${ev.id}">
              <div>${ev.title}</div>
              <div class="meta-mono">${ev.date} · ${ev.time} · ${INTEGRATIONS[ev.integration]?.name || ''}</div>
            </div>
          `;
        }).join('')}
      </div>
    </div>
  `;

  container.querySelectorAll('[data-orbit-id]').forEach(el => {
    el.addEventListener('click', () => {
      const ev = state.events.find(x => x.id === el.getAttribute('data-orbit-id'));
      if (ev) {
        state.selectedDateStr = ev.date;
        renderRightInspector();
      }
    });
  });
}

function renderTimelineView(container, events) {
  const grouped = {};
  events.filter(ev => !ev.completed).forEach(ev => {
    if (!grouped[ev.date]) grouped[ev.date] = [];
    grouped[ev.date].push(ev);
  });

  const dates = Object.keys(grouped).sort();
  if (dates.length === 0) {
    container.innerHTML = `<div class="meta-mono" style="text-align:center; padding:48px 0;">No synced or scheduled activities yet. Click "Logins" or "+ New" to begin.</div>`;
    return;
  }

  container.innerHTML = `
    <div class="stream-list">
      ${dates.map(d => `
        <div class="stream-day">
          <div class="meta-mono">${formatReadableDate(d)} — ${formatWeekday(d)}</div>
          ${grouped[d].map(ev => {
            const col = getEventColor(ev);
            return `
              <div class="stream-item" style="border-left-color:${col};" data-select-date="${ev.date}">
                <div style="display:flex;align-items:center;gap:10px;">
                  <button type="button" class="circle-check ${ev.completed ? 'checked' : ''}" style="border-color:${col}; ${ev.completed ? `background:${col};` : ''}" data-toggle-id="${ev.id}"></button>
                  <span style="${ev.completed ? 'text-decoration:line-through;color:var(--muted);' : ''}">${ev.title}</span>
                </div>
                <span class="meta-mono">${ev.time} · ${CATEGORIES[ev.category]?.name || ''} · ${INTEGRATIONS[ev.integration]?.name || ''}</span>
              </div>
            `;
          }).join('')}
        </div>
      `).join('')}
    </div>
  `;

  container.querySelectorAll('[data-toggle-id]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleEventComplete(btn.getAttribute('data-toggle-id'));
    });
  });
}

// =========================================================
// RIGHT INSPECTOR (Includes Individual Activity Color Picker)
// =========================================================
function renderRightInspector() {
  const d = state.selectedDateStr;
  document.getElementById('inspectorWeekday').textContent = formatWeekday(d);
  document.getElementById('inspectorDateTitle').textContent = formatReadableDate(d);

  const dayEvents = getFilteredEvents().filter(e => e.date === d);
  const listEl = document.getElementById('inspectorEventsList');

  if (dayEvents.length === 0) {
    listEl.innerHTML = `<div class="meta-mono" style="padding:12px 0;">No events scheduled.</div>`;
    return;
  }

  listEl.innerHTML = dayEvents.map(ev => {
    const col = getEventColor(ev);
    return `
      <div class="inspector-item" style="border-left-color:${col};">
        <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:8px;">
          <div style="display:flex;align-items:flex-start;gap:8px;min-width:0;flex:1;">
            <button type="button" class="circle-check ${ev.completed ? 'checked' : ''}" style="border-color:${col}; ${ev.completed ? `background:${col};` : ''}" data-inspector-check="${ev.id}"></button>
            <div style="min-width:0;flex:1;">
              <div style="${ev.completed ? 'text-decoration:line-through;color:var(--muted);' : ''}">${ev.title}</div>
              <div class="meta-mono">${ev.time} · ${INTEGRATIONS[ev.integration]?.name || ''}</div>
              <div style="margin-top:4px;">
                <select class="inspector-cat-select" data-ev-cat="${ev.id}" title="Change classification">
                  ${Object.values(CATEGORIES).map(c => `<option value="${c.id}" ${ev.category === c.id ? 'selected' : ''}>${c.name}</option>`).join('')}
                </select>
              </div>
              ${ev.location ? `<div class="meta-mono" style="margin-top:2px;">${ev.location}</div>` : ''}
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:6px;">
            <label class="color-picker-wrap" style="background:${col};" title="Customize this activity's color">
              <input type="color" class="color-picker-input" data-ev-color="${ev.id}" value="${col}" />
            </label>
            <button type="button" data-delete-event="${ev.id}" style="background:none;border:none;color:var(--muted);cursor:pointer;">✕</button>
          </div>
        </div>
      </div>
    `;
  }).join('');

  listEl.querySelectorAll('[data-inspector-check]').forEach(btn => {
    btn.addEventListener('click', () => toggleEventComplete(btn.getAttribute('data-inspector-check')));
  });

  listEl.querySelectorAll('[data-ev-cat]').forEach(select => {
    select.addEventListener('change', (e) => {
      const id = select.getAttribute('data-ev-cat');
      const ev = state.events.find(x => x.id === id);
      if (ev) {
        ev.category = e.target.value;
        ev.categoryOverridden = true;
        savePersistedState();
        renderAll();
      }
    });
  });

  listEl.querySelectorAll('[data-ev-color]').forEach(input => {
    input.addEventListener('input', (e) => {
      const id = input.getAttribute('data-ev-color');
      const ev = state.events.find(x => x.id === id);
      if (ev) {
        ev.customColor = e.target.value;
        savePersistedState();
        renderAll();
      }
    });
  });

  listEl.querySelectorAll('[data-delete-event]').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.getAttribute('data-delete-event');
      state.events = state.events.filter(e => e.id !== id);
      savePersistedState();
      renderAll();
    });
  });
}

function toggleEventComplete(id) {
  const ev = state.events.find(e => e.id === id);
  if (ev) {
    ev.completed = !ev.completed;
    savePersistedState();
    renderAll();
  }
}

// =========================================================
// LOCAL .ICS FILE IMPORT
// =========================================================
function parseIcsContent(icsText) {
  const unfolded = icsText.replace(/\r?\n[ \t]/g, '');
  const blocks = unfolded.split('BEGIN:VEVENT').slice(1);
  let imported = 0;

  blocks.forEach(block => {
    const summaryMatch = block.match(/^SUMMARY(?:;[^:]*)?:(.*)$/im);
    const dtStartMatch = block.match(/^DTSTART(?:;[^:]*)?:(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2}))?/im);
    const locationMatch = block.match(/^LOCATION(?:;[^:]*)?:(.*)$/im);
    if (!summaryMatch || !dtStartMatch) return;

    const title = summaryMatch[1].trim().replace(/\\,/g, ',').replace(/\\n/g, ' ');
    const date = `${dtStartMatch[1]}-${dtStartMatch[2]}-${dtStartMatch[3]}`;
    const time = dtStartMatch[4] ? `${dtStartMatch[4]}:${dtStartMatch[5]}` : '23:59';
    const location = locationMatch ? locationMatch[1].trim().replace(/\\,/g, ',') : 'iCal Import';

    state.events.push({
      id: 'ics-local-' + Date.now() + '-' + imported++,
      title,
      category: 'duedates',
      integration: 'canvas',
      date,
      time,
      location,
      completed: false
    });
  });

  if (imported > 0) {
    savePersistedState();
    renderAll();
  }
}

function renderAll() {
  renderLeftSidebar();
  document.getElementById('currentPeriodLabel').textContent =
    state.viewDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  const container = document.getElementById('viewContainer');
  const events = getFilteredEvents();

  if (state.currentView === 'month') renderMonthView(container, events);
  else if (state.currentView === 'week') renderWeekView(container, events);
  else if (state.currentView === 'todo') renderTodoView(container, events);
  else if (state.currentView === 'orbit') renderOrbitView(container, events);
  else if (state.currentView === 'timeline') renderTimelineView(container, events);

  renderRightInspector();
}

document.addEventListener('DOMContentLoaded', () => {
  const hasSavedUser = loadPersistedState();
  renderOnboardingIntegrations();
  updateOnboardingSummaryCount();

  if (hasSavedUser) {
    document.getElementById('onboardingOverlay').classList.add('hidden');
  }

  renderAll();

  document.getElementById('launchCalendarBtn').addEventListener('click', closeOnboardingAndLaunch);
  document.getElementById('testLiveSyncBtn').addEventListener('click', triggerLiveBackendSync);
  document.getElementById('syncNowTopBtn').addEventListener('click', triggerLiveBackendSync);
  document.getElementById('sidebarSyncBtn').addEventListener('click', triggerLiveBackendSync);
  document.getElementById('completePastBtn').addEventListener('click', () => {
    state.events = state.events.map(ev => (ev.date < todayStr ? { ...ev, completed: true } : ev));
    savePersistedState();
    renderAll();
  });

  document.getElementById('studentNameInput').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      closeOnboardingAndLaunch();
    }
  });

  document.getElementById('profileTrigger').addEventListener('click', openOnboardingModal);
  document.getElementById('manageLoginsBtn').addEventListener('click', openOnboardingModal);

  document.querySelectorAll('.view-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.view-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      state.currentView = tab.getAttribute('data-view');
      renderAll();
    });
  });

  document.getElementById('prevPeriodBtn').addEventListener('click', () => {
    if (state.currentView === 'week') state.viewDate.setDate(state.viewDate.getDate() - 7);
    else state.viewDate.setMonth(state.viewDate.getMonth() - 1);
    renderAll();
  });

  document.getElementById('nextPeriodBtn').addEventListener('click', () => {
    if (state.currentView === 'week') state.viewDate.setDate(state.viewDate.getDate() + 7);
    else state.viewDate.setMonth(state.viewDate.getMonth() + 1);
    renderAll();
  });

  document.getElementById('todayBtn').addEventListener('click', () => {
    state.viewDate = new Date(todayDate.getFullYear(), todayDate.getMonth(), todayDate.getDate());
    state.selectedDateStr = todayStr;
    renderAll();
  });

  document.getElementById('searchInput').addEventListener('input', (e) => {
    state.searchQuery = e.target.value;
    renderAll();
  });

  document.getElementById('resetCategoriesBtn').addEventListener('click', () => {
    Object.keys(CATEGORIES).forEach(k => state.activeCategories.add(k));
    renderAll();
  });

  const eventModal = document.getElementById('eventModal');
  const evCategorySelect = document.getElementById('evCategory');
  const evColorInput = document.getElementById('evColor');

  evCategorySelect.addEventListener('change', () => {
    const cat = evCategorySelect.value;
    evColorInput.value = state.categoryColors[cat] || '#6366f1';
  });

  const openModal = () => {
    document.getElementById('evDate').value = state.selectedDateStr;
    evColorInput.value = state.categoryColors[evCategorySelect.value] || '#6366f1';
    eventModal.classList.remove('hidden');
  };
  const closeModal = () => eventModal.classList.add('hidden');

  document.getElementById('openNewEventBtn').addEventListener('click', openModal);
  document.getElementById('quickAddSelectedDayBtn').addEventListener('click', openModal);
  document.getElementById('closeEventModalBtn').addEventListener('click', closeModal);
  document.getElementById('cancelEventModalBtn').addEventListener('click', closeModal);

  document.getElementById('addEventForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const cat = document.getElementById('evCategory').value;
    const pickedColor = document.getElementById('evColor').value;
    const defaultCatColor = state.categoryColors[cat] || '#6366f1';

    const newEv = {
      id: 'manual-' + Date.now(),
      title: document.getElementById('evTitle').value.trim(),
      category: cat,
      integration: document.getElementById('evIntegration').value,
      date: document.getElementById('evDate').value || state.selectedDateStr,
      time: document.getElementById('evTime').value || '15:00',
      location: document.getElementById('evLocation').value.trim(),
      customColor: pickedColor !== defaultCatColor ? pickedColor : undefined,
      completed: false
    };
    state.events.push(newEv);
    state.selectedDateStr = newEv.date;
    savePersistedState();
    e.target.reset();
    closeModal();
    renderAll();
  });

  document.getElementById('icsFileInput').addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => parseIcsContent(String(reader.result || ''));
    reader.readAsText(file);
  });
});
