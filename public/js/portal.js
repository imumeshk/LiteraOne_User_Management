/* ═══════════════════════════════════════════════════════════════
   Litera One Manager Portal — Frontend Application
   All API calls go through the Node.js server (no direct Graph calls)
═══════════════════════════════════════════════════════════════ */
'use strict';

const DEPLOY_DEFAULT_URLS = {
  Outlook: 'https://marketplace.microsoft.com/en-us/product/office/wa200008417',
  Word: 'https://marketplace.microsoft.com/en-us/product/WA200008193',
};

// ─── Global State ─────────────────────────────────────────────────────────────
const S = {
  isAuthenticated: !!window.__IS_AUTHENTICATED__,
  // Entra Apps
  allApps: [],
  selectedApp: null,
  selectedAppRaw: null,            // full detail {sp, permissions, hasConsent, isHidden}
  assignments: [],                 // enriched assignments for selected app
  appAssignmentsCache: {},         // { [spId]: assignment[] }

  // Assign modal
  assignTab: 'users',
  assignSelected: [],              // [{id, name, email, type}]
  assignPaging: { nextLink: null, items: [] },

  // Deploy
  deployTargets: [],               // [{id, name, emailOrId, type}]
  deployAssignTab: 'users',
  deployAssignSelected: [],
  deployAssignPaging: { nextLink: null, items: [] },
  deployStatusItems: [],
  deployManifestFileContent: '',
  deployManifestFileName: '',

  // SCIM wizard
  scimStep: 1,
  scimAssignTab: 'users',
  scimAssignSelected: [],
  scimAssignPaging: { nextLink: null, items: [] },

  // Litera One Users Management
  usersManagementMode: 'assign',
  lumTab: 'users',
  lumSelected: [],
  lumPaging: { nextLink: null, items: [] },
  lumApps: [],
  userManagementRecords: [],

  // Enablement
  enFilter: 'All',
  enResources: [],
  enExportButtons: [],

  // Console
  logs: [],

  // Timers
  _assignTimer: null,
  _deployAssignTimer: null,
  _scimAssignTimer: null,
  _lumAssignTimer: null,
  _assignmentPollTimer: null,
  enablementAdminConfig: null,
  selectedEnablementResourceType: null,
  enablementJsonVisible: false,
  enablementAdminEditResourceId: null,
  enablementAdminEditExportButtonId: null,
};

// ─── API Helpers ───────────────────────────────────────────────────────────────
async function api(method, path, body) {
  const csrfToken = window.__CSRF_TOKEN__ || '';
  const opts = {
    method,
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
    credentials: 'same-origin',
  };
  if (csrfToken) opts.headers['x-csrf-token'] = csrfToken;
  if (body !== undefined) opts.body = JSON.stringify(body);
  const res = await fetch('/api' + path, opts);
  if (res.status === 401) throw new Error('Not authenticated. Click Sign In to connect.');
  const ct = res.headers.get('content-type') || '';
  if (ct.includes('text/csv') || ct.includes('text/html')) {
    // Binary download — handled separately
    return res;
  }
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
  return data;
}
const GET  = (p)     => api('GET', p);
const POST = (p, b)  => api('POST', p, b);
const DEL  = (p)     => api('DELETE', p);
const PATCH = (p, b) => api('PATCH', p, b);
const getAssignments = (spId) => GET(`/apps/${spId}/assignments?_ts=${Date.now()}`);
const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function getAssignmentsWithRetry(spId, { attempts = 4, delayMs = 900, minCount = null } = {}) {
  let lastData = { assignments: [] };
  let lastError = null;
  for (let i = 0; i < attempts; i++) {
    try {
      const data = await getAssignments(spId);
      const assignments = data?.assignments || [];
      lastData = { assignments };
      if (minCount == null || assignments.length >= minCount || i === attempts - 1) {
        return lastData;
      }
    } catch (e) {
      lastError = e;
      if (i === attempts - 1) throw e;
    }
    await wait(delayMs);
  }
  if (lastError) throw lastError;
  return lastData;
}

async function refreshSelectedAssignments({ minCount = null, logLabel = 'Assignments reloaded', silent = false } = {}) {
  if (!S.selectedApp) return;
  try {
    const before = (S.assignments || []).length;
    const data = await getAssignmentsWithRetry(S.selectedApp.id, { minCount });
    S.assignments = data.assignments || [];
    S.appAssignmentsCache[S.selectedApp.id] = S.assignments.slice();
    renderAssignments();
    if (!silent && (before !== S.assignments.length || logLabel)) {
      log(`${logLabel}: ${S.assignments.length}`);
    }
  } catch (e) {
    log(`Assignments refresh error: ${e.message}`);
  }
}

function stopAssignmentAutoRefresh() {
  if (S._assignmentPollTimer) {
    clearInterval(S._assignmentPollTimer);
    S._assignmentPollTimer = null;
  }
}

function startAssignmentAutoRefresh() {
  stopAssignmentAutoRefresh();
  if (!S.selectedApp) return;
  S._assignmentPollTimer = setInterval(() => {
    if (!S.selectedApp) return;
    refreshSelectedAssignments({ silent: true });
  }, 5000);
}

async function downloadApi(method, path, body, filename) {
  const csrfToken = window.__CSRF_TOKEN__ || '';
  const opts = { method, headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' }, credentials: 'same-origin' };
  if (csrfToken) opts.headers['x-csrf-token'] = csrfToken;
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch('/api' + path, opts);
  if (!res.ok) {
    if (res.status === 401) toast('Please sign in to continue.', 'error');
    else toast('Download failed', 'error');
    return;
  }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

// ─── Navigation ────────────────────────────────────────────────────────────────
function navTo(page) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById('page-' + page)?.classList.add('active');
  document.querySelectorAll('.nav-btn[data-page]').forEach(b => {
    b.classList.toggle('active', b.dataset.page === page);
  });
  if (page === 'enablement') loadEnablement();
  if (page === 'enablement-admin') refreshEnablementAdmin();
  if (page === 'deploy') populateCopySourceDrop();
  if (page === 'apps' && !S.allApps.length) loadApps();
  if (page === 'apps' && S.selectedApp) startAssignmentAutoRefresh();
  if (page !== 'apps') stopAssignmentAutoRefresh();
  if (page === 'users-management' && S.isAuthenticated) loadUsersManagementApps();
}

document.querySelectorAll('.nav-btn[data-page]').forEach(btn => {
  btn.addEventListener('click', () => navTo(btn.dataset.page));
});

// ─── Toast ─────────────────────────────────────────────────────────────────────
function toast(msg, type) {
  const el = document.createElement('div');
  el.className = 'toast-msg' + (type ? ' ' + type : '');
  el.textContent = msg;
  document.getElementById('toastContainer').appendChild(el);
  setTimeout(() => el.remove(), 3500);
}

// ─── Console / Logging ─────────────────────────────────────────────────────────
function log(msg) {
  const ts = new Date().toLocaleTimeString();
  const line = `[${ts}] ${msg}`;
  S.logs.push(line);
  const el = document.getElementById('consoleLog');
  if (el) {
    el.textContent += line + '\n';
    el.scrollTop = el.scrollHeight;
  }
}

function clearConsole() {
  const el = document.getElementById('consoleLog');
  if (el) el.textContent = '';
  S.logs = [];
}
function exportLogs() {
  const blob = new Blob([S.logs.join('\n')], { type: 'text/plain' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = 'litera_one_logs.txt'; a.click();
}

// ─── Modal Helpers ─────────────────────────────────────────────────────────────
function openModal(id) { document.getElementById(id).style.display = 'flex'; }
function closeModal(id) { document.getElementById(id).style.display = 'none'; }

// Close modals on backdrop click
document.querySelectorAll('.modal-backdrop').forEach(bd => {
  bd.addEventListener('click', e => { if (e.target === bd) bd.style.display = 'none'; });
});

// ─── Utility ───────────────────────────────────────────────────────────────────
function escHtml(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function fmtType(t) {
  return t === 'User'
    ? '<span class="badge badge-blue">User</span>'
    : '<span class="badge badge-purple">Group</span>';
}
function toggleAll(tbodyId, masterCb) {
  document.querySelectorAll(`#${tbodyId} input[type=checkbox]`).forEach(c => c.checked = masterCb.checked);
}
function copyText(text) {
  navigator.clipboard.writeText(text).then(() => toast('Copied!')).catch(() => toast('Copy failed', 'error'));
}

function isAuthError(err) {
  const msg = String(err?.message || err || '').toLowerCase();
  return msg.includes('not authenticated') || msg.includes('session expired') || msg.includes('sign in');
}

function updateAppsAuthUi() {
  const scimBtn = document.getElementById('btnNewScimApp');
  const appDetailEmpty = document.getElementById('appDetailEmpty');
  const appSigninPrompt = document.getElementById('appSigninPrompt');
  const appList = document.getElementById('appList');
  const appSearchBtn = document.getElementById('btnAppsSearch');
  const usersMgmtNav = document.querySelector('.nav-btn[data-page="users-management"]');
  const usersMgmtPrompt = document.getElementById('usersManagementPrompt');
  const usersMgmtWorkspace = document.getElementById('usersManagementWorkspace');
  if (scimBtn) scimBtn.style.display = S.isAuthenticated ? 'inline-flex' : 'none';
  if (appSearchBtn) appSearchBtn.disabled = !S.isAuthenticated;
  if (appDetailEmpty) appDetailEmpty.style.display = S.isAuthenticated ? 'flex' : 'none';
  if (appSigninPrompt) appSigninPrompt.style.display = S.isAuthenticated ? 'none' : 'flex';
  if (usersMgmtNav) usersMgmtNav.style.display = S.isAuthenticated ? 'flex' : 'none';
  if (usersMgmtPrompt) usersMgmtPrompt.style.display = S.isAuthenticated ? 'none' : 'flex';
  if (usersMgmtWorkspace) usersMgmtWorkspace.style.display = S.isAuthenticated ? 'block' : 'none';
  if (!S.isAuthenticated && appList) {
    appList.innerHTML = '<div class="empty-state"><div class="em-icon">🔐</div><div>Please sign in to view and manage applications.</div></div>';
    stopAssignmentAutoRefresh();
  }
  if (!S.isAuthenticated && document.getElementById('page-users-management')?.classList.contains('active')) {
    navTo('dashboard');
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  ENTRA APPS
// ═══════════════════════════════════════════════════════════════════════════════

async function loadApps() {
  const q = document.getElementById('appSearch').value.trim() || 'Litera';
  const listEl = document.getElementById('appList');
  listEl.innerHTML = '<div class="empty-state"><span class="spinner spinner-dark"></span></div>';
  log('Searching apps: ' + q);
  try {
    const data = await GET(`/apps/search?q=${encodeURIComponent(q)}`);
    S.isAuthenticated = true;
    updateAppsAuthUi();
    S.allApps = data.apps || [];
    populateCopySourceDrop();
    if (!S.allApps.length) {
      listEl.innerHTML = `<div class="empty-state"><div class="em-icon">🔍</div><div>No apps found for "${escHtml(q)}"</div></div>`;
      return;
    }
    listEl.innerHTML = '';
    S.allApps.forEach(app => {
      const d = document.createElement('div');
      d.className = 'app-item';
      d.innerHTML = `<div class="app-name truncate">${escHtml(app.displayName)}</div><div class="app-id truncate">${escHtml(app.id)}</div>`;
      d.onclick = () => selectApp(app, d);
      listEl.appendChild(d);
    });
    log(`Found ${S.allApps.length} app(s).`);
  } catch (e) {
    log('ERROR loading apps: ' + e.message);
    if (isAuthError(e)) {
      S.isAuthenticated = false;
      updateAppsAuthUi();
      return;
    }
    toast(e.message, 'error');
    listEl.innerHTML = `<div class="empty-state"><div class="em-icon">⚠️</div><div>${escHtml(e.message)}</div></div>`;
  }
}

async function selectApp(app, el) {
  stopAssignmentAutoRefresh();
  document.querySelectorAll('.app-item').forEach(i => i.classList.remove('active'));
  el.classList.add('active');
  S.selectedApp = app;
  S.assignments = [];

  // Show skeleton
  document.getElementById('appDetail').style.display = 'block';
  document.getElementById('appDetailEmpty').style.display = 'none';
  document.getElementById('appSigninPrompt').style.display = 'none';
  document.getElementById('detailAppName').textContent = app.displayName;
  document.getElementById('detailAppId').textContent = `SP ID: ${app.id}`;
  document.getElementById('permsList').innerHTML = '<span class="perm-chip">Loading...</span>';
  document.getElementById('btnGrantConsent').style.display = 'none';
  document.getElementById('consentBadge').style.display = 'none';
  document.getElementById('noConsentBadge').style.display = 'none';
  document.getElementById('btnToggleVis').textContent = 'Checking...';
  document.getElementById('btnToggleVis').className = 'btn btn-sm btn-secondary';
  document.getElementById('assignTbody').innerHTML = '<tr><td colspan="4" class="tbl-empty"><span class="spinner spinner-dark"></span></td></tr>';

  log(`Loading details for: ${app.displayName}`);

  try {
    const [detail, assignments] = await Promise.all([
      GET(`/apps/${app.id}`),
      getAssignments(app.id),
    ]);

    S.selectedAppRaw = detail;
    S.assignments = assignments.assignments || [];
    S.appAssignmentsCache[app.id] = S.assignments.slice();

    // Permissions
    const perms = detail.permissions || [];
    document.getElementById('permsList').innerHTML = perms.length
      ? perms.map(p => `<span class="perm-chip" title="${escHtml(p.type)}">${escHtml(p.name)}</span>`).join('')
      : '<span class="perm-chip">No permissions found</span>';

    // Consent
    document.getElementById('detailAppId').textContent = `SP ID: ${app.id} · App ID: ${detail.sp?.appId || ''}`;
    if (detail.hasConsent) {
      document.getElementById('consentBadge').style.display = 'inline';
      document.getElementById('noConsentBadge').style.display = 'none';
      document.getElementById('btnGrantConsent').style.display = 'none';
    } else {
      document.getElementById('consentBadge').style.display = 'none';
      document.getElementById('noConsentBadge').style.display = 'inline-flex';
      document.getElementById('btnGrantConsent').style.display = 'inline-flex';
    }

    // Visibility
    updateVisButton(detail.isHidden);

    // Assignments
    renderAssignments();
    log(`Loaded ${S.assignments.length} assignment(s).`);
    startAssignmentAutoRefresh();
  } catch (e) {
    log('ERROR loading app details: ' + e.message);
    if (isAuthError(e)) {
      S.isAuthenticated = false;
      const scimBtn = document.getElementById('btnNewScimApp');
      if (scimBtn) scimBtn.style.display = 'none';
      document.getElementById('appDetail').style.display = 'none';
      document.getElementById('appDetailEmpty').style.display = 'none';
      document.getElementById('appSigninPrompt').style.display = 'flex';
      toast('Please sign in to view app details.', 'error');
      return;
    }
    toast(e.message, 'error');
  }
}

function updateVisButton(isHidden) {
  const btn = document.getElementById('btnToggleVis');
  const status = document.getElementById('visStatus');
  if (isHidden) {
    btn.textContent = '👁 Show in MyApps';
    btn.className = 'btn btn-sm btn-secondary';
    status.textContent = 'Status: Hidden (Recommended)';
    status.style.color = 'var(--accent3)';
  } else {
    btn.textContent = '🙈 Hide from MyApps';
    btn.className = 'btn btn-sm btn-warn';
    status.textContent = 'Status: Visible to Users';
    status.style.color = 'var(--danger)';
  }
}

function renderAssignments() {
  const q = (document.getElementById('assignFilter').value || '').toLowerCase();
  const filtered = q
    ? S.assignments.filter(a =>
        (a.principalDisplayName || '').toLowerCase().includes(q) ||
        (a.emailOrId || '').toLowerCase().includes(q))
    : S.assignments;

  const tb = document.getElementById('assignTbody');
  if (!filtered.length) {
    tb.innerHTML = '<tr><td colspan="4" class="tbl-empty">No assignments found</td></tr>';
    return;
  }
  tb.innerHTML = filtered.map((a, i) => `
    <tr>
      <td class="chk-col"><input type="checkbox" data-idx="${i}"></td>
      <td>
        <span class="link-text" onclick="viewPrincipal('${escHtml(a.principalId)}','${escHtml(a.principalType)}','${escHtml(a.principalDisplayName || '')}')"
              style="font-weight:600;cursor:pointer;color:var(--accent)">
          ${escHtml(a.principalDisplayName || '(unknown)')}
        </span>
      </td>
      <td class="mono truncate" style="max-width:200px;font-size:11.5px">${escHtml(a.emailOrId || a.principalId)}</td>
      <td>${fmtType(a.principalType)}</td>
    </tr>`).join('');
}

function filterAssignments() { renderAssignments(); }

async function grantConsent() {
  if (!S.selectedApp || !S.selectedAppRaw) return;
  const perms = S.selectedAppRaw.permissions || [];
  const appNameEl = document.getElementById('grantConsentAppName');
  const permsEl = document.getElementById('grantConsentPerms');
  const statusEl = document.getElementById('grantConsentStatus');
  if (appNameEl) appNameEl.textContent = `Application: ${S.selectedApp.displayName}`;
  if (permsEl) {
    permsEl.innerHTML = perms.length
      ? perms.map(p => `<span class="perm-chip" title="${escHtml(p.type || '')}">${escHtml(p.name || '')}</span>`).join('')
      : '<span class="text-xs">No explicit permission list found. Consent may still be required for configured API access.</span>';
  }
  if (statusEl) statusEl.textContent = '';
  openModal('grantConsentModal');
}

async function confirmGrantConsent() {
  if (!S.selectedApp) return;
  log('Granting admin consent for ' + S.selectedApp.displayName + '...');
  const statusEl = document.getElementById('grantConsentStatus');
  if (statusEl) statusEl.textContent = 'Granting consent...';
  try {
    const data = await POST(`/apps/${S.selectedApp.id}/consent`);
    const results = data.results || [];
    const failed = results.filter(r => r?.error);
    const okCount = results.length - failed.length;
    log('Consent results: ' + results.map(r => r.resource || r.error).join(', '));
    if (failed.length) {
      toast(`Admin consent partially completed (${okCount} success, ${failed.length} failed)`, 'error');
      if (statusEl) statusEl.textContent = `Partial success: ${okCount} granted, ${failed.length} failed.`;
    } else {
      toast('Admin consent granted successfully', 'success');
      if (statusEl) statusEl.textContent = `Success: ${okCount || 1} permission set(s) granted.`;
      closeModal('grantConsentModal');
    }

    // Refresh app detail state so badges/buttons reflect new consent status.
    const detail = await GET(`/apps/${S.selectedApp.id}`);
    S.selectedAppRaw = detail;
    if (detail.hasConsent) {
      document.getElementById('consentBadge').style.display = 'inline';
      document.getElementById('noConsentBadge').style.display = 'none';
      document.getElementById('btnGrantConsent').style.display = 'none';
    }
  } catch (e) {
    log('Consent error: ' + e.message);
    if (statusEl) statusEl.textContent = `Failed: ${e.message}`;
    toast('Admin consent failed', 'error');
  }
}

async function toggleVisibility() {
  if (!S.selectedApp || !S.selectedAppRaw) return;
  const currentlyHidden = S.selectedAppRaw.isHidden;
  const appId = S.selectedAppRaw.sp?.appId;
  log((currentlyHidden ? 'Showing' : 'Hiding') + ' app: ' + S.selectedApp.displayName);
  try {
    const data = await PATCH(`/apps/${S.selectedApp.id}/visibility`, { hide: !currentlyHidden, appId });
    S.selectedAppRaw.isHidden = data.isHidden;
    S.selectedAppRaw.tags = data.tags;
    updateVisButton(data.isHidden);
    toast('Visibility updated', 'success');
    log('Visibility tags: ' + (data.tags || []).join(', '));
  } catch (e) {
    log('Visibility error: ' + e.message);
    toast(e.message, 'error');
  }
}

async function checkSync() {
  if (!S.selectedApp) return;
  log('Checking sync jobs for ' + S.selectedApp.displayName + '...');
  try {
    const data = await GET(`/apps/${S.selectedApp.id}/sync-jobs`);
    const jobs = data.jobs || [];
    if (jobs.length) {
      const job = jobs[0];
      const status = job.status?.code || job.status?.state || 'Unknown';
      log(`Sync job: ${job.id} — Status: ${status} — Schedule: ${job.schedule?.state || 'N/A'}`);
      toast(`Sync status: ${status}`, status === 'Active' ? 'success' : undefined);
      const btnRestart = document.getElementById('btnRestartProv');
      if (btnRestart) btnRestart.style.display = 'inline-flex';
    } else {
      log('No sync jobs found for this app.');
      toast('No sync jobs found');
    }
  } catch (e) {
    log('Sync check error: ' + e.message);
    toast(e.message, 'error');
  }
}

async function restartProvisioning() {
  if (!S.selectedApp) return;
  try {
    const { jobs } = await GET(`/apps/${S.selectedApp.id}/sync-jobs`);
    if (!jobs?.length) { toast('No sync job found', 'error'); return; }
    await POST(`/apps/${S.selectedApp.id}/sync-jobs/${jobs[0].id}/restart`);
    log('Provisioning restart requested.');
    toast('Provisioning restarted', 'success');
  } catch (e) {
    log('Restart error: ' + e.message);
    toast(e.message, 'error');
  }
}

async function confirmDeleteApp() {
  if (!S.selectedApp) return;
  if (!confirm(`Permanently delete "${S.selectedApp.displayName}"?\n\nThis removes the service principal from your tenant and cannot be undone.`)) return;
  log('Deleting: ' + S.selectedApp.displayName);
  try {
    await DEL(`/apps/${S.selectedApp.id}`);
    toast('App deleted', 'success');
    log('Deleted: ' + S.selectedApp.displayName);
    S.selectedApp = null;
    S.selectedAppRaw = null;
    document.getElementById('appDetail').style.display = 'none';
    document.getElementById('appDetailEmpty').style.display = S.isAuthenticated ? 'flex' : 'none';
    document.querySelectorAll('.app-item.active').forEach(el => el.remove());
  } catch (e) {
    log('Delete error: ' + e.message);
    toast(e.message, 'error');
  }
}

async function exportUsersCsv() {
  if (!S.selectedApp) { toast('Select an app first', 'error'); return; }
  log('Exporting assignments CSV...');
  await downloadApi('GET', `/apps/${S.selectedApp.id}/export-csv`, null, `${S.selectedApp.displayName}_assignments.csv`);
}

async function removeChecked() {
  if (!S.selectedApp) return;
  const checked = [...document.querySelectorAll('#assignTbody input[type=checkbox]:checked')];
  if (!checked.length) { toast('Check assignments to remove', 'error'); return; }

  const q = (document.getElementById('assignFilter').value || '').toLowerCase();
  const filtered = q
    ? S.assignments.filter(a => (a.principalDisplayName||'').toLowerCase().includes(q) || (a.emailOrId||'').toLowerCase().includes(q))
    : S.assignments;
  const toRemove = checked.map(cb => filtered[parseInt(cb.dataset.idx)]).filter(Boolean);

  if (!confirm(`Remove ${toRemove.length} assignment(s)?`)) return;
  log(`Removing ${toRemove.length} assignment(s)...`);
  let ok = 0, fail = 0;
  const removedIds = new Set();
  for (const a of toRemove) {
    try {
      await DEL(`/apps/${S.selectedApp.id}/assignments/${a.id}`);
      log('Removed: ' + (a.principalDisplayName || a.principalId));
      removedIds.add(a.id);
      ok++;
    } catch (e) {
      log('Remove error: ' + e.message);
      fail++;
    }
  }

  if (removedIds.size) {
    S.assignments = S.assignments.filter(a => !removedIds.has(a.id));
    renderAssignments();
    log(`Removed locally from UI: ${removedIds.size} assignment(s).`);
  }

  toast(`Removed ${ok}${fail ? ` (${fail} failed)` : ''}`, ok ? 'success' : 'error');
  // Reload assignments immediately and once more after a short delay to absorb Graph propagation lag.
  const reload = async () => {
    try {
      const data = await getAssignments(S.selectedApp.id);
      S.assignments = data.assignments || [];
      renderAssignments();
      log(`Assignments reloaded: ${S.assignments.length}`);
    } catch (_) {}
  };
  await reload();
  setTimeout(() => { reload(); }, 1200);
}

async function viewPrincipal(id, type, name) {
  if (type === 'Group') {
    document.getElementById('groupMembersTitle').textContent = `Members of "${name}"`;
    document.getElementById('groupMembersTbody').innerHTML = '<tr><td colspan="2" class="tbl-empty"><span class="spinner spinner-dark"></span></td></tr>';
    openModal('groupMembersModal');
    try {
      const data = await GET(`/apps/${S.selectedApp.id}/group-members/${id}`);
      const members = data.members || [];
      document.getElementById('groupMembersTbody').innerHTML = members.length
        ? members.map(m => `<tr><td>${escHtml(m.displayName||'')}</td><td class="mono" style="font-size:11.5px">${escHtml(m.userPrincipalName||'')}</td></tr>`).join('')
        : '<tr><td colspan="2" class="tbl-empty">No members found</td></tr>';
    } catch (e) {
      document.getElementById('groupMembersTbody').innerHTML = `<tr><td colspan="2" class="tbl-empty">${escHtml(e.message)}</td></tr>`;
    }
  } else {
    document.getElementById('userAppsTitle').textContent = `Litera App Assignments for "${name}"`;
    document.getElementById('userAppsTbody').innerHTML = '<tr><td class="tbl-empty"><span class="spinner spinner-dark"></span></td></tr>';
    openModal('userAppsModal');
    try {
      const data = await GET(`/apps/user-assignments/${id}`);
      const litera = (data.assignments || []).filter(a => S.allApps.some(app => app.id === a.resourceId));
      document.getElementById('userAppsTbody').innerHTML = litera.length
        ? litera.map(a => `<tr><td>${escHtml(a.resourceDisplayName||'')}</td></tr>`).join('')
        : '<tr><td class="tbl-empty">No Litera app assignments</td></tr>';
    } catch (e) {
      document.getElementById('userAppsTbody').innerHTML = `<tr><td class="tbl-empty">${escHtml(e.message)}</td></tr>`;
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  ASSIGN MODAL (for existing apps)
// ═══════════════════════════════════════════════════════════════════════════════

function toggleCopyAssignRow(context) {
  const map = {
    assign: { checkId: 'chkAssignCopyFromApp', rowId: 'assignCopyRow' },
    deploy: { checkId: 'chkDeployAssignCopyFromApp', rowId: 'deployAssignCopyRow' },
    scim: { checkId: 'chkScimAssignCopyFromApp', rowId: 'scimAssignCopyRow' },
  };
  const cfg = map[context];
  if (!cfg) return;
  const checked = document.getElementById(cfg.checkId)?.checked;
  const row = document.getElementById(cfg.rowId);
  if (row) row.style.display = checked ? 'flex' : 'none';
}

async function copyAssignmentsIntoSelection(context) {
  const map = {
    assign: { selectId: 'assignCopySourceApp', selected: S.assignSelected, countId: 'selectedCount', chipsId: 'selectedChips' },
    deploy: { selectId: 'deployAssignCopySourceApp', selected: S.deployAssignSelected, countId: 'deploySelectedCount', chipsId: 'deploySelectedChips' },
    scim: { selectId: 'scimAssignCopySourceApp', selected: S.scimAssignSelected, countId: null, chipsId: 'scimAssignChips' },
  };
  const cfg = map[context];
  if (!cfg) return;
  const selectedValue = document.getElementById(cfg.selectId)?.value;
  const sourceAppId = selectedValue || S.selectedApp?.id || '';
  if (!sourceAppId) { toast('Select a source app', 'error'); return; }
  log(`Copy assignments requested (${context}) from app: ${sourceAppId}`);

  try {
    let data = await getAssignments(sourceAppId);
    let fromApp = data.assignments || [];
    if (!fromApp.length) {
      data = await getAssignmentsWithRetry(sourceAppId, { attempts: 6, delayMs: 900, minCount: 1 });
      fromApp = data.assignments || [];
    }
    if (!fromApp.length && Array.isArray(S.appAssignmentsCache[sourceAppId])) {
      fromApp = S.appAssignmentsCache[sourceAppId];
      log(`Source app assignments loaded from cache: ${fromApp.length}`);
    }
    if (!fromApp.length && S.selectedApp?.id === sourceAppId && Array.isArray(S.assignments) && S.assignments.length) {
      fromApp = S.assignments;
      log(`Source app assignments loaded from selected app state: ${fromApp.length}`);
    }
    log(`Source app assignments fetched: ${fromApp.length}`);
    let added = 0;
    let eligible = 0;
    for (const a of fromApp) {
      if (!a?.principalId) continue;
      const rawType = String(a.principalType || '').toLowerCase();
      if (rawType === 'serviceprincipal') continue;
      eligible++;
      const normalizedType = rawType === 'group' ? 'Group' : 'User';
      if (cfg.selected.some(s => s.id === a.principalId)) continue;
      cfg.selected.push({
        id: a.principalId,
        name: a.principalDisplayName || a.principalId,
        email: a.emailOrId || a.principalId,
        emailOrId: a.emailOrId || a.principalId,
        type: normalizedType,
      });
      added++;
    }
    if (added === 0 && eligible > 0) {
      toast(`Nothing new to copy (${eligible} already selected)`, 'error');
      if (cfg.countId) document.getElementById(cfg.countId).textContent = cfg.selected.length;
      _renderChips(cfg.chipsId, cfg.selected, context);
      return;
    }
    if (!added) {
      toast('No assignable user/group entries found in source app', 'error');
      log('No assignable user/group entries found to copy.');
      return;
    }
    if (cfg.countId) document.getElementById(cfg.countId).textContent = cfg.selected.length;
    _renderChips(cfg.chipsId, cfg.selected, context);
    toast(`Copied ${added} assignment(s)`, 'success');
    log(`Copied ${added} assignment(s) into selection.`);
  } catch (e) {
    log(`Copy assignments error: ${e.message}`);
    toast(e.message, 'error');
  }
}

function openAssignModal() {
  if (!S.selectedApp) { toast('Select an app first', 'error'); return; }
  S.assignSelected = [];
  S.assignPaging = { nextLink: null, items: [] };
  document.getElementById('assignSearchInput').value = '';
  document.getElementById('assignResultList').innerHTML = '<div class="search-hint">Type to search...</div>';
  document.getElementById('selectedChips').innerHTML = '';
  document.getElementById('selectedCount').textContent = '0';
  const assignSelectAll = document.getElementById('assignSelectAllResults');
  if (assignSelectAll) { assignSelectAll.checked = false; assignSelectAll.indeterminate = false; }
  const assignCopyCheck = document.getElementById('chkAssignCopyFromApp');
  if (assignCopyCheck) assignCopyCheck.checked = false;
  toggleCopyAssignRow('assign');
  populateCopySourceDrop();
  setAssignTab('users', document.querySelector('#assignModal .tab-btn'));
  openModal('assignModal');
}

function setAssignTab(type, btn) {
  S.assignTab = type;
  document.querySelectorAll('#assignModal .tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  searchAssignables();
}

function searchAssignables() {
  clearTimeout(S._assignTimer);
  S._assignTimer = setTimeout(() => _fetchPrincipals('assign', false), 280);
}

function _prefixState(prefix) {
  if (prefix === 'assign') {
    return {
      selected: S.assignSelected,
      paging: S.assignPaging,
      countId: 'selectedCount',
      chipsId: 'selectedChips',
      selectAllId: 'assignSelectAllResults',
    };
  }
  if (prefix === 'deploy') {
    return {
      selected: S.deployAssignSelected,
      paging: S.deployAssignPaging,
      countId: 'deploySelectedCount',
      chipsId: 'deploySelectedChips',
      selectAllId: 'deploySelectAllResults',
    };
  }
  if (prefix === 'scim') {
    return {
      selected: S.scimAssignSelected,
      paging: S.scimAssignPaging,
      countId: null,
      chipsId: 'scimAssignChips',
      selectAllId: 'scimSelectAllResults',
    };
  }
  if (prefix === 'lum') {
    return {
      selected: S.lumSelected,
      paging: S.lumPaging,
      countId: 'lumSelectedCount',
      chipsId: 'lumSelectedChips',
      selectAllId: 'lumSelectAllResults',
    };
  }
  return null;
}

function _refreshSelectAllResultsUi(prefix) {
  const cfg = _prefixState(prefix);
  if (!cfg) return;
  const checkbox = document.getElementById(cfg.selectAllId);
  if (!checkbox) return;
  const items = cfg.paging.items || [];
  if (!items.length) {
    checkbox.checked = false;
    checkbox.indeterminate = false;
    return;
  }
  const selectedIds = new Set(cfg.selected.map(s => s.id));
  const selectedCount = items.filter(i => selectedIds.has(i.id)).length;
  checkbox.checked = selectedCount > 0 && selectedCount === items.length;
  checkbox.indeterminate = selectedCount > 0 && selectedCount < items.length;
}

function toggleSelectAllResults(prefix, masterCb) {
  const cfg = _prefixState(prefix);
  if (!cfg) return;
  const items = cfg.paging.items || [];
  if (!items.length) return;
  const selectedIds = new Set(cfg.selected.map(s => s.id));
  for (const item of items) {
    const sub = item.userPrincipalName || item.mail || item.id;
    const type = item['@odata.type']?.includes('group') ? 'Group' : 'User';
    if (masterCb.checked) {
      if (!selectedIds.has(item.id)) {
        cfg.selected.push({ id: item.id, name: item.displayName, email: sub, type });
        selectedIds.add(item.id);
      }
    } else {
      const idx = cfg.selected.findIndex(s => s.id === item.id);
      if (idx >= 0) cfg.selected.splice(idx, 1);
    }
  }
  if (cfg.countId) document.getElementById(cfg.countId).textContent = cfg.selected.length;
  if (cfg.chipsId) _renderChips(cfg.chipsId, cfg.selected, prefix);
  if (prefix === 'lum') updateUsersManagementSummary();
  _renderSearchResults(document.getElementById({
    assign: 'assignResultList',
    deploy: 'deployAssignResultList',
    scim: 'scimAssignResultList',
    lum: 'lumResultList',
  }[prefix]), items, cfg.selected, prefix, !!cfg.paging.nextLink);
}

function _renderSearchResults(el, items, selectedArr, prefix, hasMore = false) {
  if (!items.length) {
    el.innerHTML = '<div class="search-hint">No results found</div>';
    _refreshSelectAllResultsUi(prefix);
    return;
  }
  const list = items.map(item => {
    const sub = item.userPrincipalName || item.mail || item.id;
    const sel = selectedArr.some(s => s.id === item.id);
    return `<div class="search-item" onclick="_toggleSelect('${prefix}','${escHtml(item.id)}','${escHtml(item.displayName||'')}','${escHtml(sub||'')}','${item['@odata.type']?.includes('group') ? 'Group' : 'User'}')">
      <div>
        <div class="search-item-name">${escHtml(item.displayName||'')}</div>
        <div class="search-item-sub">${escHtml(sub||'')}</div>
      </div>
      <span class="search-item-action ${sel ? 'added' : ''}" id="${prefix}-sri-${escHtml(item.id)}">${sel ? '✓ Added' : '+ Add'}</span>
    </div>`;
  }).join('');
  const more = hasMore
    ? `<div class="search-hint"><button class="btn btn-secondary btn-sm" onclick="loadMorePrincipals('${prefix}')">Load More</button></div>`
    : '';
  el.innerHTML = list + more;
  _refreshSelectAllResultsUi(prefix);
}

function loadMorePrincipals(prefix) {
  _fetchPrincipals(prefix, true);
}

async function _fetchPrincipals(prefix, append) {
  const map = {
    assign: { inputId: 'assignSearchInput', resultId: 'assignResultList', tab: () => S.assignTab, selected: () => S.assignSelected, paging: 'assignPaging' },
    deploy: { inputId: 'deployAssignInput', resultId: 'deployAssignResultList', tab: () => S.deployAssignTab, selected: () => S.deployAssignSelected, paging: 'deployAssignPaging' },
    scim: { inputId: 'scimAssignInput', resultId: 'scimAssignResultList', tab: () => S.scimAssignTab, selected: () => S.scimAssignSelected, paging: 'scimAssignPaging' },
    lum: { inputId: 'lumSearchInput', resultId: 'lumResultList', tab: () => S.lumTab, selected: () => S.lumSelected, paging: 'lumPaging' },
  };
  const cfg = map[prefix];
  if (!cfg) return;

  const q = document.getElementById(cfg.inputId).value.trim();
  const el = document.getElementById(cfg.resultId);
  const paging = S[cfg.paging];

  if (!append) {
    paging.nextLink = null;
    paging.items = [];
    el.innerHTML = '<div class="search-hint"><span class="spinner spinner-dark"></span></div>';
  }

  if (prefix === 'lum' && S.usersManagementMode === 'remove') {
    const selectedApps = getUsersManagementSelectedApps();
    if (!selectedApps.length) {
      paging.items = [];
      paging.nextLink = null;
      el.innerHTML = '<div class="search-hint">Select at least one app to load current assignments</div>';
      _refreshSelectAllResultsUi(prefix);
      return;
    }

    try {
      const principalMap = new Map();
      for (const app of selectedApps) {
        const data = await getAssignments(app.id);
        const rows = data.assignments || [];
        for (const a of rows) {
          if (!a?.principalId) continue;
          const type = a.principalType === 'Group' ? 'Group' : 'User';
          if (principalMap.has(a.principalId)) continue;
          principalMap.set(a.principalId, {
            id: a.principalId,
            displayName: a.principalDisplayName || a.emailOrId || a.principalId,
            userPrincipalName: type === 'User' ? (a.emailOrId || '') : '',
            mail: type === 'Group' ? (a.emailOrId || '') : '',
            '@odata.type': type === 'Group' ? '#microsoft.graph.group' : '#microsoft.graph.user',
          });
        }
      }

      const modeType = cfg.tab() === 'groups' ? '#microsoft.graph.group' : '#microsoft.graph.user';
      const query = q.toLowerCase();
      const filtered = [...principalMap.values()]
        .filter(p => p['@odata.type'] === modeType)
        .filter(p => {
          if (!query) return true;
          const sub = p.userPrincipalName || p.mail || p.id || '';
          return (p.displayName || '').toLowerCase().includes(query) || sub.toLowerCase().includes(query);
        })
        .sort((a, b) => String(a.displayName || '').localeCompare(String(b.displayName || '')));

      paging.items = filtered;
      paging.nextLink = null;
      _renderSearchResults(el, paging.items, cfg.selected(), prefix, false);
    } catch (e) {
      el.innerHTML = `<div class="search-hint" style="color:var(--danger)">${escHtml(e.message)}</div>`;
    }
    return;
  }

  let url = `/principals/search?type=${cfg.tab()}&top=20`;
  if (append && paging.nextLink) {
    url += `&nextLink=${encodeURIComponent(paging.nextLink)}`;
  } else if (q) {
    url += `&q=${encodeURIComponent(q)}`;
  }

  try {
    const data = await GET(url);
    const pageItems = data.principals || [];
    paging.items = append ? [...paging.items, ...pageItems] : pageItems;
    paging.nextLink = data.nextLink || null;
    _renderSearchResults(el, paging.items, cfg.selected(), prefix, !!paging.nextLink);
  } catch (e) {
    el.innerHTML = `<div class="search-hint" style="color:var(--danger)">${escHtml(e.message)}</div>`;
  }
}

function _toggleSelect(prefix, id, name, email, type) {
  const cfg = _prefixState(prefix);
  if (!cfg) return;
  const arr = cfg.selected;

  const idx = arr.findIndex(s => s.id === id);
  if (idx >= 0) arr.splice(idx, 1); else arr.push({ id, name, email, type });

  const isSel = arr.some(s => s.id === id);
  const span = document.getElementById(`${prefix}-sri-${id}`);
  if (span) { span.textContent = isSel ? '✓ Added' : '+ Add'; span.className = 'search-item-action' + (isSel ? ' added' : ''); }
  if (cfg.countId) document.getElementById(cfg.countId).textContent = arr.length;
  if (cfg.chipsId) _renderChips(cfg.chipsId, arr, prefix);
  if (prefix === 'lum') updateUsersManagementSummary();
  _refreshSelectAllResultsUi(prefix);
}

function _renderChips(elId, arr, prefix) {
  if (!arr.length) {
    document.getElementById(elId).innerHTML = '<div class="search-hint" style="padding:0;color:var(--t3);text-align:left">No users or groups selected</div>';
    return;
  }
  document.getElementById(elId).innerHTML = arr.map(s =>
    `<span class="chip">
      ${escHtml(s.name)} <span class="badge ${s.type==='Group'?'badge-purple':'badge-blue'}" style="padding:1px 6px;font-size:10px">${s.type}</span>
      <button onclick="_toggleSelect('${prefix}','${s.id}','${escHtml(s.name)}','${escHtml(s.email)}','${s.type}')" title="Remove">×</button>
    </span>`).join('');
}

async function confirmAssign() {
  if (!S.assignSelected.length) { toast('Nothing selected', 'error'); return; }
  if (!S.selectedApp) return;
  const initialCount = S.assignments.length;
  closeModal('assignModal');
  log(`Assigning ${S.assignSelected.length} principal(s) to ${S.selectedApp.displayName}...`);
  let ok = 0, fail = 0;
  for (const p of S.assignSelected) {
    try {
      await POST(`/apps/${S.selectedApp.id}/assignments`, { principalId: p.id, principalType: p.type });
      log('Assigned: ' + p.name); ok++;
    } catch (e) {
      log(`Assign error (${p.name}): ${e.message}`); fail++;
    }
  }
  toast(`Assigned ${ok}${fail ? ` (${fail} failed)` : ''}`, ok ? 'success' : 'error');
  await refreshSelectedAssignments({ minCount: initialCount + ok });
  setTimeout(() => { refreshSelectedAssignments({ logLabel: 'Assignments auto-refreshed' }); }, 1200);
}

// ═══════════════════════════════════════════════════════════════════════════════
//  SCIM CREATION WIZARD
// ═══════════════════════════════════════════════════════════════════════════════

function openScimModal() {
  if (!S.isAuthenticated) {
    toast('Please sign in to create a SCIM app.', 'error');
    return;
  }
  S.scimStep = 1;
  S.scimAssignSelected = [];
  S.scimAssignPaging = { nextLink: null, items: [] };
  _updateScimSteps();
  ['scimPane1','scimPane2','scimPane3','scimPane4'].forEach((id, i) => {
    document.getElementById(id).style.display = i === 0 ? 'block' : 'none';
  });
  document.getElementById('scimBackBtn').style.display = 'none';
  document.getElementById('scimNextBtn').style.display = 'inline-flex';
  document.getElementById('scimNextBtn').textContent = 'Next →';
  document.getElementById('scimCancelBtn').textContent = 'Cancel';
  document.getElementById('scimAssignInput').value = '';
  document.getElementById('scimAssignResultList').innerHTML = '<div class="search-hint">Type to search...</div>';
  document.getElementById('scimAssignChips').innerHTML = '';
  const scimSelectAll = document.getElementById('scimSelectAllResults');
  if (scimSelectAll) { scimSelectAll.checked = false; scimSelectAll.indeterminate = false; }
  const scimCopyCheck = document.getElementById('chkScimAssignCopyFromApp');
  if (scimCopyCheck) scimCopyCheck.checked = false;
  toggleCopyAssignRow('scim');
  populateCopySourceDrop();
  document.getElementById('scimTestResult').textContent = '';
  document.getElementById('scimCreationLog').textContent = 'Waiting to start...';
  document.getElementById('scimCreateProgress').style.width = '0%';
  document.getElementById('scimResultBox').style.display = 'none';
  openModal('scimModal');
}

function _updateScimSteps() {
  [1, 2, 3, 4].forEach(n => {
    const el = document.getElementById(`scimStep${n}Ind`);
    el.className = 'step' + (n < S.scimStep ? ' done' : n === S.scimStep ? ' active' : '');
  });
}

function scimCancel() { closeModal('scimModal'); }
function scimBack() {
  if (S.scimStep === 2) { S.scimStep = 1; }
  else if (S.scimStep === 3) { S.scimStep = 2; }
  _showScimPane();
}
function scimNext() {
  if (S.scimStep === 1) {
    if (!document.getElementById('scimName').value.trim()) { toast('App name is required', 'error'); return; }
    S.scimStep = 2;
  } else if (S.scimStep === 2) {
    if (!document.getElementById('scimUrl').value.trim()) { toast('SCIM tenant URL is required', 'error'); return; }
    if (!document.getElementById('scimSecretToken').value.trim()) { toast('SCIM secret token is required', 'error'); return; }
    S.scimStep = 3;
  } else if (S.scimStep === 3) {
    S.scimStep = 4;
    _showScimPane();
    _startScimCreation();
    return;
  }
  _showScimPane();
}
function _showScimPane() {
  [1, 2, 3, 4].forEach(n => {
    document.getElementById(`scimPane${n}`).style.display = n === S.scimStep ? 'block' : 'none';
  });
  document.getElementById('scimBackBtn').style.display = S.scimStep > 1 && S.scimStep < 4 ? 'inline-flex' : 'none';
  document.getElementById('scimNextBtn').style.display = S.scimStep < 4 ? 'inline-flex' : 'none';
  document.getElementById('scimNextBtn').textContent = S.scimStep === 3 ? 'Create App →' : 'Next →';
  _updateScimSteps();
}

function setScimAssignTab(type, btn) {
  S.scimAssignTab = type;
  document.querySelectorAll('#scimPane3 .tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  searchScimAssignables();
}

function searchScimAssignables() {
  clearTimeout(S._scimAssignTimer);
  S._scimAssignTimer = setTimeout(() => _fetchPrincipals('scim', false), 280);
}

async function _startScimCreation() {
  const appName    = document.getElementById('scimName').value.trim();
  const scimUrl    = document.getElementById('scimUrl').value.trim();
  const scimToken  = document.getElementById('scimSecretToken').value.trim();
  const principals = S.scimAssignSelected;
  const probeSpId  = S.allApps[0]?.id || null;
  const csrfToken  = window.__CSRF_TOKEN__ || '';

  const logEl  = document.getElementById('scimCreationLog');
  const progEl = document.getElementById('scimCreateProgress');
  logEl.textContent = '';

  function scimLog(msg) { logEl.textContent += new Date().toLocaleTimeString() + ' — ' + msg + '\n'; logEl.scrollTop = logEl.scrollHeight; log('SCIM: ' + msg); }
  function scimProg(pct) { progEl.style.width = pct + '%'; }

  try {
    const res = await fetch('/api/scim/create', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...(csrfToken ? { 'x-csrf-token': csrfToken } : {}),
      },
      credentials: 'same-origin',
      body: JSON.stringify({ appName, scimUrl, scimToken, principals, probeSpId }),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: 'Request failed' }));
      throw new Error(err.error || 'Creation failed');
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      const lines = buf.split('\n');
      buf = lines.pop();
      for (const line of lines) {
        if (!line.startsWith('data:')) continue;
        try {
          const evt = JSON.parse(line.slice(5).trim());
          if (evt.type === 'progress') { scimLog(evt.message); scimProg(evt.percent || 0); }
          else if (evt.type === 'diag')   { scimLog('  → ' + evt.message); }
          else if (evt.type === 'result') {
            if (evt.success) {
              scimProg(100);
              scimLog(`✅ App "${evt.appName}" created successfully!`);
              scimLog(`   SP ID: ${evt.spId} | Assigned: ${evt.assignedCount}/${evt.assignedRequested}`);
              _showScimResult(true, evt);
              toast('SCIM app created: ' + evt.appName, 'success');
              loadApps(); // refresh app list
            } else {
              scimLog('❌ Creation failed: ' + (evt.error || 'Unknown error'));
              _showScimResult(false, evt);
            }
          }
        } catch (_) {}
      }
    }
  } catch (e) {
    scimLog('❌ Error: ' + e.message);
    _showScimResult(false, { error: e.message });
    toast(e.message, 'error');
  }
  document.getElementById('scimCancelBtn').textContent = 'Close';
}

function _showScimResult(success, data) {
  const box = document.getElementById('scimResultBox');
  box.style.display = 'block';
  if (success) {
    box.innerHTML = `<div class="info-box" style="background:#f0fdf4;border-color:#86efac;color:#15803d">
      <strong>✅ SCIM app provisioned successfully!</strong><br/>
      App Name: ${escHtml(data.appName)} | SP ID: <span class="mono">${escHtml(data.spId)}</span><br/>
      Assigned: ${data.assignedCount || 0} of ${data.assignedRequested || 0} principals
    </div>`;
  } else {
    box.innerHTML = `<div class="alert alert-error">❌ ${escHtml(data.error || 'Creation failed')}</div>`;
  }
}

function toggleScimToken() {
  const el = document.getElementById('scimSecretToken');
  el.type = el.type === 'password' ? 'text' : 'password';
}

async function testScimConnection() {
  const scimUrl = document.getElementById('scimUrl').value.trim();
  const scimToken = document.getElementById('scimSecretToken').value.trim();
  const resultEl = document.getElementById('scimTestResult');
  resultEl.textContent = 'Testing connection...';
  resultEl.style.color = 'var(--t3)';
  try {
    const data = await POST('/scim/test-connection', { scimUrl, scimToken });
    resultEl.textContent = `Connection successful (${data.status}) via ${data.url}`;
    resultEl.style.color = 'var(--accent3)';
    toast('SCIM connection successful', 'success');
  } catch (e) {
    resultEl.textContent = e.message || 'SCIM connection test failed';
    resultEl.style.color = 'var(--danger)';
    toast('SCIM connection failed', 'error');
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  DEPLOYMENTS
// ═══════════════════════════════════════════════════════════════════════════════

function updateDeployProduct() {
  const product = document.querySelector('[name=deployProduct]:checked')?.value || 'Outlook';
  const sourceInput = document.getElementById('sourceUrl');
  if (sourceInput) sourceInput.value = DEPLOY_DEFAULT_URLS[product] || DEPLOY_DEFAULT_URLS.Outlook;
}

function updateSourceUi() {
  const src = document.querySelector('[name=deploySource]:checked')?.value;
  const manifestRow = document.getElementById('manifestRow');
  const sourceUrlRow = document.getElementById('sourceUrlRow');
  const appSourceOption = document.getElementById('deploySourceAppSourceOption');
  if (manifestRow) manifestRow.style.display = src === 'manifest' ? 'block' : 'none';
  if (sourceUrlRow) sourceUrlRow.style.display = src === 'manifest' ? 'none' : 'block';
  if (appSourceOption) appSourceOption.style.opacity = src === 'manifest' ? '0.5' : '1';
}

async function onManifestFilePicked(event) {
  const file = event?.target?.files?.[0];
  const fileLabel = document.getElementById('manifestFileName');
  S.deployManifestFileContent = '';
  S.deployManifestFileName = '';
  if (!file) {
    if (fileLabel) fileLabel.textContent = '';
    return;
  }
  try {
    const txt = await file.text();
    S.deployManifestFileContent = txt;
    S.deployManifestFileName = file.name;
    if (fileLabel) fileLabel.textContent = `Selected file: ${file.name}`;
    const manifestUrlInput = document.getElementById('manifestUrl');
    if (manifestUrlInput && !manifestUrlInput.value.trim()) manifestUrlInput.value = file.name;
  } catch (e) {
    if (fileLabel) fileLabel.textContent = `Failed to read file: ${e.message}`;
    toast('Unable to read manifest file', 'error');
  }
}

function updateMethodUi() {
  const m = document.querySelector('[name=deployMethod]:checked')?.value;
  document.getElementById('integratedPanel').style.display = m === 'integrated' ? 'block' : 'none';
  const targetsSection = document.getElementById('deployTargetsSection');
  if (targetsSection) targetsSection.style.display = m === 'integrated' ? 'none' : 'block';
  const btn = document.getElementById('btnDeploy');
  if (btn) btn.textContent = m === 'integrated' ? '☁ Log Deployment Request' : '☁ Deploy Add-in';
}

function populateCopySourceDrop() {
  const byId = new Map();
  [...(S.allApps || []), ...(S.lumApps || [])].forEach(app => {
    if (app?.id && !byId.has(app.id)) byId.set(app.id, app);
  });
  const sourceApps = [...byId.values()];
  const options = '<option value="">— Select app —</option>' +
    sourceApps.map(a => `<option value="${escHtml(a.id)}">${escHtml(a.displayName)}</option>`).join('');
  ['copySourceApp', 'assignCopySourceApp', 'deployAssignCopySourceApp', 'scimAssignCopySourceApp'].forEach(id => {
    const sel = document.getElementById(id);
    if (sel) sel.innerHTML = options;
  });
}

function toggleCopyTargetsRow() {
  const show = document.getElementById('chkCopyTargets').checked;
  document.getElementById('copyTargetsRow').style.display = show ? 'flex' : 'none';
}

async function copyTargetsFromApp() {
  const spId = document.getElementById('copySourceApp').value;
  if (!spId) { toast('Select a source app', 'error'); return; }
  log('Copying targets from app: ' + spId);
  try {
    let data = await getAssignments(spId);
    let fromApp = data.assignments || [];
    if (!fromApp.length) {
      data = await getAssignmentsWithRetry(spId, { attempts: 6, delayMs: 900, minCount: 1 });
      fromApp = data.assignments || [];
    }
    if (!fromApp.length && Array.isArray(S.appAssignmentsCache[spId])) {
      fromApp = S.appAssignmentsCache[spId];
      log(`Targets source assignments loaded from cache: ${fromApp.length}`);
    }
    const targets = fromApp
      .filter(a => {
        if (!a?.principalId) return false;
        const t = String(a.principalType || '').toLowerCase();
        return t === 'user' || t === 'group';
      })
      .map(a => ({
        id: a.principalId,
        name: a.principalDisplayName,
        emailOrId: a.emailOrId,
        type: String(a.principalType || '').toLowerCase() === 'group' ? 'Group' : 'User',
      }));
    if (!targets.length) {
      toast('No user/group assignments found in source app', 'error');
      return;
    }
    let added = 0;
    targets.forEach(t => { if (!S.deployTargets.some(d => d.id === t.id)) { S.deployTargets.push(t); added++; } });
    renderDeployTargets();
    toast(`Copied ${added} target(s)`, 'success');
    log(`Copied ${added} targets.`);
  } catch (e) {
    toast(e.message, 'error');
    log('Copy targets error: ' + e.message);
  }
}

function renderDeployTargets() {
  const tb = document.getElementById('deployTargetsTbody');
  const count = S.deployTargets.length;
  document.getElementById('deployTargetCount').textContent = `${count} selected`;
  document.getElementById('clearTargetsBtn').style.display = count ? 'inline-flex' : 'none';
  if (!count) { tb.innerHTML = '<tr><td colspan="4" class="tbl-empty">No targets selected</td></tr>'; return; }
  tb.innerHTML = S.deployTargets.map((t, i) => `
    <tr>
      <td class="chk-col"><input type="checkbox" data-di="${i}"></td>
      <td>${escHtml(t.name || '(unknown)')}</td>
      <td class="mono truncate" style="max-width:180px;font-size:11.5px">${escHtml(t.emailOrId || t.id || '')}</td>
      <td>${fmtType(t.type || 'User')}</td>
    </tr>`).join('');
}

function clearDeployTargets() { S.deployTargets = []; renderDeployTargets(); }

function openDeployAssignModal() {
  S.deployAssignSelected = [];
  S.deployAssignPaging = { nextLink: null, items: [] };
  document.getElementById('deployAssignInput').value = '';
  document.getElementById('deployAssignResultList').innerHTML = '<div class="search-hint">Type to search...</div>';
  document.getElementById('deploySelectedChips').innerHTML = '';
  document.getElementById('deploySelectedCount').textContent = '0';
  const deploySelectAll = document.getElementById('deploySelectAllResults');
  if (deploySelectAll) { deploySelectAll.checked = false; deploySelectAll.indeterminate = false; }
  const deployCopyCheck = document.getElementById('chkDeployAssignCopyFromApp');
  if (deployCopyCheck) deployCopyCheck.checked = false;
  toggleCopyAssignRow('deploy');
  populateCopySourceDrop();
  setDeployAssignTab('users', document.querySelector('#deployAssignModal .tab-btn'));
  openModal('deployAssignModal');
}

function setDeployAssignTab(type, btn) {
  S.deployAssignTab = type;
  document.querySelectorAll('#deployAssignModal .tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  searchDeployAssignables();
}

function searchDeployAssignables() {
  clearTimeout(S._deployAssignTimer);
  S._deployAssignTimer = setTimeout(() => _fetchPrincipals('deploy', false), 280);
}

function confirmDeployAssign() {
  S.deployAssignSelected.forEach(s => {
    if (!S.deployTargets.some(d => d.id === s.id)) S.deployTargets.push({ ...s, emailOrId: s.email });
  });
  renderDeployTargets();
  closeModal('deployAssignModal');
  toast(`Added ${S.deployAssignSelected.length} target(s)`, 'success');
}

function renderDeploymentStatusList() {
  const listEl = document.getElementById('deployStatusList');
  if (!listEl) return;
  if (!S.deployStatusItems.length) {
    listEl.innerHTML = '<div class="empty-state"><div class="em-icon">☁</div><div>Click Refresh to load deployment status</div></div>';
    return;
  }
  listEl.innerHTML = S.deployStatusItems.map((d, i) => `
    <div class="deploy-status-click" onclick="openDeployStatusDetail(${i})">
      <div style="font-weight:700;margin-bottom:4px">${escHtml(d.displayName || d.description || d.id || '(unnamed)')}</div>
      <div class="text-xs">${escHtml(d.state || d.status || d.lifecycleStatus || 'Status unavailable')}</div>
    </div>`).join('');
}

async function loadDeployments() {
  const listEl = document.getElementById('deployStatusList');
  listEl.innerHTML = '<div class="empty-state"><span class="spinner spinner-dark"></span></div>';
  log('Loading deployment status...');
  try {
    const data = await GET('/deploy/status');
    if (!data.supported || !data.deployments?.length) {
      listEl.innerHTML = `<div class="empty-state">
        <div class="em-icon">ℹ️</div>
        <div>${escHtml(data.message || 'No deployments found.')}</div>
        <div class="text-xs" style="margin-top:8px">This API endpoint may not be available for your tenant.</div>
      </div>`;
      return;
    }
    S.deployStatusItems = data.deployments || [];
    renderDeploymentStatusList();
    log(`Loaded ${S.deployStatusItems.length} deployment(s).`);
  } catch (e) {
    log('Deployments error: ' + e.message);
    listEl.innerHTML = `<div class="empty-state">
      <div class="em-icon">⚠️</div>
      <div>Deployment status unavailable</div>
      <div class="text-xs" style="margin-top:8px">${escHtml(e.message)}</div>
    </div>`;
  }
}

async function submitDeploy() {
  const src = document.querySelector('[name=deploySource]:checked')?.value || 'appsource';
  const sourceValue = src === 'manifest'
    ? (S.deployManifestFileContent || document.getElementById('manifestUrl').value.trim())
    : document.getElementById('sourceUrl').value.trim();
  if (!sourceValue) { toast(src === 'manifest' ? 'Manifest URL/path or uploaded file is required' : 'Source URL is required', 'error'); return; }

  const product = document.querySelector('[name=deployProduct]:checked')?.value || 'Outlook';
  const method  = document.querySelector('[name=deployMethod]:checked')?.value || 'prepare';
  if (method !== 'integrated' && !S.deployTargets.length) { toast('Add deployment targets first', 'error'); return; }
  const selectedTargets = method === 'integrated' ? [] : S.deployTargets;

  log(`Submitting deployment: ${product} | ${src}${S.deployManifestFileName ? ` (${S.deployManifestFileName})` : ''} | ${selectedTargets.length} targets`);
  try {
    const data = await POST('/deploy/prepare', { product, source: src, sourceValue, targets: selectedTargets, method });
    toast(data.message, 'success');
    log('Deployment prepared: ' + data.message);
    const detailItem = {
      displayName: `Deployment Prepared — ${product}`,
      state: method === 'integrated' ? 'Pending in M365 Admin Center' : 'Prepared in portal',
      preparedAt: new Date().toISOString(),
      summary: data.summary,
      method,
      source: src,
      m365Url: data.m365Url,
    };
    S.deployStatusItems = [detailItem, ...S.deployStatusItems];
    renderDeploymentStatusList();
  } catch (e) {
    toast(e.message, 'error');
    log('Deploy error: ' + e.message);
  }
}

async function exportTargetsCsv() {
  if (!S.deployTargets.length) { toast('No targets to export', 'error'); return; }
  await downloadApi('POST', '/deploy/export-csv', { targets: S.deployTargets }, 'deploy_targets.csv');
}

function openDeployStatusDetail(index) {
  const item = S.deployStatusItems[index];
  if (!item) return;
  const detail = {
    name: item.displayName || item.description || item.id || '(unnamed)',
    status: item.state || item.status || item.lifecycleStatus || 'Unknown',
    id: item.id || null,
    sourceType: item.source || null,
    method: item.method || null,
    lastUpdated: new Date().toLocaleString(),
    raw: item,
  };
  const box = document.getElementById('deployStatusDetailBody');
  box.textContent = JSON.stringify(detail, null, 2);
  openModal('deployStatusModal');
}

// ═══════════════════════════════════════════════════════════════════════════════
//  LITERA ONE USERS MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════════════

function setUsersManagementMode(mode, btn) {
  S.usersManagementMode = mode === 'remove' ? 'remove' : 'assign';
  document.querySelectorAll('#page-users-management .users-mgmt-mode-toggle .users-mgmt-solid-btn')
    .forEach(b => b.classList.remove('active', 'green', 'blue', 'ghost'));
  const assignBtn = document.getElementById('lumModeAssignBtn');
  const removeBtn = document.getElementById('lumModeRemoveBtn');
  if (assignBtn) assignBtn.classList.add(S.usersManagementMode === 'assign' ? 'green' : 'ghost');
  if (removeBtn) removeBtn.classList.add(S.usersManagementMode === 'remove' ? 'blue' : 'ghost');
  if (btn) btn.classList.add('active');
  const actionBtn = document.getElementById('lumRunActionBtn');
  if (actionBtn) actionBtn.textContent = S.usersManagementMode === 'assign' ? 'Assign Selected' : 'Remove Selected';
  const summaryTitle = document.getElementById('lumSummaryTitle');
  if (summaryTitle) summaryTitle.textContent = S.usersManagementMode === 'assign' ? 'Ready to assign' : 'Ready to remove';
  const searchInput = document.getElementById('lumSearchInput');
  if (searchInput) {
    searchInput.placeholder = S.usersManagementMode === 'remove'
      ? 'Filter current assignments...'
      : 'Search users or groups...';
  }
  updateUsersManagementSummary();
  searchUsersManagementPrincipals();
}

function setUsersManagementPrincipalTab(type, btn) {
  S.lumTab = type === 'groups' ? 'groups' : 'users';
  document.querySelectorAll('#page-users-management .users-mgmt-principal-toggle .users-mgmt-solid-btn')
    .forEach(b => b.classList.remove('active', 'active-filter', 'ghost'));
  const usersBtn = document.getElementById('lumUsersTabBtn');
  const groupsBtn = document.getElementById('lumGroupsTabBtn');
  if (usersBtn) usersBtn.classList.add(S.lumTab === 'users' ? 'active-filter' : 'ghost');
  if (groupsBtn) groupsBtn.classList.add(S.lumTab === 'groups' ? 'active-filter' : 'ghost');
  if (btn) btn.classList.add('active');
  searchUsersManagementPrincipals();
}

function searchUsersManagementPrincipals() {
  clearTimeout(S._lumAssignTimer);
  S._lumAssignTimer = setTimeout(() => _fetchPrincipals('lum', false), 280);
}

function updateUsersManagementSummary() {
  const summary = document.getElementById('lumSummaryText');
  if (!summary) return;
  const selectedApps = getUsersManagementSelectedApps().length;
  const selectedPrincipals = S.lumSelected.length;
  const modeLabel = S.usersManagementMode === 'assign' ? 'Assign' : 'Remove';
  summary.textContent = `Mode: ${modeLabel} • Selected Principals: ${selectedPrincipals} • Selected Apps: ${selectedApps}`;
  const peopleStat = document.getElementById('lumStatPeople');
  const appsStat = document.getElementById('lumStatApps');
  const peopleCount = document.getElementById('lumSelectedPeopleCount');
  const appsCount = document.getElementById('lumSelectedAppsCount');
  if (peopleStat) peopleStat.textContent = String(selectedPrincipals);
  if (appsStat) appsStat.textContent = String(selectedApps);
  if (peopleCount) peopleCount.textContent = String(selectedPrincipals);
  if (appsCount) appsCount.textContent = String(selectedApps);
  renderUsersManagementSelectedApps();
}

async function loadUsersManagementApps() {
  if (!S.isAuthenticated) return;
  const tbody = document.getElementById('lumAppsTbody');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="4" class="tbl-empty"><span class="spinner spinner-dark"></span></td></tr>';
  try {
    const data = await GET('/apps/search?q=Litera');
    S.lumApps = data.apps || [];
    if (!S.allApps.length) S.allApps = S.lumApps.slice();
    populateCopySourceDrop();
    if (!S.lumApps.length) {
      tbody.innerHTML = '<tr><td colspan="4" class="tbl-empty">No Litera apps found</td></tr>';
      return;
    }
    renderUsersManagementApps();
  } catch (e) {
    tbody.innerHTML = `<tr><td colspan="4" class="tbl-empty">${escHtml(e.message)}</td></tr>`;
  }
}

function renderUsersManagementApps() {
  const tbody = document.getElementById('lumAppsTbody');
  if (!tbody) return;
  if (!S.lumApps.length) {
    tbody.innerHTML = '<tr><td colspan="4" class="tbl-empty">No Litera apps found</td></tr>';
    return;
  }
  tbody.innerHTML = S.lumApps.map((app, i) => `
    <tr>
      <td class="chk-col"><input type="checkbox" data-lum-app-idx="${i}" data-app-id="${escHtml(app.id || '')}" onchange="updateUsersManagementSummary(); if(S.usersManagementMode==='remove')searchUsersManagementPrincipals()"></td>
      <td>${escHtml(app.displayName || '')}</td>
      <td class="mono text-xs">${escHtml(app.id || '')}</td>
      <td>Not selected</td>
    </tr>`).join('');
  const allApps = document.getElementById('lumAppsSelectAll');
  if (allApps) allApps.checked = false;
  updateUsersManagementSummary();
}

function toggleUsersManagementAppsSelectAll(masterCb) {
  document.querySelectorAll('#lumAppsTbody input[type=checkbox][data-lum-app-idx]')
    .forEach(cb => { cb.checked = !!masterCb.checked; });
  updateUsersManagementSummary();
  if (S.usersManagementMode === 'remove') searchUsersManagementPrincipals();
}

function getUsersManagementSelectedApps() {
  return [...document.querySelectorAll('#lumAppsTbody input[type=checkbox][data-lum-app-idx]:checked')]
    .map(cb => S.lumApps[parseInt(cb.dataset.lumAppIdx, 10)])
    .filter(Boolean);
}

function renderUsersManagementHistory() {
  const tbody = document.getElementById('lumHistoryTbody');
  if (!tbody) return;
  const historyStat = document.getElementById('lumStatHistory');
  if (historyStat) historyStat.textContent = String(S.userManagementRecords.length);
  if (!S.userManagementRecords.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="tbl-empty">No history yet</td></tr>';
    return;
  }
  tbody.innerHTML = S.userManagementRecords.slice().reverse().map(r => `
    <tr>
      <td class="text-xs">${escHtml(r.time)}</td>
      <td>${escHtml(r.action)}</td>
      <td>${escHtml(r.appName)}</td>
      <td>
        <div>${escHtml(r.principalName)}</div>
        <div class="text-xs mono">${escHtml(r.emailOrId || '')}</div>
      </td>
      <td>${escHtml(r.principalType)}</td>
      <td><span class="badge ${r.status === 'Success' ? 'badge-green' : 'badge-red'}">${escHtml(r.status)}</span></td>
    </tr>
  `).join('');
}

function renderUsersManagementSelectedApps() {
  const el = document.getElementById('lumSelectedAppsChips');
  if (!el) return;
  const selectedApps = getUsersManagementSelectedApps();
  document.querySelectorAll('#lumAppsTbody tr').forEach(row => {
    const cb = row.querySelector('input[data-lum-app-idx]');
    if (!cb) return;
    row.classList.toggle('users-mgmt-row-selected', !!cb.checked);
    const statusCell = row.children[3];
    if (statusCell) statusCell.textContent = cb.checked ? 'Selected' : 'Not selected';
  });
  if (!selectedApps.length) {
    el.innerHTML = '<div class="search-hint" style="padding:0;color:var(--t3);text-align:left">No apps selected</div>';
    return;
  }
  el.innerHTML = selectedApps.map(app => `
    <span class="chip">
      ${escHtml(app.displayName || app.id)}
      <button onclick="toggleUsersManagementAppById('${escHtml(app.id)}')" title="Remove">×</button>
    </span>`).join('');
}

function toggleUsersManagementAppById(appId) {
  const checkbox = document.querySelector(`#lumAppsTbody input[data-lum-app-idx][data-app-id="${CSS.escape(appId)}"]`);
  if (checkbox) {
    checkbox.checked = false;
    updateUsersManagementSummary();
  }
}

function clearUsersManagementPrincipals() {
  S.lumSelected = [];
  document.getElementById('lumSelectedCount').textContent = '0';
  _renderChips('lumSelectedChips', S.lumSelected, 'lum');
  _refreshSelectAllResultsUi('lum');
  updateUsersManagementSummary();
}

function clearUsersManagementApps() {
  document.querySelectorAll('#lumAppsTbody input[type=checkbox][data-lum-app-idx]')
    .forEach(cb => { cb.checked = false; });
  const allApps = document.getElementById('lumAppsSelectAll');
  if (allApps) allApps.checked = false;
  updateUsersManagementSummary();
  if (S.usersManagementMode === 'remove') searchUsersManagementPrincipals();
}

function recordUsersManagementResult(entry) {
  S.userManagementRecords.push(entry);
  renderUsersManagementHistory();
}

async function runUsersManagementAction() {
  if (!S.isAuthenticated) { toast('Please sign in first', 'error'); return; }
  if (!S.lumSelected.length) { toast('Select at least one user/group', 'error'); return; }
  const selectedApps = getUsersManagementSelectedApps();
  if (!selectedApps.length) { toast('Select at least one Litera app', 'error'); return; }

  const action = S.usersManagementMode === 'remove' ? 'Remove' : 'Assign';
  log(`${action} users/groups across ${selectedApps.length} app(s)...`);
  let success = 0;
  let failed = 0;
  const touchesSelectedApp = !!S.selectedApp && selectedApps.some(a => a.id === S.selectedApp.id);

  for (const app of selectedApps) {
    if (S.usersManagementMode === 'assign') {
      for (const principal of S.lumSelected) {
        try {
          await POST(`/apps/${app.id}/assignments`, { principalId: principal.id, principalType: principal.type });
          success++;
          recordUsersManagementResult({
            time: new Date().toLocaleString(),
            action: 'Assigned',
            appName: app.displayName || app.id,
            appId: app.id,
            principalName: principal.name || principal.id,
            principalType: principal.type || 'User',
            principalId: principal.id,
            emailOrId: principal.email || principal.id,
            status: 'Success',
          });
        } catch (e) {
          failed++;
          recordUsersManagementResult({
            time: new Date().toLocaleString(),
            action: 'Assigned',
            appName: app.displayName || app.id,
            appId: app.id,
            principalName: principal.name || principal.id,
            principalType: principal.type || 'User',
            principalId: principal.id,
            emailOrId: principal.email || principal.id,
            status: 'Failed',
            error: e.message,
          });
          log(`Assign error [${app.displayName} - ${principal.name}]: ${e.message}`);
        }
      }
      continue;
    }

    let appAssignments = [];
    try {
      const data = await getAssignments(app.id);
      appAssignments = data.assignments || [];
    } catch (e) {
      for (const principal of S.lumSelected) {
        failed++;
        recordUsersManagementResult({
          time: new Date().toLocaleString(),
          action: 'Removed',
          appName: app.displayName || app.id,
          appId: app.id,
          principalName: principal.name || principal.id,
          principalType: principal.type || 'User',
          principalId: principal.id,
          emailOrId: principal.email || principal.id,
          status: 'Failed',
          error: `Unable to load assignments: ${e.message}`,
        });
      }
      continue;
    }

    for (const principal of S.lumSelected) {
      const existing = appAssignments.filter(a => a.principalId === principal.id);
      if (!existing.length) {
        failed++;
        recordUsersManagementResult({
          time: new Date().toLocaleString(),
          action: 'Removed',
          appName: app.displayName || app.id,
          appId: app.id,
          principalName: principal.name || principal.id,
          principalType: principal.type || 'User',
          principalId: principal.id,
          emailOrId: principal.email || principal.id,
          status: 'Failed',
          error: 'Assignment not found',
        });
        continue;
      }

      for (const assignment of existing) {
        try {
          await DEL(`/apps/${app.id}/assignments/${assignment.id}`);
          success++;
          recordUsersManagementResult({
            time: new Date().toLocaleString(),
            action: 'Removed',
            appName: app.displayName || app.id,
            appId: app.id,
            principalName: principal.name || principal.id,
            principalType: principal.type || 'User',
            principalId: principal.id,
            emailOrId: principal.email || principal.id,
            status: 'Success',
          });
        } catch (e) {
          failed++;
          recordUsersManagementResult({
            time: new Date().toLocaleString(),
            action: 'Removed',
            appName: app.displayName || app.id,
            appId: app.id,
            principalName: principal.name || principal.id,
            principalType: principal.type || 'User',
            principalId: principal.id,
            emailOrId: principal.email || principal.id,
            status: 'Failed',
            error: e.message,
          });
          log(`Remove error [${app.displayName} - ${principal.name}]: ${e.message}`);
        }
      }
    }
  }

  toast(`${action} completed: ${success} success${failed ? `, ${failed} failed` : ''}`, success ? 'success' : 'error');
  if (touchesSelectedApp) {
    await refreshSelectedAssignments({ logLabel: 'Selected app assignments auto-refreshed' });
    setTimeout(() => { refreshSelectedAssignments({ logLabel: 'Selected app assignments rechecked' }); }, 1200);
  }
  await loadUsersManagementApps();
  await _fetchPrincipals('lum', false);
  updateUsersManagementSummary();
  log('Users Management view refreshed after action.');
}

function exportUsersManagementRecords() {
  if (!S.userManagementRecords.length) {
    toast('No history to export', 'error');
    return;
  }
  const rows = [
    ['Time', 'Action', 'App Name', 'App ID', 'Principal Name', 'Principal Type', 'Principal ID', 'Email / ID', 'Status', 'Error'],
    ...S.userManagementRecords.map(r => [
      r.time || '',
      r.action || '',
      r.appName || '',
      r.appId || '',
      r.principalName || '',
      r.principalType || '',
      r.principalId || '',
      r.emailOrId || '',
      r.status || '',
      r.error || '',
    ]),
  ];
  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'litera_one_users_management_history.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

// ═══════════════════════════════════════════════════════════════════════════════
//  ENABLEMENT HUB
// ═══════════════════════════════════════════════════════════════════════════════

async function loadEnablement() {
  const q      = document.getElementById('enSearch').value.trim();
  const type   = S.enFilter;
  const el     = document.getElementById('enContent');
  el.innerHTML = '<div class="empty-state"><span class="spinner spinner-dark"></span></div>';
  try {
    const params = new URLSearchParams();
    if (q) params.set('q', q);
    if (type !== 'All') params.set('type', type);
    const data = await GET(`/enablement/resources?${params}`);
    S.enResources = data.resources || [];
    S.enExportButtons = data.exportButtons || [];
    renderEnablementExportButtons(S.enExportButtons);
    renderEnCards(S.enResources);
  } catch (e) {
    el.innerHTML = `<div class="empty-state"><div class="em-icon">⚠️</div><div>${escHtml(e.message)}</div></div>`;
  }
}

function renderEnablementExportButtons(buttons) {
  const wrap = document.getElementById('enExportButtons');
  if (!wrap) return;
  if (!buttons?.length) {
    wrap.innerHTML = '';
    return;
  }
  wrap.innerHTML = '';
  buttons.forEach(btn => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'btn btn-xs en-export-pack-btn';
    button.textContent = `📦 ${btn.text || btn.id}`;
    button.onclick = () => exportEnablementByButton(btn.id, btn.text || btn.id);
    wrap.appendChild(button);
  });
}

function filterEn(type, btn) {
  S.enFilter = type;
  document.querySelectorAll('#enFilterBtns button').forEach(b => {
    b.className = 'btn btn-xs ' + (b.dataset.filter === type ? 'btn-primary' : 'btn-secondary');
  });
  loadEnablement();
}

function renderEnCards(resources) {
  window._enResources = resources;
  const el = document.getElementById('enContent');
  const perRow = parseInt(document.getElementById('enPerRow').value) || 3;
  const typeClass = { video: 'en-type-video', guide: 'en-type-guide', doc: 'en-type-doc', template: 'en-type-template' };
  if (!resources.length) {
    el.innerHTML = '<div class="empty-state"><div class="em-icon">🔍</div><div>No resources match your filter</div></div>';
    return;
  }
  el.innerHTML = `<div style="display:grid;grid-template-columns:repeat(${perRow},1fr);gap:14px">` +
    resources.map(r => `
      <div style="border:1px solid var(--border);border-radius:10px;background:var(--surface);padding:16px;display:flex;flex-direction:column;transition:box-shadow .15s" onmouseover="this.style.boxShadow='0 4px 18px rgba(0,0,0,.08)'" onmouseout="this.style.boxShadow=''">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
          <label class="chk-label" style="font-size:12px"><input type="checkbox" data-en-id="${r.id}" style="width:13px;height:13px"> </label>
          <span class="en-type-badge ${typeClass[r.type] || 'en-type-guide'}">${escHtml(r.type)}</span>
        </div>
        <div style="font-size:15px;font-weight:700;margin-bottom:6px;line-height:1.3">${escHtml(r.icon || '')} ${escHtml(r.title)}</div>
        <div style="font-size:12.5px;color:var(--t2);flex:1;line-height:1.5;margin-bottom:12px">${escHtml(r.description)}</div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <a href="${escHtml(r.url)}" target="_blank" rel="noopener" class="btn btn-secondary btn-xs">↗ Open</a>
          <button class="btn btn-secondary btn-xs" onclick="copyText('${escHtml(r.url)}')">🔗 Copy Link</button>
        </div>
      </div>`).join('') + '</div>';
}

function selectAllEn(cb) {
  document.querySelectorAll('#enContent input[type=checkbox][data-en-id]').forEach(c => c.checked = cb.checked);
}

function getSelectedEnablementResources() {
  const checkedIds = [...document.querySelectorAll('#enContent input[data-en-id]:checked')].map(c => String(c.dataset.enId));
  if (checkedIds.length) {
    const selectedIdSet = new Set(checkedIds);
    return (S.enResources || []).filter(r => selectedIdSet.has(String(r.id)));
  }
  return S.enResources || [];
}

function openEnablementExportModal() {
  const selectedItems = getSelectedEnablementResources();
  const selectedCount = [...document.querySelectorAll('#enContent input[data-en-id]:checked')].length;
  const summary = document.getElementById('enablementExportSummary');
  if (summary) {
    summary.textContent = selectedCount
      ? `${selectedItems.length} selected resource(s) ready. Choose export or email.`
      : `${selectedItems.length} filtered resource(s) ready. Choose export or email.`;
  }
  openModal('enablementExportModal');
}

async function performEnablementExport() {
  const checked = [...document.querySelectorAll('#enContent input[data-en-id]:checked')].map(c => c.dataset.enId);
  const layout = 'cards';
  log(`Exporting ${checked.length || 'all'} enablement resource(s)...`);
  await downloadApi('POST', '/enablement/export', { ids: checked, layout, title: 'Litera One Enablement Resources' }, 'litera_enablement.html');
  closeModal('enablementExportModal');
  toast(`Exported ${checked.length || S.enResources.length} resource(s)`, 'success');
}

function buildEnablementEmailBody(items) {
  return [
    'Hi,',
    '',
    'Sharing the selected Enablement Hub resources:',
    '',
    ...items.map((r, index) => `${index + 1}. ${r.title}\n${r.url}${r.description ? `\n${r.description}` : ''}`),
    '',
    'Regards,'
  ].join('\n');
}

function buildEnablementEmailHtml(items) {
  const rows = items.map((r, index) => `
    <tr>
      <td style="padding:12px 0;border-top:1px solid #e2e8f0">
        <div style="font-size:12px;font-weight:700;color:#0284c7;text-transform:uppercase;margin-bottom:6px">${escHtml(r.type || 'Resource')}</div>
        <div style="font-size:18px;font-weight:700;color:#0f172a;margin-bottom:6px">${index + 1}. ${escHtml(r.title || '')}</div>
        ${r.description ? `<div style="font-size:14px;line-height:1.6;color:#475569;margin-bottom:10px">${escHtml(r.description)}</div>` : ''}
        <a href="${escHtml(r.url || '')}" style="display:inline-block;padding:8px 14px;background:#0284c7;color:#ffffff;text-decoration:none;border-radius:8px;font-size:13px;font-weight:600">Open Resource</a>
        <div style="margin-top:8px;font-size:12px;color:#64748b">${escHtml(r.url || '')}</div>
      </td>
    </tr>`).join('');

  return `<!DOCTYPE html>
<html>
  <body style="margin:0;padding:24px;background:#f8fafc;font-family:Segoe UI,Arial,sans-serif;color:#0f172a">
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:760px;margin:0 auto;background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;padding:28px">
      <tr>
        <td>
          <div style="font-size:26px;font-weight:800;margin-bottom:6px">Litera One Enablement Resources</div>
          <div style="font-size:14px;line-height:1.6;color:#475569;margin-bottom:18px">Sharing ${items.length} selected resource(s) from the Enablement Hub.</div>
          <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
            ${rows}
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>`;
}

async function copyEnablementEmailToClipboard(html, text) {
  if (navigator.clipboard?.write && typeof ClipboardItem !== 'undefined') {
    const item = new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([text], { type: 'text/plain' }),
    });
    await navigator.clipboard.write([item]);
    return 'html';
  }

  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return 'text';
  }

  throw new Error('Clipboard access is unavailable.');
}

async function emailEnablementSelection() {
  const items = getSelectedEnablementResources();
  if (!items.length) {
    toast('No resources available to email', 'error');
    return;
  }
  const subject = `Litera One Enablement Resources (${items.length})`;
  const textBody = buildEnablementEmailBody(items);
  const htmlBody = buildEnablementEmailHtml(items);

  try {
    const copyMode = await copyEnablementEmailToClipboard(htmlBody, textBody);
    const mailtoBody = copyMode === 'html'
      ? 'HTML email content has been copied to your clipboard. Paste it into the message body.'
      : textBody;
    const mailto = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(mailtoBody)}`;
    closeModal('enablementExportModal');
    window.location.href = mailto;
    toast(
      copyMode === 'html'
        ? `Opened email draft and copied HTML for ${items.length} resource(s)`
        : `Opened email draft for ${items.length} resource(s)`,
      'success'
    );
  } catch (e) {
    toast(`Unable to prepare email: ${e.message}`, 'error');
  }
}

async function exportEnablementByButton(exportButtonId, text) {
  if (!exportButtonId) return;
  await downloadApi('POST', '/enablement/export', { exportButtonId, layout: 'cards', title: 'Litera One Enablement Resources' }, 'litera_enablement.html');
  toast(`Exported: ${text || exportButtonId}`, 'success');
}

// ═══════════════════════════════════════════════════════════════════════════════
//  ENABLEMENT ADMIN
// ═══════════════════════════════════════════════════════════════════════════════

function _setEnablementAdminStatus(msg, isError = false) {
  const el = document.getElementById('enablementAdminStatus');
  if (!el) return;
  el.textContent = msg;
  el.style.color = isError ? 'var(--danger)' : 'var(--t2)';
}

function _setEnablementAdminSubtitle(loggedIn) {
  const el = document.getElementById('enablementAdminSubtitle');
  if (!el) return;
  el.textContent = loggedIn
    ? 'Manage enablement resources, export buttons, and settings'
    : 'Sign in to manage enablement resources, export buttons, and settings';
}

function _toggleEnablementAdminWorkspace(loggedIn) {
  const authCard = document.getElementById('enablementAdminAuthCard');
  const workspace = document.getElementById('enablementAdminWorkspace');
  if (authCard) authCard.style.display = loggedIn ? 'none' : 'block';
  if (workspace) workspace.style.display = loggedIn ? 'block' : 'none';
  _setEnablementAdminSubtitle(loggedIn);
  if (window.__STANDALONE_ENABLEMENT_ADMIN__) {
    document.body.classList.toggle('enablement-admin-auth-view', !loggedIn);
  }
}

function toggleEnablementJsonEditor() {
  S.enablementJsonVisible = !S.enablementJsonVisible;
  const card = document.getElementById('enablementJsonEditorCard');
  const btn = document.getElementById('toggleEnablementJsonBtn');
  if (card) card.style.display = S.enablementJsonVisible ? 'block' : 'none';
  if (btn) btn.textContent = S.enablementJsonVisible ? 'Hide JSON Editor' : 'Show JSON Editor';
}

function openEnablementResourceForm(type) {
  S.selectedEnablementResourceType = String(type || '').trim() || 'Custom';
  const form = document.getElementById('enablementResourceForm');
  const label = document.getElementById('selectedEnablementResourceType');
  const typeInput = document.getElementById('newResourceType');
  if (form) form.style.display = 'block';
  if (label) label.textContent = S.selectedEnablementResourceType;
  if (typeInput) typeInput.value = S.selectedEnablementResourceType;
  _markSelectedEnablementTypeButton(S.selectedEnablementResourceType);
}

function _renderEnablementResourceTypeButtons() {
  const wrap = document.getElementById('enablementResourceTypeButtons');
  if (!wrap) return;
  const resources = S.enablementAdminConfig?.resources || [];
  const types = [...new Set(resources.map(r => String(r.type || '').trim()).filter(Boolean))];
  const clsByType = {
    video: 'enablement-type-video',
    guide: 'enablement-type-guide',
    doc: 'enablement-type-doc',
    template: 'enablement-type-template',
    article: 'enablement-type-article',
  };
  wrap.innerHTML = '';
  if (!types.length) {
    wrap.innerHTML = '<span class="text-xs">No existing resource types found.</span>';
    return;
  }
  types.forEach(type => {
    const key = String(type).toLowerCase();
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `btn btn-xs enablement-type-btn ${clsByType[key] || 'enablement-type-default'}`;
    btn.textContent = type;
    btn.onclick = () => openEnablementResourceForm(type);
    wrap.appendChild(btn);
  });
}

function _markSelectedEnablementTypeButton(type) {
  const buttons = [...document.querySelectorAll('#enablementResourceTypeButtons .enablement-type-btn')];
  buttons.forEach(btn => {
    btn.classList.toggle('active', btn.textContent === type);
  });
}

function _renderCreatedResourcePreview(resource) {
  const box = document.getElementById('enablementResourcePreview');
  if (!box || !resource) return;
  const type = String(resource.type || 'Resource');
  const badgeClass = {
    video: 'en-type-video',
    guide: 'en-type-guide',
    doc: 'en-type-doc',
    template: 'en-type-template',
  }[type.toLowerCase()] || 'badge-gray';
  box.style.display = 'block';
  box.innerHTML = `
    <div class="text-sm" style="font-weight:700;margin-bottom:6px;color:#166534">Resource created successfully</div>
    <div class="en-type-badge ${badgeClass}">${escHtml(type)}</div>
    <div style="font-weight:700;margin-bottom:4px">${escHtml(resource.icon || '')} ${escHtml(resource.title || '')}</div>
    <div class="text-xs mb-2">${escHtml(resource.description || '')}</div>
    <a href="${escHtml(resource.url || '#')}" target="_blank" rel="noopener" class="btn btn-secondary btn-xs">↗ Open Resource</a>
  `;
}

function _setEnablementResourceFormMode(isEdit) {
  const mode = document.getElementById('enablementResourceFormMode');
  const submit = document.getElementById('submitEnablementResourceBtn');
  const cancel = document.getElementById('cancelEnablementResourceEditBtn');
  const idInput = document.getElementById('newResourceId');
  if (mode) mode.textContent = isEdit ? 'Edit' : 'Create';
  if (submit) submit.textContent = isEdit ? 'Update Resource' : '+ Create Resource';
  if (cancel) cancel.style.display = isEdit ? 'inline-flex' : 'none';
  if (idInput) idInput.disabled = !!isEdit;
}

function cancelEnablementResourceEdit() {
  S.enablementAdminEditResourceId = null;
  _setEnablementResourceFormMode(false);
  const preview = document.getElementById('enablementResourcePreview');
  if (preview) preview.style.display = 'none';
  ['newResourceId', 'newResourceTitle', 'newResourceType', 'newResourceIcon', 'newResourceUrl', 'newResourceDescription', 'newResourceTags'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
}

function _renderEnablementAdminResourceList() {
  const tbody = document.getElementById('enablementAdminResourceTbody');
  if (!tbody) return;
  const resources = S.enablementAdminConfig?.resources || [];
  if (!resources.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="tbl-empty">No resources found</td></tr>';
    return;
  }
  tbody.innerHTML = '';
  resources.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="mono">${escHtml(r.id)}</td>
      <td>${escHtml(r.title)}</td>
      <td><span class="badge badge-blue">${escHtml(r.type)}</span></td>
      <td class="truncate" style="max-width:260px">${escHtml(r.url)}</td>
      <td>
        <div class="flex gap-2">
          <button type="button" class="btn btn-secondary btn-xs">Edit</button>
          <button type="button" class="btn btn-secondary btn-xs">Copy to New</button>
          <button type="button" class="btn btn-danger btn-xs">Delete</button>
        </div>
      </td>
    `;
    const [editBtn, copyBtn, delBtn] = tr.querySelectorAll('button');
    editBtn.onclick = () => startEditEnablementResource(r.id);
    copyBtn.onclick = () => copyEnablementResourceToNew(r.id);
    delBtn.onclick = () => deleteEnablementResource(r.id);
    tbody.appendChild(tr);
  });
}

function startEditEnablementResource(id) {
  const resources = S.enablementAdminConfig?.resources || [];
  const item = resources.find(r => String(r.id) === String(id));
  if (!item) {
    toast('Resource not found', 'error');
    return;
  }
  S.enablementAdminEditResourceId = item.id;
  openEnablementResourceForm(item.type || 'Custom');
  _setEnablementResourceFormMode(true);
  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = val || '';
  };
  setVal('newResourceId', item.id);
  setVal('newResourceTitle', item.title);
  setVal('newResourceType', item.type);
  setVal('newResourceIcon', item.icon);
  setVal('newResourceUrl', item.url);
  setVal('newResourceDescription', item.description);
  setVal('newResourceTags', (item.tags || []).join(', '));
}

function copyEnablementResourceToNew(id) {
  const resources = S.enablementAdminConfig?.resources || [];
  const item = resources.find(r => String(r.id) === String(id));
  if (!item) {
    toast('Resource not found', 'error');
    return;
  }
  S.enablementAdminEditResourceId = null;
  openEnablementResourceForm(item.type || 'Custom');
  _setEnablementResourceFormMode(false);
  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = val || '';
  };
  setVal('newResourceId', `${item.id}-copy`);
  setVal('newResourceTitle', item.title);
  setVal('newResourceType', item.type);
  setVal('newResourceIcon', item.icon);
  setVal('newResourceUrl', item.url);
  setVal('newResourceDescription', item.description);
  setVal('newResourceTags', (item.tags || []).join(', '));
}

function _setEnablementExportButtonFormMode(isEdit) {
  const mode = document.getElementById('enablementExportButtonFormMode');
  const submit = document.getElementById('submitEnablementExportButtonBtn');
  const cancel = document.getElementById('cancelEnablementExportButtonEditBtn');
  const idInput = document.getElementById('newExportButtonId');
  if (mode) mode.textContent = isEdit ? 'Edit' : 'Create';
  if (submit) submit.textContent = isEdit ? 'Update Export Button' : '+ Create Export Button';
  if (cancel) cancel.style.display = isEdit ? 'inline-flex' : 'none';
  if (idInput) idInput.disabled = !!isEdit;
}

function cancelEnablementExportButtonEdit() {
  S.enablementAdminEditExportButtonId = null;
  _setEnablementExportButtonFormMode(false);
  ['newExportButtonId', 'newExportButtonText', 'newExportButtonResourceIds', 'newExportButtonType', 'newExportButtonTag'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
}

function _exportButtonRuleText(button) {
  if ((button.resourceIds || []).length) return `IDs: ${(button.resourceIds || []).join(', ')}`;
  if (button.resourceType) return `Type: ${button.resourceType}`;
  if (button.resourceTag) return `Tag: ${button.resourceTag}`;
  return 'All resources';
}

function _renderEnablementAdminExportButtonList() {
  const tbody = document.getElementById('enablementAdminExportButtonTbody');
  if (!tbody) return;
  const items = S.enablementAdminConfig?.exportButtons || [];
  if (!items.length) {
    tbody.innerHTML = '<tr><td colspan="4" class="tbl-empty">No export buttons found</td></tr>';
    return;
  }
  tbody.innerHTML = '';
  items.forEach(btn => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="mono">${escHtml(btn.id)}</td>
      <td>${escHtml(btn.text)}</td>
      <td class="text-xs">${escHtml(_exportButtonRuleText(btn))}</td>
      <td>
        <div class="flex gap-2">
          <button type="button" class="btn btn-secondary btn-xs">Edit</button>
          <button type="button" class="btn btn-secondary btn-xs">Copy to New</button>
          <button type="button" class="btn btn-danger btn-xs">Delete</button>
        </div>
      </td>
    `;
    const [editBtn, copyBtn, delBtn] = tr.querySelectorAll('button');
    editBtn.onclick = () => startEditEnablementExportButton(btn.id);
    copyBtn.onclick = () => copyEnablementExportButtonToNew(btn.id);
    delBtn.onclick = () => deleteEnablementExportButton(btn.id);
    tbody.appendChild(tr);
  });
}

function startEditEnablementExportButton(id) {
  const items = S.enablementAdminConfig?.exportButtons || [];
  const item = items.find(b => String(b.id) === String(id));
  if (!item) {
    toast('Export button not found', 'error');
    return;
  }
  S.enablementAdminEditExportButtonId = item.id;
  _setEnablementExportButtonFormMode(true);
  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = val || '';
  };
  setVal('newExportButtonId', item.id);
  setVal('newExportButtonText', item.text);
  setVal('newExportButtonResourceIds', (item.resourceIds || []).join(', '));
  setVal('newExportButtonType', item.resourceType);
  setVal('newExportButtonTag', item.resourceTag);
}

function copyEnablementExportButtonToNew(id) {
  const items = S.enablementAdminConfig?.exportButtons || [];
  const item = items.find(b => String(b.id) === String(id));
  if (!item) {
    toast('Export button not found', 'error');
    return;
  }
  S.enablementAdminEditExportButtonId = null;
  _setEnablementExportButtonFormMode(false);
  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = val || '';
  };
  setVal('newExportButtonId', `${item.id}-copy`);
  setVal('newExportButtonText', item.text);
  setVal('newExportButtonResourceIds', (item.resourceIds || []).join(', '));
  setVal('newExportButtonType', item.resourceType);
  setVal('newExportButtonTag', item.resourceTag);
}

function toggleEnablementAdminPassword() {
  const pass = document.getElementById('enablementAdminPass');
  if (!pass) return;
  pass.type = pass.type === 'password' ? 'text' : 'password';
}

function showEnablementAdminResetHint() {
  toast('Reset admin credentials in .env using ENABLEMENT_ADMIN_USERNAME and ENABLEMENT_ADMIN_PASSWORD', 'error');
}

async function refreshEnablementAdmin() {
  try {
    const data = await GET('/enablement/admin/status');
    if (!data.configured) {
      _setEnablementAdminStatus('Admin auth is not configured. Set ENABLEMENT_ADMIN_USERNAME and ENABLEMENT_ADMIN_PASSWORD in .env.', true);
      _toggleEnablementAdminWorkspace(false);
      return;
    }
    if (data.loggedIn) {
      _setEnablementAdminStatus(`Logged in as ${data.username}`);
      const ws = document.getElementById('enablementAdminWorkspaceStatus');
      if (ws) ws.textContent = `Logged in as ${data.username}`;
      _toggleEnablementAdminWorkspace(true);
      await loadEnablementAdminConfig();
    } else {
      _setEnablementAdminStatus('Not logged in');
      _toggleEnablementAdminWorkspace(false);
    }
  } catch (e) {
    _setEnablementAdminStatus(e.message || 'Unable to load admin status', true);
    _toggleEnablementAdminWorkspace(false);
  }
}

async function loginEnablementAdmin() {
  const username = document.getElementById('enablementAdminUser')?.value?.trim();
  const password = document.getElementById('enablementAdminPass')?.value?.trim();
  if (!username || !password) {
    toast('Enter admin username and password', 'error');
    return;
  }
  try {
    await POST('/enablement/admin/login', { username, password });
    toast('Enablement admin login successful', 'success');
    await refreshEnablementAdmin();
  } catch (e) {
    _setEnablementAdminStatus(e.message || 'Login failed', true);
    toast('Enablement admin login failed', 'error');
  }
}

async function logoutEnablementAdmin() {
  try {
    await POST('/enablement/admin/logout', {});
    _setEnablementAdminStatus('Logged out');
    _toggleEnablementAdminWorkspace(false);
    S.enablementAdminConfig = null;
    S.selectedEnablementResourceType = null;
    S.enablementJsonVisible = false;
    S.enablementAdminEditResourceId = null;
    S.enablementAdminEditExportButtonId = null;
    const jsonCard = document.getElementById('enablementJsonEditorCard');
    if (jsonCard) jsonCard.style.display = 'none';
    const jsonBtn = document.getElementById('toggleEnablementJsonBtn');
    if (jsonBtn) jsonBtn.textContent = 'Show JSON Editor';
    const resourceForm = document.getElementById('enablementResourceForm');
    if (resourceForm) resourceForm.style.display = 'none';
    _setEnablementResourceFormMode(false);
    _setEnablementExportButtonFormMode(false);
    const editor = document.getElementById('enablementAdminEditor');
    if (editor) editor.value = '';
    toast('Enablement admin logged out', 'success');
  } catch (e) {
    _setEnablementAdminStatus(e.message || 'Logout failed', true);
  }
}

async function loadEnablementAdminConfig() {
  try {
    const data = await GET('/enablement/admin/config');
    S.enablementAdminConfig = data.config || {};
    if (!S.enablementAdminEditResourceId) _setEnablementResourceFormMode(false);
    if (!S.enablementAdminEditExportButtonId) _setEnablementExportButtonFormMode(false);
    _renderEnablementResourceTypeButtons();
    _renderEnablementAdminResourceList();
    _renderEnablementAdminExportButtonList();
    const editor = document.getElementById('enablementAdminEditor');
    if (editor) editor.value = JSON.stringify(data.config || {}, null, 2);
  } catch (e) {
    _setEnablementAdminStatus(e.message || 'Load config failed', true);
  }
}

async function saveEnablementAdminConfig() {
  const editor = document.getElementById('enablementAdminEditor');
  if (!editor) return;
  let parsed;
  try {
    parsed = JSON.parse(editor.value || '{}');
  } catch (e) {
    toast('Invalid JSON in editor', 'error');
    return;
  }
  try {
    const data = await api('PUT', '/enablement/admin/config', parsed);
    S.enablementAdminConfig = data.config || {};
    _renderEnablementResourceTypeButtons();
    _renderEnablementAdminResourceList();
    _renderEnablementAdminExportButtonList();
    editor.value = JSON.stringify(data.config || {}, null, 2);
    toast('Enablement config saved', 'success');
    await loadEnablement();
  } catch (e) {
    _setEnablementAdminStatus(e.message || 'Save failed', true);
    toast('Failed to save enablement config', 'error');
  }
}

function _splitCsv(value) {
  return String(value || '')
    .split(',')
    .map(v => v.trim())
    .filter(Boolean);
}

async function createEnablementResource() {
  const id = document.getElementById('newResourceId')?.value?.trim();
  const title = document.getElementById('newResourceTitle')?.value?.trim();
  const type = document.getElementById('newResourceType')?.value?.trim() || S.selectedEnablementResourceType;
  const icon = document.getElementById('newResourceIcon')?.value?.trim();
  const url = document.getElementById('newResourceUrl')?.value?.trim();
  const description = document.getElementById('newResourceDescription')?.value?.trim();
  const tags = _splitCsv(document.getElementById('newResourceTags')?.value);

  if (!id || !title || !type || !url) {
    toast('Resource id, title, type, and url are required', 'error');
    return;
  }

  const payload = { id, title, type, url, description };
  if (icon) payload.icon = icon;
  if (tags.length) payload.tags = tags;

  try {
    const data = S.enablementAdminEditResourceId
      ? await api('PUT', `/enablement/admin/resources/${encodeURIComponent(S.enablementAdminEditResourceId)}`, payload)
      : await POST('/enablement/admin/resources', payload);
    toast(S.enablementAdminEditResourceId ? 'Resource updated' : 'Resource created', 'success');
    await refreshEnablementAdmin();
    await loadEnablement();
    _renderCreatedResourcePreview(data.resource || payload);
    S.enablementAdminEditResourceId = null;
    _setEnablementResourceFormMode(false);
    ['newResourceId', 'newResourceTitle', 'newResourceIcon', 'newResourceUrl', 'newResourceDescription', 'newResourceTags'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
  } catch (e) {
    toast(e.message || 'Failed to create resource', 'error');
  }
}

async function createEnablementExportButton() {
  const id = document.getElementById('newExportButtonId')?.value?.trim();
  const text = document.getElementById('newExportButtonText')?.value?.trim();
  const resourceIds = _splitCsv(document.getElementById('newExportButtonResourceIds')?.value);
  const resourceType = document.getElementById('newExportButtonType')?.value?.trim();
  const resourceTag = document.getElementById('newExportButtonTag')?.value?.trim();

  if (!text) {
    toast('Export button text is required', 'error');
    return;
  }

  const payload = { text };
  if (id) payload.id = id;
  if (resourceIds.length) payload.resourceIds = resourceIds;
  if (resourceType) payload.resourceType = resourceType;
  if (resourceTag) payload.resourceTag = resourceTag;

  try {
    if (S.enablementAdminEditExportButtonId) {
      await api('PUT', `/enablement/admin/export-buttons/${encodeURIComponent(S.enablementAdminEditExportButtonId)}`, payload);
    } else {
      await POST('/enablement/admin/export-buttons', payload);
    }
    toast(S.enablementAdminEditExportButtonId ? 'Export button updated' : 'Export button created', 'success');
    S.enablementAdminEditExportButtonId = null;
    _setEnablementExportButtonFormMode(false);
    ['newExportButtonId', 'newExportButtonText', 'newExportButtonResourceIds', 'newExportButtonType', 'newExportButtonTag'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    await refreshEnablementAdmin();
    await loadEnablement();
  } catch (e) {
    toast(e.message || 'Failed to create export button', 'error');
  }
}

async function deleteEnablementResource(fromRowId) {
  const id = String(fromRowId || document.getElementById('deleteResourceId')?.value?.trim() || '');
  if (!id) {
    toast('Enter a resource id to delete', 'error');
    return;
  }
  try {
    await DEL(`/enablement/admin/resources/${encodeURIComponent(id)}`);
    toast('Resource deleted', 'success');
    await refreshEnablementAdmin();
    await loadEnablement();
  } catch (e) {
    toast(e.message || 'Failed to delete resource', 'error');
  }
}

async function deleteEnablementExportButton(fromRowId) {
  const id = String(fromRowId || document.getElementById('deleteExportButtonId')?.value?.trim() || '');
  if (!id) {
    toast('Enter an export button id to delete', 'error');
    return;
  }
  try {
    await DEL(`/enablement/admin/export-buttons/${encodeURIComponent(id)}`);
    toast('Export button deleted', 'success');
    await refreshEnablementAdmin();
    await loadEnablement();
  } catch (e) {
    toast(e.message || 'Failed to delete export button', 'error');
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════════════════════════════
(function init() {
  log('Litera One Manager Portal loaded.');
  updateDeployProduct();
  updateSourceUi();
  updateMethodUi();

  const bindTenantForm = (formId, selectId, hiddenId, resetId, rememberId) => {
    const form = document.getElementById(formId);
    if (!form) return;
    const tenantSelect = document.getElementById(selectId);
    const tenantHidden = document.getElementById(hiddenId);
    const tenantResetBtn = document.getElementById(resetId);
    const rememberTenant = rememberId ? document.getElementById(rememberId) : null;

    const savedTenant = localStorage.getItem('lo_tenant');
    if (savedTenant && tenantSelect && formId === 'signinForm') tenantSelect.value = savedTenant;
    if (rememberTenant) rememberTenant.checked = !!savedTenant;
    if (tenantHidden && tenantSelect) tenantHidden.value = (tenantSelect.value || 'organizations').trim() || 'organizations';

    if (tenantSelect && tenantHidden) {
      tenantSelect.addEventListener('change', () => {
        tenantHidden.value = (tenantSelect.value || 'organizations').trim() || 'organizations';
      });
      tenantSelect.addEventListener('input', () => {
        tenantHidden.value = (tenantSelect.value || 'organizations').trim() || 'organizations';
      });
    }

    if (tenantResetBtn && tenantSelect && tenantHidden) {
      tenantResetBtn.addEventListener('click', () => {
        tenantSelect.value = 'organizations';
        tenantHidden.value = 'organizations';
      });
    }

    form.addEventListener('submit', () => {
      const tenant = (tenantSelect?.value || 'organizations').trim() || 'organizations';
      if (tenantHidden) tenantHidden.value = tenant;
      if (rememberTenant?.checked) localStorage.setItem('lo_tenant', tenant);
      else if (rememberTenant) localStorage.removeItem('lo_tenant');
    });
  };

  bindTenantForm('signinForm', 'tenantSelect', 'tenantHidden', 'tenantResetBtn', 'rememberTenant');
  bindTenantForm('switchTenantForm', 'switchTenantSelect', 'switchTenantHidden', 'switchTenantResetBtn');
  updateAppsAuthUi();
  setUsersManagementMode('assign', document.getElementById('lumModeAssignBtn'));
  renderUsersManagementHistory();

  const initialPage = window.__INITIAL_PAGE__ || 'dashboard';
  if (initialPage !== 'dashboard') navTo(initialPage);
  if (window.__STANDALONE_ENABLEMENT_ADMIN__) navTo('enablement-admin');

  loadEnablement();

  // Keyboard: Escape closes top-most modal
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      const open = [...document.querySelectorAll('.modal-backdrop')].filter(m => m.style.display === 'flex');
      if (open.length) open[open.length - 1].style.display = 'none';
    }
  });

  // Enter in app search triggers load
  document.getElementById('appSearch')?.addEventListener('keydown', e => {
    if (e.key === 'Enter') loadApps();
  });

  // Enter on enablement admin login fields submits login
  ['enablementAdminUser', 'enablementAdminPass'].forEach(id => {
    document.getElementById(id)?.addEventListener('keydown', e => {
      if (e.key === 'Enter') loginEnablementAdmin();
    });
  });
})();
