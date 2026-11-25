/**********************************************************************
 *  EMPLOYEE TASK MANAGEMENT SYSTEM — Full Merge + WhatsApp Cloud API
 *  + AUTO ATTENDANCE (first login of the day = PRESENT)
 *
 *  Search this file for:
 *    "WHATSAPP CLOUD API"  -> WhatsApp changes
 *    "ATTENDANCE"          -> attendance section (end of file)
 *    "AUTO PUNCH"          -> the hook inside loginUser()
 **********************************************************************/

const SHEET_ID = ''; // <-- the sheet with Users + SheetLinks tabs
const USERS_SHEET = 'Users';
const SHEETLINKS_SHEET = 'SheetLinks';
const NOTIFICATIONS_SHEET = 'Notifications';
const APPS_SHEET = 'Apps';
const DONE_KEYWORDS = ['done', 'complete', 'completed', 'closed', 'yes'];

const ADMIN_BROADCAST_KEY = '__ADMIN__';

const ADMIN_ROLES_ = ['admin', 'manager'];
function isAdminRole_(role) {
  return ADMIN_ROLES_.indexOf((role || '').toString().trim().toLowerCase()) !== -1;
}

const COLUMN_KEYWORDS_ = {
  date:      { keywords: ['due', 'deadline', 'planned', 'target', 'date'], exclude: ['actual', 'done date', 'completed on', '%'] },
  status:    { keywords: ['status', 'done', 'complete'], exclude: [] },
  task:      { keywords: ['task', 'item', 'work'], exclude: ['id'] },
  assigned:  { keywords: ['doer', 'assigned', 'owner', 'responsible', 'accountable', 'name'], exclude: [] },
  desc:      { keywords: ['description', 'detail', 'note', 'remark'], exclude: [] },
  priority:  { keywords: ['priority'], exclude: [] }
};

// ============================================================
// TASK-SYSTEM CONFIGURATION: one entry per employee's own task sheet
// ============================================================
const SHEETS_CONFIG = {
  'person1': { id: '', name: 'Brijpal_ji', sheetName: 'Sheet1' },
  'person2': { id: '', name: 'Mukesh',     sheetName: 'Sheet1' },
  'person3': { id: '', name: 'Sayeed',     sheetName: 'Sheet1' },
  'person4': { id: '', name: 'Prashant',   sheetName: 'Sheet1' },
  'person5': { id: '', name: 'Piyush',     sheetName: 'Sheet1' },
  'person6': { id: '', name: 'Sonu',       sheetName: 'Ajay sir' },
  'person7': { id: '', name: 'Nitesh',     sheetName: 'Sheet1' }
};

const DATA_START_ROW = 4;

// ============================================================
// WEB APP ENTRY POINT
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Employee Task Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// NAME NORMALIZATION
// ============================================================
function normalizeName_(s) {
  return (s || '').toString().trim().toLowerCase().replace(/[_\s]+/g, ' ').trim();
}

// ============================================================
// BIG CACHE: CacheService allows only 100 KB per key, so large lists are split into chunks.
// (The old single-key cache silently failed for the task list, so every load re-read all the sheets.)
// ============================================================
const BIG_CHUNK_ = 30000;
const TASKS_CACHE_KEY = 'allTasks_v2';
const TASKS_CACHE_SECONDS = 600;

function cachePutBig_(key, value, ttl) {
  try {
    const str = JSON.stringify(value);
    const n = Math.max(1, Math.ceil(str.length / BIG_CHUNK_));
    const obj = {};
    for (let i = 0; i < n; i++) obj[key + '_' + i] = str.substr(i * BIG_CHUNK_, BIG_CHUNK_);
    obj[key] = String(n);
    CacheService.getScriptCache().putAll(obj, ttl || 300);
    return true;
  } catch (e) { Logger.log('cachePutBig_ ' + key + ': ' + e); return false; }
}

function cacheGetBig_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(key));
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const got = cache.getAll(keys);
    let str = '';
    for (let i = 0; i < n; i++) { const part = got[key + '_' + i]; if (part == null) return null; str += part; }
    return JSON.parse(str);
  } catch (e) { return null; }
}

function cacheRemoveBig_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(key)) || 0;
    const keys = [key];
    for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    cache.removeAll(keys);
  } catch (e) { /* ignore */ }
}

// Pending-work caches carry a version number; bumping it drops every pending cache at once.
function pendCacheKey_(name) {
  return 'pend_' + (CacheService.getScriptCache().get('pendVer') || '0') + '_' + name;
}
function bumpPendCache_() { CacheService.getScriptCache().put('pendVer', String(Date.now()), 21600); }

// ============================================================
// CENTRAL SHEET HELPER
// ============================================================
let _centralSS_ = null;                       // opened once per run instead of on every helper call
function getCentralSS_() { return _centralSS_ || (_centralSS_ = SpreadsheetApp.openById(SHEET_ID)); }

function getCentralSheet_(name) {
  const ss = getCentralSS_();
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  const target = name.toLowerCase().replace(/\s+/g, '');
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) {
    if (all[i].getName().toLowerCase().replace(/\s+/g, '') === target) return all[i];
  }
  return null;
}

function getTabFuzzy_(ss, tabName) {
  let tab = ss.getSheetByName(tabName);
  if (tab) return tab;
  const target = tabName.toLowerCase().replace(/\s+/g, '');
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) {
    if (all[i].getName().toLowerCase().replace(/\s+/g, '') === target) return all[i];
  }
  return null;
}

function pickDefaultTab_(ss) {
  const allSheets = ss.getSheets();
  if (allSheets.length <= 1) return allSheets[0];

  const exactPreferred = allSheets.find(s => s.getName().trim().toLowerCase() === 'fms');
  if (exactPreferred) return exactPreferred;

  const skipPatterns = [
    /^form\s*responses?/i,
    /^config$/i,
    /^holidays?$/i,
    /^instructions$/i,
    /^setup\s*sheet$/i,
    /^dashboard$/i
  ];
  const candidates = allSheets.filter(s => !skipPatterns.some(p => p.test(s.getName().trim())));
  const pool = candidates.length > 0 ? candidates : allSheets;

  let best = pool[0];
  let bestScore = -1;
  pool.forEach(s => {
    const score = s.getLastRow() * s.getLastColumn();
    if (score > bestScore) { bestScore = score; best = s; }
  });
  return best;
}

// ============================================================
// AUTH
// ============================================================
function loginUser(username, password) {
  username = (username || '').toString().trim();
  password = (password || '').toString().trim();

  if (!username || !password) {
    return { success: false, message: 'Please enter both username and password.' };
  }

  const sheet = getCentralSheet_(USERS_SHEET);
  if (!sheet) {
    return { success: false, message: 'Users tab not found. Check SHEET_ID and tab name.' };
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sheetUser = (row[0] || '').toString().trim();
    const sheetPass = (row[1] || '').toString().trim();

    if (sheetUser.toLowerCase() === username.toLowerCase() && sheetPass === password) {
      const fullName = row[2] || sheetUser;
      let role = (row[3] || 'Employee').toString().trim();
      const roleLower = role.toLowerCase();
      if (roleLower === 'admin') role = 'Admin';
      else if (roleLower === 'manager') role = 'Manager';
      const personKey = (row[4] || '').toString().trim() || null;

      const token = Utilities.getUuid();
      const session = { username: sheetUser, fullName: fullName, role: role, personKey: personKey, exp: sessionExpiry_() };
      sessionSave_(token, session);

      // === AUTO PUNCH === now done by the app right after login (markAttendance), so login itself
      // returns immediately instead of waiting for the attendance spreadsheet to open.

      return { success: true, token: token, username: sheetUser, fullName: fullName, role: role, personKey: personKey };
    }
  }
  return { success: false, message: 'Invalid username or password. Please try again.' };
}

// === SESSION LENGTH ===
// A login stays valid until 7:00 PM India time the same day (Google's cache alone can only keep
// a login for 6 hours, so the session is also saved in Script Properties with its own end time).
// A login made AFTER 7 PM lasts SESSION_LATE_MIN minutes, so late work is still possible.
const SESSION_END_HM   = 1900;   // 7:00 PM (24-hour HHmm)
const SESSION_LATE_MIN = 120;    // logins after the end time last this many minutes

function sessionExpiry_() {
  const now = new Date();
  const p = Utilities.formatDate(now, ATTENDANCE_TZ, 'HH:mm:ss').split(':').map(Number);
  const nowSec = p[0] * 3600 + p[1] * 60 + p[2];
  const endSec = Math.floor(SESSION_END_HM / 100) * 3600 + (SESSION_END_HM % 100) * 60;
  if (nowSec >= endSec) return now.getTime() + SESSION_LATE_MIN * 60000;
  return now.getTime() + (endSec - nowSec) * 1000;
}

function sessionSave_(token, session) {
  const props = PropertiesService.getScriptProperties();
  try {                                                   // remove old finished sessions (keeps the store small)
    const all = props.getProperties();
    const nowMs = Date.now();
    Object.keys(all).forEach(function (k) {
      if (k.indexOf('sess_') !== 0) return;
      try { if (JSON.parse(all[k]).exp < nowMs) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
    });
  } catch (e) { Logger.log('sessionSave_ cleanup: ' + e); }
  const raw = JSON.stringify(session);
  props.setProperty('sess_' + token, raw);
  const ttl = Math.max(1, Math.min(21600, Math.floor((session.exp - Date.now()) / 1000)));
  CacheService.getScriptCache().put(token, raw, ttl);
}

function sessionEnd_(token) {
  if (!token) return;
  try { CacheService.getScriptCache().remove(token); } catch (e) { /* ignore */ }
  try { PropertiesService.getScriptProperties().deleteProperty('sess_' + token); } catch (e) { /* ignore */ }
}

function verifySession_(token) {
  token = token || '';
  const cache = CacheService.getScriptCache();
  let raw = cache.get(token);
  let fromCache = !!raw;
  if (!raw && token) {                                    // cache lost it (6-hour limit or eviction): use the saved copy
    raw = PropertiesService.getScriptProperties().getProperty('sess_' + token);
  }
  if (!raw) throw new Error('Session expired. Please log in again.');

  const session = JSON.parse(raw);
  if (session.exp && Date.now() > session.exp) {
    sessionEnd_(token);
    throw new Error('Session expired. Please log in again.');
  }
  if (!fromCache && session.exp) {                        // put it back in the fast cache
    const ttl = Math.max(1, Math.min(21600, Math.floor((session.exp - Date.now()) / 1000)));
    try { cache.put(token, raw, ttl); } catch (e) { /* ignore */ }
  }
  return session;
}

function assertCanAccess_(session, personKey) {
  if (isAdminRole_(session.role)) return;
  if (session.personKey !== personKey) {
    throw new Error('You do not have permission to access this crew member\'s sheet.');
  }
}

function assertIsAdmin_(session) {
  if (!isAdminRole_(session.role)) {
    throw new Error('Only Admins/Managers can perform this action.');
  }
}

// ============================================================
// TASK ENGINE
// ============================================================
function rowToTask_(row, index, key, config) {
  return {
    rowNumber: index + DATA_START_ROW,
    TimestampRaw: row[0] ? new Date(row[0]).getTime() : 0,
    Timestamp:    row[0] ? formatDate(row[0]) : '',
    TaskName:     row[1] || '',
    AssignedTo:   row[2] || '',
    Description:  row[3] || '',
    Deadline:     row[4] ? formatDate(row[4]) : '',
    Priority:     row[5] || 'Medium',
    Status:       row[6] || 'Pending',
    DoneDate:     row[7] ? formatDate(row[7]) : '',
    WorkProgress: (row[8] !== '' && row[8] !== null && row[8] !== undefined) ? row[8] : 0,
    Remark:       row[9] || '',
    FollowUp:     row[10] ? String(formatDate(row[10])).slice(0, 10) : '',   // K  Next follow-up
    LastUpdate:   row[11] ? String(formatDate(row[11])).slice(0, 10) : '',   // L  Last update
    Notes:        row[12] ? String(row[12]) : '',                            // M  Notes log
    Owner:        config.name,
    OwnerKey:     key,
    SpreadsheetId: config.id,
    SheetName:    config.sheetName
  };
}

function withRetry_(fn, maxAttempts) {
  maxAttempts = maxAttempts || 3;
  let lastErr;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      if (attempt < maxAttempts) Utilities.sleep(300 * attempt);
    }
  }
  throw lastErr;
}

// Reads every employee's task sheet (slow: opens each spreadsheet) and refreshes the cache.
// Returns { tasks, errors }. Also used by the 5-minute warm-up trigger so users almost never wait for it.
function loadAllTasksFresh_() {
  const allTasks = [];
  const errors = [];
  const validConfigs = Object.entries(SHEETS_CONFIG).filter(([key, config]) =>
    config.id && config.id.length >= 20 && !config.id.includes('PASTE'));

  for (const [key, config] of validConfigs) {
    try {
      const ss = withRetry_(() => SpreadsheetApp.openById(config.id), 3);
      const sheet = ss.getSheetByName(config.sheetName);
      if (!sheet) { errors.push(config.name + ': tab not found'); continue; }

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < DATA_START_ROW || lastCol === 0) continue;

      const numCols = Math.min(lastCol, 13);
      const totalDataRows = lastRow - DATA_START_ROW + 1;
      const data = withRetry_(() => sheet.getRange(DATA_START_ROW, 1, totalDataRows, numCols).getValues(), 3);
      data.forEach((row, index) => {
        if (!row[1] && !row[2]) return;
        allTasks.push(rowToTask_(row, index, key, config));
      });
    } catch (error) {
      errors.push(config.name + ': ' + error.message);
      Logger.log('Error: ' + config.name + ' - ' + error);
    }
  }
  allTasks.sort((a, b) => b.TimestampRaw - a.TimestampRaw);
  const clean = sanitizeForClient_(allTasks);
  if (errors.length === 0) cachePutBig_(TASKS_CACHE_KEY, clean, TASKS_CACHE_SECONDS);
  else Logger.log('Task load finished with errors, not cached: ' + errors.join(' | '));
  return { tasks: clean, errors: errors };
}

function getAllTasks(token, forceRefresh) {
  const session = verifySession_(token);
  if (!isAdminRole_(session.role)) return getTasksByPerson(token, session.personKey);

  if (!forceRefresh) {
    const hit = cacheGetBig_(TASKS_CACHE_KEY);
    if (hit) return hit;
  }

  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(20000);
  try {
    if (gotLock && !forceRefresh) {                       // someone else may have just refreshed it
      const hit2 = cacheGetBig_(TASKS_CACHE_KEY);
      if (hit2) return hit2;
    }
    return loadAllTasksFresh_().tasks;
  } finally {
    if (gotLock) lock.releaseLock();
  }
}

// Time trigger (every 5 min): keeps the task cache warm so the app opens instantly.
function warmTaskCache() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try { loadAllTasksFresh_(); } finally { lock.releaseLock(); }
}

// RUN THIS ONCE from the editor
function createCacheWarmTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'warmTaskCache') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('warmTaskCache').timeBased().everyMinutes(5).create();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'warmEmployeeCaches') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('warmEmployeeCaches').timeBased().everyMinutes(10).create();
  Logger.log('Created: warmTaskCache every 5 min + warmEmployeeCaches every 10 min.');
}

function getTasksByPerson(token, personKey, force) {
  const session = verifySession_(token);
  assertCanAccess_(session, personKey);

  const config = SHEETS_CONFIG[personKey];
  if (!config) throw new Error('Person not found: ' + personKey);

  if (!force) {                                           // fast path: the warm all-tasks cache already has it
    const hit = cacheGetBig_(TASKS_CACHE_KEY);
    if (hit) return hit.filter(function (t) { return t.OwnerKey === personKey; });
  }

  const ss = SpreadsheetApp.openById(config.id);
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) throw new Error('Sheet tab not found: ' + config.sheetName);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < DATA_START_ROW || lastCol === 0) return [];

  const numCols = Math.min(lastCol, 13);
  const totalDataRows = lastRow - DATA_START_ROW + 1;
  const data = sheet.getRange(DATA_START_ROW, 1, totalDataRows, numCols).getValues();

  const tasks = [];
  data.forEach((row, index) => {
    if (!row[1] && !row[2]) return;
    tasks.push(rowToTask_(row, index, personKey, config));
  });
  tasks.sort((a, b) => b.TimestampRaw - a.TimestampRaw);
  return tasks;
}

// Writes one task row (A-J as before + K Next follow-up, L Last update, M Notes). Used by addTaskToPerson and the recurring engine.
function appendTaskRow_(sheet, now, taskData) {
  const fu = taskData.followUp ? new Date(taskData.followUp + 'T00:00:00') : '';
  sheet.appendRow([
    now, taskData.taskName || '', taskData.assignedTo || '', taskData.description || '',
    taskData.deadline || '', taskData.priority || 'Medium', 'Pending', '', 0, '',
    (fu && !isNaN(fu.getTime())) ? fu : '', now, taskData.notes || ''
  ]);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  sheet.getRange(lastRow, 5).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(lastRow, 11, 1, 2).setNumberFormat('dd/mm/yyyy');
  try {                                              // label the 3 new columns once
    const h = sheet.getRange(DATA_START_ROW - 1, 11, 1, 3);
    if (!String(h.getValues()[0].join('')).trim()) {
      h.setValues([['Next Follow-up', 'Last Update', 'Notes']]);
      sheet.getRange(DATA_START_ROW - 1, 10).copyFormatToRange(sheet, 11, 13, DATA_START_ROW - 1, DATA_START_ROW - 1);
    }
  } catch (e) { /* header labels are cosmetic */ }
  return lastRow;
}

// Employees are VIEW-ONLY — only Admin/Manager roles may add/update/delete tasks.
function addTaskToPerson(token, personKey, taskData) {
  const session = verifySession_(token);
  assertIsAdmin_(session);
  assertCanAccess_(session, personKey);

  const config = SHEETS_CONFIG[personKey];
  if (!config) throw new Error('Person not found: ' + personKey);

  const ss = SpreadsheetApp.openById(config.id);
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) throw new Error('Sheet not found');

  const now = new Date();
  appendTaskRow_(sheet, now, taskData);

  // === WHATSAPP REMINDERS === save the reminder plan chosen in the "New task" form
  saveReminderPlan_(personKey, taskData, now.getTime());

  // Email the employee that a task was just assigned to them.
  sendTaskAssignedEmail_(personKey, taskData);

  // === WHATSAPP CLOUD API === send to employee + admin whenever a task is added
  sendTaskAssignedWhatsAppCloud_(personKey, taskData);

  // In-portal notification scoped ONLY to this personKey.
  createNotification_(personKey, 'New task assigned', buildTaskAssignedNotifText_(taskData));

  // Broadcast to every Admin/Manager account.
  createNotification_(ADMIN_BROADCAST_KEY, 'Task assigned to ' + config.name,
    buildTaskAssignedNotifText_(taskData));

  cacheRemoveBig_(TASKS_CACHE_KEY);

  return { success: true, message: 'Task added to ' + config.name };
}

function updateTask(token, personKey, rowNumber, updateData) {
  const session = verifySession_(token);
  assertIsAdmin_(session);
  assertCanAccess_(session, personKey);

  const config = SHEETS_CONFIG[personKey];
  if (!config) throw new Error('Person not found: ' + personKey);

  const ss = SpreadsheetApp.openById(config.id);
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) throw new Error('Sheet not found');

  if (rowNumber < DATA_START_ROW || rowNumber > sheet.getLastRow()) throw new Error('Invalid row number: ' + rowNumber);

  if (updateData.status !== undefined) sheet.getRange(rowNumber, 7).setValue(updateData.status);
  if (updateData.status === 'Done') {
    sheet.getRange(rowNumber, 8).setValue(new Date());
    sheet.getRange(rowNumber, 8).setNumberFormat('dd/mm/yyyy');
  }
  if (updateData.workProgress !== undefined) sheet.getRange(rowNumber, 9).setValue(updateData.workProgress);
  if (updateData.remark !== undefined) sheet.getRange(rowNumber, 10).setValue(updateData.remark);
  if (updateData.priority !== undefined) sheet.getRange(rowNumber, 6).setValue(updateData.priority);
  // follow-up date ('' clears it), a dated note appended to the notes log, and Last Update = today
  if (updateData.followUp !== undefined) {
    const fu = updateData.followUp ? new Date(updateData.followUp + 'T00:00:00') : '';
    sheet.getRange(rowNumber, 11).setValue((fu && !isNaN(fu.getTime())) ? fu : '').setNumberFormat('dd/mm/yyyy');
  }
  if (updateData.note && String(updateData.note).trim()) {
    const cell = sheet.getRange(rowNumber, 13);
    const line = Utilities.formatDate(new Date(), ATTENDANCE_TZ, 'dd/MM/yyyy') + ': ' + String(updateData.note).trim();
    const prev = String(cell.getValue() || '');
    cell.setValue(prev ? prev + '\n' + line : line);
  }
  sheet.getRange(rowNumber, 12).setValue(new Date()).setNumberFormat('dd/mm/yyyy');
  cacheRemoveBig_(TASKS_CACHE_KEY);

  return { success: true, message: 'Task updated successfully' };
}

function deleteTask(token, personKey, rowNumber) {
  const session = verifySession_(token);
  assertIsAdmin_(session);
  assertCanAccess_(session, personKey);

  const config = SHEETS_CONFIG[personKey];
  if (!config) throw new Error('Person not found: ' + personKey);

  const ss = SpreadsheetApp.openById(config.id);
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) throw new Error('Sheet not found');

  if (rowNumber < DATA_START_ROW || rowNumber > sheet.getLastRow()) throw new Error('Invalid row number: ' + rowNumber);
  sheet.deleteRow(rowNumber);
  cacheRemoveBig_(TASKS_CACHE_KEY);
  return { success: true, message: 'Task deleted' };
}

function getPeopleList(token) {
  const session = verifySession_(token);
  if (isAdminRole_(session.role)) {
    return Object.keys(SHEETS_CONFIG).map(key => ({ key: key, name: SHEETS_CONFIG[key].name }));
  }
  const config = SHEETS_CONFIG[session.personKey];
  return config ? [{ key: session.personKey, name: config.name }] : [];
}

// ============================================================
// APPS TAB (Admin/Manager only)
// ============================================================
function getApps(token) {
  const session = verifySession_(token);
  assertIsAdmin_(session);

  const sheet = getCentralSheet_(APPS_SHEET);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const apps = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = (row[0] || '').toString().trim();
    const url = (row[1] || '').toString().trim();
    if (!name || !url) continue;

    apps.push({
      name: name,
      url: url,
      description: (row[2] || '').toString().trim(),
      icon: (row[3] || '').toString().trim() || '🔗'
    });
  }
  return sanitizeForClient_(apps);
}

// ============================================================
// IN-PORTAL + BROWSER NOTIFICATIONS
// ============================================================
function getOrCreateNotificationsSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = getCentralSheet_(NOTIFICATIONS_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(NOTIFICATIONS_SHEET);
  sheet.appendRow(['Id', 'Timestamp', 'PersonKey', 'Title', 'Message', 'Read']);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  return sheet;
}

function createNotification_(personKey, title, message) {
  try {
    if (!personKey) return;
    const sheet = getOrCreateNotificationsSheet_();
    const id = Date.now() + '_' + Math.floor(Math.random() * 1000);
    sheet.appendRow([id, new Date(), personKey, title || '', message || '', false]);
    CacheService.getScriptCache().remove('notif_' + personKey);
  } catch (e) {
    Logger.log('createNotification_ error: ' + e);
  }
}

function buildTaskAssignedNotifText_(taskData) {
  const parts = [];
  if (taskData.deadline) parts.push('due ' + taskData.deadline);
  if (taskData.priority) parts.push(taskData.priority + ' priority');
  const meta = parts.length ? ' (' + parts.join(', ') + ')' : '';
  return '"' + (taskData.taskName || '(untitled)') + '"' + meta;
}

function getMyNotifications(token) {
  const session = verifySession_(token);
  try {
    const scopeKey = isAdminRole_(session.role) ? ADMIN_BROADCAST_KEY : session.personKey;
    if (!scopeKey) return { notifications: [] };

    const nCache = CacheService.getScriptCache();
    const cachedN = nCache.get('notif_' + scopeKey);
    if (cachedN) { try { return JSON.parse(cachedN); } catch (e) { /* rebuild */ } }

    const sheet = getCentralSheet_(NOTIFICATIONS_SHEET);
    if (!sheet) return { notifications: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { notifications: [] };

    const startRow = Math.max(2, lastRow - 199);          // only the newest 200 rows
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 6).getValues();
    const mine = [];
    data.forEach((row, idx) => {
      const rowPersonKey = (row[2] || '').toString().trim();
      if (rowPersonKey !== scopeKey) return;
      mine.push({
        id: row[0],
        rowNumber: idx + startRow,
        timestamp: formatDate(row[1]),
        title: row[3] || '',
        message: row[4] || '',
        read: row[5] === true || row[5] === 'TRUE' || row[5] === 'true'
      });
    });

    mine.sort((a, b) => b.rowNumber - a.rowNumber);
    const outN = sanitizeForClient_({ notifications: mine.slice(0, 40) });
    try { nCache.put('notif_' + scopeKey, JSON.stringify(outN), 30); } catch (e) { /* too big: skip */ }
    return outN;
  } catch (e) {
    Logger.log('getMyNotifications error: ' + e);
    return { notifications: [] };
  }
}

function markNotificationsRead(token, ids) {
  const session = verifySession_(token);
  const scopeKey = isAdminRole_(session.role) ? ADMIN_BROADCAST_KEY : session.personKey;
  if (!scopeKey || !ids || !ids.length) return { success: true };

  const sheet = getCentralSheet_(NOTIFICATIONS_SHEET);
  if (!sheet) return { success: true };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true };

  const idSet = new Set(ids.map(String));
  const startRow = Math.max(2, lastRow - 199);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 6).getValues();
  data.forEach((row, idx) => {
    const rowPersonKey = (row[2] || '').toString().trim();
    if (rowPersonKey !== scopeKey) return;
    if (idSet.has(String(row[0]))) {
      sheet.getRange(idx + startRow, 6).setValue(true);
    }
  });
  CacheService.getScriptCache().remove('notif_' + scopeKey);

  return { success: true };
}

// ============================================================
// TEAM CHAT
// ============================================================
const CHAT_SHEET = 'Chat';

function getOrCreateChatSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = getCentralSheet_(CHAT_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(CHAT_SHEET);
  sheet.appendRow(['Timestamp', 'Username', 'FullName', 'Message']);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  return sheet;
}

function getChatMessages(token, sinceRow) {
  const session = verifySession_(token);
  try {
    sinceRow = Number(sinceRow) || 1;
    const chatCache = CacheService.getScriptCache();
    const knownLast = chatCache.get('chatLast');         // idle polls are answered from cache without opening the sheet
    if (knownLast !== null && Number(knownLast) <= sinceRow) {
      return sanitizeForClient_({ messages: [], lastRow: Number(knownLast) });
    }
    const sheet = getOrCreateChatSheet_();
    const lastRow = sheet.getLastRow();
    chatCache.put('chatLast', String(lastRow), 300);

    if (lastRow <= sinceRow) {
      return sanitizeForClient_({ messages: [], lastRow: lastRow });
    }

    const startRow = sinceRow + 1;
    const numRows = lastRow - sinceRow;
    const data = sheet.getRange(startRow, 1, numRows, 4).getValues();

    const messages = data.map((row, idx) => ({
      row: startRow + idx,
      timestamp: formatDate(row[0]),
      username: row[1] || '',
      fullName: row[2] || '',
      message: row[3] || ''
    }));

    return sanitizeForClient_({ messages: messages, lastRow: lastRow });
  } catch (e) {
    Logger.log('getChatMessages error: ' + e);
    return { messages: [], lastRow: Number(sinceRow) || 1 };
  }
}

function sendChatMessage(token, message) {
  const session = verifySession_(token);
  message = (message || '').toString().trim();
  if (!message) throw new Error('Message cannot be empty.');
  if (message.length > 2000) message = message.substring(0, 2000);

  const sheet = getOrCreateChatSheet_();
  sheet.appendRow([new Date(), session.username, session.fullName, message]);
  CacheService.getScriptCache().put('chatLast', String(sheet.getLastRow()), 300);
  return { success: true };
}

// ============================================================
// EMAIL NOTIFICATIONS
// ============================================================
function getEmployeeEmail_(personKey) {
  const usersSheet = getCentralSheet_(USERS_SHEET);
  if (!usersSheet) return null;
  const data = usersSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][4] || '').toString().trim() === personKey) {
      const email = (data[i][5] || '').toString().trim();
      return email || null;
    }
  }
  return null;
}

function sendTaskAssignedEmail_(personKey, taskData) {
  try {
    const config = SHEETS_CONFIG[personKey];
    const email = getEmployeeEmail_(personKey);
    if (!email) { Logger.log('No email on file for ' + (config ? config.name : personKey) + ' — skipping notification.'); return; }

    const subject = 'New Task Assigned: ' + (taskData.taskName || '(untitled)');
    const body =
      'Hi ' + (config ? config.name : '') + ',\n\n' +
      'A new task has been assigned to you:\n\n' +
      'Task: ' + (taskData.taskName || '') + '\n' +
      'Assigned To: ' + (taskData.assignedTo || '') + '\n' +
      'Description: ' + (taskData.description || '') + '\n' +
      'Deadline: ' + (taskData.deadline || '') + '\n' +
      'Priority: ' + (taskData.priority || 'Medium') + '\n\n' +
      'Please log in to the Task Portal to view details.\n\n' +
      '— Employee Task Management System';

    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('sendTaskAssignedEmail_ error: ' + e);
  }
}

function getTasksByPersonRaw_(personKey) {
  const config = SHEETS_CONFIG[personKey];
  if (!config) return [];
  const ss = SpreadsheetApp.openById(config.id);
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < DATA_START_ROW || lastCol === 0) return [];
  const numCols = Math.min(lastCol, 13);
  const totalDataRows = lastRow - DATA_START_ROW + 1;
  const data = sheet.getRange(DATA_START_ROW, 1, totalDataRows, numCols).getValues();
  const tasks = [];
  data.forEach((row, index) => {
    if (!row[1] && !row[2]) return;
    tasks.push(rowToTask_(row, index, personKey, config));
  });
  return tasks;
}

// ============================================================
// DAILY REMINDER JOB (email, once/day via trigger)
// ============================================================
function sendDailyTaskReminders() {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate() + 1);

  Object.keys(SHEETS_CONFIG).forEach(function (personKey) {
    try {
      const config = SHEETS_CONFIG[personKey];
      const email = getEmployeeEmail_(personKey);
      if (!email) return;

      const tasks = getTasksByPersonRaw_(personKey);
      const dueToday = [];
      const dueTomorrow = [];

      tasks.forEach(function (t) {
        if (t.Status === 'Done' || !t.Deadline) return;
        const parts = String(t.Deadline).split('/');
        if (parts.length < 3) return;
        const d = new Date(parts[2].split(' ')[0], parts[1] - 1, parts[0]);
        d.setHours(0, 0, 0, 0);
        if (d.getTime() === today.getTime()) dueToday.push(t);
        else if (d.getTime() === tomorrow.getTime()) dueTomorrow.push(t);
      });

      if (dueToday.length === 0 && dueTomorrow.length === 0) return;

      let body = 'Hi ' + config.name + ',\n\nHere is your task reminder:\n\n';
      if (dueToday.length) {
        body += 'DUE TODAY:\n';
        dueToday.forEach(t => body += '- ' + t.TaskName + ' (Priority: ' + t.Priority + ')\n');
        body += '\n';
      }
      if (dueTomorrow.length) {
        body += 'DUE TOMORROW:\n';
        dueTomorrow.forEach(t => body += '- ' + t.TaskName + ' (Priority: ' + t.Priority + ')\n');
        body += '\n';
      }
      body += 'Please log in to the Task Portal to update your progress.\n\n— Employee Task Management System';

      MailApp.sendEmail(email, 'Task Reminder: ' + (dueToday.length + dueTomorrow.length) + ' task(s) due', body);
    } catch (e) {
      Logger.log('sendDailyTaskReminders error for ' + personKey + ': ' + e);
    }
  });
}

function createDailyReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyTaskReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyTaskReminders')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

// ============================================================
// ADMIN/MANAGER: get full data for ONE employee
// ============================================================
// Cheap lookup (no sheet scanning): username of this person + the sheets linked to them in SheetLinks.
function empLinks_(personKey) {
  const config = SHEETS_CONFIG[personKey];
  let username = null, fullName = config ? config.name : personKey;
  const usersSheet = getCentralSheet_(USERS_SHEET);
  if (usersSheet) {
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][4] || '').toString().trim() === personKey) {
        username = (data[i][0] || '').toString().trim();
        fullName = data[i][2] || fullName;
        break;
      }
    }
  }
  const personalSheets = [];
  if (username) {
    const normUsername = normalizeName_(username);
    const normFullName = normalizeName_(fullName);
    const linksSheet = getCentralSheet_(SHEETLINKS_SHEET);
    if (linksSheet) {
      const data = linksSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const owner = (row[0] || '').toString().trim();
        if (!owner) continue;
        const normOwner = normalizeName_(owner);
        if (normOwner === normUsername || (normFullName && normOwner === normFullName)) {
          personalSheets.push({ owner: owner, label: row[1] || 'Untitled Sheet', url: row[2] || '', tabName: row[3] || '' });
        }
      }
    }
  }
  return { username: username, fullName: fullName, personalSheets: personalSheets };
}

// Instant call for the dashboard: the employee's sheet list (cached 5 min), so the cards appear right after the click.
function getEmployeeLinks(token, personKey) {
  const session = verifySession_(token);
  assertCanAccess_(session, personKey);
  if (!SHEETS_CONFIG[personKey]) throw new Error('Person not found: ' + personKey);
  const ck = 'links_' + personKey;
  const hit = cacheGetBig_(ck);
  if (hit) return hit;
  const l = empLinks_(personKey);
  const out = sanitizeForClient_({ personKey: personKey, username: l.username, personalSheets: l.personalSheets });
  cachePutBig_(ck, out, 300);
  return out;
}

// The slow part: scans every personal sheet of one employee for today's pending work, and caches it.
function loadEmployeeSheets_(personKey) {
  const l = empLinks_(personKey);
  let pendingWork = [];
  l.personalSheets.forEach(function (link) {
    pendingWork = pendingWork.concat(scanSheetForPendingToday_(link, l.username, l.fullName));
  });
  pendingWork.sort((a, b) => new Date(a.dueDate) - new Date(b.dueDate));
  const out = { username: l.username, personalSheets: l.personalSheets, pendingWork: sanitizeForClient_(pendingWork) };
  cachePutBig_(pendCacheKey_('emp_' + personKey), out, 900);
  return out;
}

function getEmployeeFullData(token, personKey, skipTasks, force) {
  const session = verifySession_(token);
  assertCanAccess_(session, personKey);
  const config = SHEETS_CONFIG[personKey];
  if (!config) throw new Error('Person not found: ' + personKey);

  let tasks = [];
  let tasksError = null;
  try {
    if (!skipTasks) tasks = getTasksByPerson(token, personKey, force);
  } catch (e) {
    tasksError = e.message || e.toString();
    Logger.log('getEmployeeFullData: could not load tasks for ' + personKey + ': ' + tasksError);
  }

  let part = force ? null : cacheGetBig_(pendCacheKey_('emp_' + personKey));
  if (!part) part = loadEmployeeSheets_(personKey);

  return sanitizeForClient_({
    personKey: personKey,
    name: config.name,
    username: part.username,
    tasks: tasks,
    personalSheets: part.personalSheets,
    pendingWork: part.pendingWork,
    tasksError: tasksError
  });
}

// Time trigger (every 10 min): pre-scans every employee's sheets so a click on an employee is instant.
function warmEmployeeCaches() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    Object.keys(SHEETS_CONFIG).forEach(function (pk) {
      try { loadEmployeeSheets_(pk); } catch (e) { Logger.log('warmEmployeeCaches ' + pk + ': ' + e); }
    });
  } finally { lock.releaseLock(); }
}

// ============================================================
// PERSONAL SHEETS
// ============================================================
function getUserSheets(token, forUsername) {
  const session = verifySession_(token);
  try {
    const sheet = getCentralSheet_(SHEETLINKS_SHEET);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const links = [];

    const scopeUsername = normalizeName_(forUsername || '');
    const sessionUsernameNorm = normalizeName_(session.username);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const owner = (row[0] || '').toString().trim();
      if (!owner) continue;
      const normOwner = normalizeName_(owner);

      if (scopeUsername) {
        if (normOwner !== scopeUsername) continue;
      } else if (!isAdminRole_(session.role) && normOwner !== sessionUsernameNorm) {
        continue;
      }

      links.push({ owner: owner, label: row[1] || 'Untitled Sheet', url: row[2] || '', tabName: row[3] || '' });
    }
    return links;
  } catch (e) {
    Logger.log('getUserSheets error: ' + e);
    return [];
  }
}

function extractSheetId_(url) {
  const match = (url || '').match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

function findColumn_(headers, keywords, exclude) {
  exclude = exclude || [];
  const norm = (h) => (h || '').toString().trim().toLowerCase();

  for (const k of keywords) {
    for (let i = 0; i < headers.length; i++) {
      if (norm(headers[i]) === k) return i;
    }
  }

  for (const k of keywords) {
    for (let i = 0; i < headers.length; i++) {
      const h = norm(headers[i]);
      if (exclude.some(x => h.includes(x))) continue;
      if (h.includes(k)) return i;
    }
  }

  return -1;
}

function findHeaderRow_(data, maxRowsToScan) {
  const limit = Math.min(maxRowsToScan || 15, data.length);
  let bestIndex = 0;
  let bestScore = 0;

  for (let r = 0; r < limit; r++) {
    const row = data[r];
    if (!row) continue;

    let score = 0;
    Object.values(COLUMN_KEYWORDS_).forEach(group => {
      if (findColumn_(row, group.keywords, group.exclude) !== -1) score++;
    });

    const hasDateLike = findColumn_(row, COLUMN_KEYWORDS_.date.keywords, COLUMN_KEYWORDS_.date.exclude) !== -1;
    if (hasDateLike && score > bestScore) {
      bestScore = score;
      bestIndex = r;
    }
  }

  return { headerRowIndex: bestIndex, headers: data[bestIndex] || data[0] };
}

function findAllColumns_(headers, exactText) {
  const target = (exactText || '').toString().trim().toLowerCase();
  const found = [];
  for (let i = 0; i < headers.length; i++) {
    const h = (headers[i] || '').toString().trim().toLowerCase();
    if (h === target) found.push(i);
  }
  return found;
}

function findColumnInRange_(headers, start, end, keywords) {
  for (let i = start; i <= end; i++) {
    const h = (headers[i] || '').toString().trim().toLowerCase();
    if (keywords.some(k => h === k)) return i;
  }
  for (let i = start; i <= end; i++) {
    const h = (headers[i] || '').toString().trim().toLowerCase();
    if (keywords.some(k => h.includes(k))) return i;
  }
  return -1;
}

function scanMultiStagePendingToday_(link, username, fullName, data, headers, headerRowIndex, firstDataRow) {
  const statusCols = findAllColumns_(headers, 'status');
  if (statusCols.length < 2) return null;

  const stages = statusCols.map(function (sc, idx) {
    const blockStart = idx === 0 ? 0 : statusCols[idx - 1] + 1;
    const labelEnd = idx === statusCols.length - 1 ? headers.length - 1 : statusCols[idx + 1] - 1;
    return {
      blockStart: blockStart,
      labelEnd: labelEnd,
      statusCol: sc,
      plannedCol: findColumnInRange_(headers, blockStart, sc, ['planned']),
      actualCol: findColumnInRange_(headers, blockStart, sc, ['actual'])
    };
  });

  const labelScanStart = Math.max(0, headerRowIndex - 8);
  stages.forEach(function (stage, idx) {
    let bestLabel = '';
    for (let r = labelScanStart; r < headerRowIndex; r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = stage.blockStart; c <= stage.labelEnd; c++) {
        const val = (row[c] || '').toString().trim();
        if (val && val.length > bestLabel.length) bestLabel = val;
      }
    }
    stage.label = bestLabel || ('Step ' + (idx + 1));
  });

  const firstAnchor = stages[0].plannedCol !== -1 ? stages[0].plannedCol : stages[0].statusCol;
  const preambleEnd = Math.max(0, firstAnchor - 1);

  let identityCol = findColumnInRange_(headers, 0, preambleEnd, ['vessel', 'supplier', 'party', 'client', 'name']);
  if (identityCol === -1) identityCol = Math.min(1, headers.length - 1);

  const fallbackDateCol = findColumnInRange_(headers, 0, preambleEnd, ['timestamp', 'date']);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const results = [];

  for (let i = firstDataRow; i < data.length; i++) {
    const row = data[i];
    const identity = (row[identityCol] || '').toString().trim();
    if (!identity) continue;

    for (let s = 0; s < stages.length; s++) {
      const stage = stages[s];
      const rawStatus = (row[stage.statusCol] || '').toString().trim().toLowerCase();
      const isExplicitlyNotDone = /\bnot\b/.test(rawStatus) ||
                                   rawStatus === 'pending' ||
                                   rawStatus === 'no' ||
                                   rawStatus.includes('incomplete');
      const isDone = !isExplicitlyNotDone && DONE_KEYWORDS.some(k => rawStatus.includes(k));
      if (isDone) continue;

      let refDate = null;
      const rawRef = (stage.plannedCol !== -1 ? row[stage.plannedCol] : null) ||
                     (fallbackDateCol !== -1 ? row[fallbackDateCol] : null);
      if (rawRef) {
        const d = rawRef instanceof Date ? rawRef : new Date(rawRef);
        if (!isNaN(d.getTime()) && d.getFullYear() >= 2000) { d.setHours(0, 0, 0, 0); refDate = d; }
      }

      if (refDate && refDate.getTime() > today.getTime()) break;

      results.push({
        sheetLabel: link.label,
        owner: link.owner,
        task: identity,
        description: 'Stuck at: ' + stage.label,
        priority: '',
        status: row[stage.statusCol] || 'Pending',
        dueDate: refDate ? Utilities.formatDate(refDate, Session.getScriptTimeZone(), 'MMM dd, yyyy') : '—',
        isOverdue: !!(refDate && refDate.getTime() < today.getTime()),
        sheetUrl: link.url,
        rowIndex: i + 1,
        statusCol: stage.statusCol + 1,
        actualCol: stage.actualCol !== -1 ? stage.actualCol + 1 : null
      });
      break;
    }
  }

  return results;
}

function scanSheetForPendingToday_(link, username, fullName) {
  const results = [];
  const sheetId = extractSheetId_(link.url);
  if (!sheetId) return results;

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const tab = link.tabName ? getTabFuzzy_(ss, link.tabName) : pickDefaultTab_(ss);
    if (!tab) return results;

    const data = tab.getDataRange().getValues();
    if (data.length < 2) return results;

    const headerInfo = findHeaderRow_(data, 15);
    const headers = headerInfo.headers;
    const firstDataRow = headerInfo.headerRowIndex + 1;

    const multiStageResults = scanMultiStagePendingToday_(
      link, username, fullName, data, headers, headerInfo.headerRowIndex, firstDataRow
    );
    if (multiStageResults !== null) {
      // Tag each row with WHERE it came from, so the frontend can write an
      // inline status edit straight back to the correct cell.
      multiStageResults.forEach(function (r) {
        r.spreadsheetId = sheetId;
        r.tabName = tab.getName();
      });
      return multiStageResults;
    }

    const dateCol = findColumn_(headers, COLUMN_KEYWORDS_.date.keywords, COLUMN_KEYWORDS_.date.exclude);
    const statusCol = findColumn_(headers, COLUMN_KEYWORDS_.status.keywords, COLUMN_KEYWORDS_.status.exclude);
    const taskCol = findColumn_(headers, COLUMN_KEYWORDS_.task.keywords, COLUMN_KEYWORDS_.task.exclude);
    const assignedCol = findColumn_(headers, COLUMN_KEYWORDS_.assigned.keywords, COLUMN_KEYWORDS_.assigned.exclude);
    const descCol = findColumn_(headers, COLUMN_KEYWORDS_.desc.keywords, COLUMN_KEYWORDS_.desc.exclude);
    const priorityCol = findColumn_(headers, COLUMN_KEYWORDS_.priority.keywords, COLUMN_KEYWORDS_.priority.exclude);

    if (dateCol === -1) return results;
    if (taskCol === -1 && assignedCol === -1) return results;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const normUsername = normalizeName_(username);
    const normFullName = normalizeName_(fullName);

    for (let i = firstDataRow; i < data.length; i++) {
      const row = data[i];
      const rawDate = row[dateCol];
      if (!rawDate) continue;
      const rowDate = rawDate instanceof Date ? rawDate : new Date(rawDate);
      if (isNaN(rowDate.getTime())) continue;
      rowDate.setHours(0, 0, 0, 0);
      if (rowDate.getFullYear() < 2000) continue;
      if (rowDate.getTime() > today.getTime()) continue;

      if (statusCol !== -1) {
        const rawStatus = (row[statusCol] || '').toString().trim().toLowerCase();
        const isExplicitlyNotDone = /\bnot\b/.test(rawStatus) ||
                                     rawStatus === 'pending' ||
                                     rawStatus === 'no' ||
                                     rawStatus.includes('incomplete');
        const isDone = !isExplicitlyNotDone && DONE_KEYWORDS.some(k => rawStatus.includes(k));
        if (isDone) continue;
      }

      if (assignedCol !== -1) {
        const assignedVal = normalizeName_(row[assignedCol]);
        const matchesUser = assignedVal === normUsername ||
                             assignedVal === normFullName ||
                             (normUsername && assignedVal.includes(normUsername)) ||
                             (normFullName && assignedVal.includes(normFullName));
        if (!matchesUser) continue;
      }

      results.push({
        sheetLabel: link.label,
        owner: link.owner,
        task: taskCol !== -1 ? (row[taskCol] || 'Untitled item') : 'Untitled item',
        description: descCol !== -1 ? (row[descCol] || '') : '',
        priority: priorityCol !== -1 ? (row[priorityCol] || '') : '',
        status: statusCol !== -1 ? (row[statusCol] || 'Pending') : 'Pending',
        dueDate: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'MMM dd, yyyy'),
        isOverdue: rowDate.getTime() < today.getTime(),
        sheetUrl: link.url,
        rowIndex: i + 1,
        statusCol: statusCol !== -1 ? statusCol + 1 : null,
        actualCol: null
      });
    }

    results.forEach(function (r) {
      r.spreadsheetId = sheetId;
      r.tabName = tab.getName();
    });
  } catch (e) {
    Logger.log('scanSheetForPendingToday_ error for ' + link.label + ': ' + e);
  }
  return results;
}

function getTodayPendingWork(token, force) {
  const session = verifySession_(token);
  try {
    const meKey = pendCacheKey_('me_' + normalizeName_(session.username));
    if (!force) { const hit = cacheGetBig_(meKey); if (hit) return hit; }
    const links = getUserSheets(token);
    let allPending = [];
    links.forEach(link => {
      try {
        allPending = allPending.concat(scanSheetForPendingToday_(link, session.username, session.fullName));
      } catch (e) {
        Logger.log('getTodayPendingWork: error scanning ' + link.label + ': ' + e);
      }
    });
    allPending.sort((a, b) => new Date(a.dueDate) - new Date(b.dueDate));
    const outPending = sanitizeForClient_(allPending);
    cachePutBig_(meKey, outPending, 600);
    return outPending;
  } catch (e) {
    Logger.log('getTodayPendingWork error: ' + e);
    return [];
  }
}

// ============================================================
// INLINE EDIT for Pending & Overdue Work (Delegation/Checklist/FMS rows)
// Admin/Manager only, same permission level as updateTask/deleteTask.
// ============================================================
function updatePendingItemStatus(token, item) {
  const session = verifySession_(token);
  assertIsAdmin_(session);

  if (!item || !item.spreadsheetId || !item.tabName || !item.rowIndex || !item.statusCol) {
    throw new Error('This row is missing the info needed to edit it in place — open the sheet instead.');
  }

  const ss = SpreadsheetApp.openById(item.spreadsheetId);
  const tab = getTabFuzzy_(ss, item.tabName) || ss.getSheetByName(item.tabName);
  if (!tab) throw new Error('Could not find tab "' + item.tabName + '" to update.');

  const newStatus = (item.newStatus || '').toString();
  tab.getRange(item.rowIndex, item.statusCol).setValue(newStatus);

  if (item.actualCol && DONE_KEYWORDS.some(k => newStatus.toLowerCase().includes(k))) {
    tab.getRange(item.rowIndex, item.actualCol).setValue(new Date());
  }

  bumpPendCache_();                       // pending lists must show the change
  return { success: true, message: 'Updated.' };
}

// ============================================================
// DEBUG (Admin/Manager only)
// ============================================================
function debugSheets(token) {
  const session = verifySession_(token);
  if (!isAdminRole_(session.role)) throw new Error('Admin/Manager only.');

  for (const [key, config] of Object.entries(SHEETS_CONFIG)) {
    try {
      const ss = SpreadsheetApp.openById(config.id);
      const sheet = ss.getSheetByName(config.sheetName);
      if (!sheet) {
        Logger.log(config.name + ': TAB NOT FOUND - available: ' + ss.getSheets().map(s => s.getName()).join(', '));
        continue;
      }
      Logger.log(config.name + ': lastRow=' + sheet.getLastRow());
    } catch (e) {
      Logger.log(config.name + ': ERROR - ' + e);
    }
  }
}

function debugAllTasksReport(token) {
  const session = verifySession_(token);
  assertIsAdmin_(session);

  const report = [];

  Object.entries(SHEETS_CONFIG).forEach(function ([key, config]) {
    const entry = {
      personKey: key,
      name: config.name,
      configuredSpreadsheetId: config.id,
      configuredTabName: config.sheetName,
      dataStartRowConfigured: DATA_START_ROW
    };

    if (!config.id || config.id.length < 20 || config.id.includes('PASTE')) {
      entry.error = 'Spreadsheet ID looks invalid, blank, or is still a placeholder — update SHEETS_CONFIG. This is why this person is silently excluded from the combined view.';
      report.push(entry);
      return;
    }

    let ss;
    try {
      ss = SpreadsheetApp.openById(config.id);
      entry.spreadsheetOpened = true;
    } catch (e) {
      entry.spreadsheetOpened = false;
      entry.error = 'Could not open this spreadsheet: ' + e.toString() +
        ' — most likely cause: the Google account that deployed this web app ' +
        '("Execute as: Me") does not have Editor access to this spreadsheet ID. ' +
        'Share the sheet with that account and try again.';
      report.push(entry);
      return;
    }

    entry.allTabsInSpreadsheet = ss.getSheets().map(s => s.getName());

    const sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) {
      entry.error = 'Tab "' + config.sheetName + '" was not found in this spreadsheet. ' +
        'Available tabs are listed in allTabsInSpreadsheet above — update SHEETS_CONFIG.sheetName to match exactly.';
      report.push(entry);
      return;
    }

    entry.tabFound = true;
    entry.lastRow = sheet.getLastRow();
    entry.lastColumn = sheet.getLastColumn();

    if (entry.lastRow < DATA_START_ROW) {
      entry.warning = 'lastRow (' + entry.lastRow + ') is below DATA_START_ROW (' + DATA_START_ROW +
        ') — this sheet will contribute ZERO tasks. Either add data starting at row ' + DATA_START_ROW +
        ', or change DATA_START_ROW at the top of Code.gs if your real data starts on a different row.';
      report.push(entry);
      return;
    }

    try {
      const numCols = Math.min(entry.lastColumn, 10);
      const sampleRowCount = Math.min(3, entry.lastRow - DATA_START_ROW + 1);
      const sample = sheet.getRange(DATA_START_ROW, 1, sampleRowCount, numCols).getValues();
      entry.sampleDataRows = sample;
      entry.sampleRowsThatWouldBeSkipped = sample.filter(r => !r[1] && !r[2]).length;
    } catch (e) {
      entry.error = 'Could not read sample rows: ' + e.toString();
    }

    report.push(entry);
  });

  return sanitizeForClient_(report);
}

function debugPendingScan(token, forPersonKey) {
  let session;
  try {
    session = verifySession_(token);
  } catch (e) {
    return [{ error: 'Session error: ' + e.toString() }];
  }

  try {
    let forUsername = null;
    if (forPersonKey && SHEETS_CONFIG[forPersonKey]) {
      const usersSheet = getCentralSheet_(USERS_SHEET);
      if (usersSheet) {
        const data = usersSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if ((data[i][4] || '').toString().trim() === forPersonKey) {
            forUsername = (data[i][0] || '').toString().trim();
            break;
          }
        }
      }
    }

    const links = getUserSheets(token, forUsername);
    const report = [];

    if (links.length === 0) {
      report.push({ info: 'No SheetLinks rows found for this scope. Check the SheetLinks tab has a row whose Username matches, and that forPersonKey/username resolved correctly.' });
    }

    links.forEach(function(link) {
      const entry = { label: link.label, tabNameRequested: link.tabName, url: link.url };
      const sheetId = extractSheetId_(link.url);
      if (!sheetId) { entry.error = 'Could not extract a Sheet ID from this URL.'; report.push(entry); return; }

      try {
        const ss = SpreadsheetApp.openById(sheetId);
        entry.allTabsInSpreadsheet = ss.getSheets().map(s => s.getName());
        const tab = link.tabName ? getTabFuzzy_(ss, link.tabName) : pickDefaultTab_(ss);
        if (!tab) { entry.error = 'Tab "' + link.tabName + '" not found.'; report.push(entry); return; }
        entry.tabActuallyUsed = tab.getName();

        const data = tab.getDataRange().getValues();
        entry.totalRows = data.length - 1;
        if (data.length < 2) { entry.error = 'No data rows.'; report.push(entry); return; }

        const headerInfo = findHeaderRow_(data, 15);
        const headers = headerInfo.headers;
        entry.headerRowDetected = headerInfo.headerRowIndex + 1;
        entry.headersFound = headers;

        const statusCols = findAllColumns_(headers, 'status');
        entry.isMultiStagePipeline = statusCols.length >= 2;
        if (entry.isMultiStagePipeline) {
          entry.stageStatusColumnsFound = statusCols.map(c => headers[c] + ' (col ' + (c + 1) + ')');
        }

        const dateCol = findColumn_(headers, COLUMN_KEYWORDS_.date.keywords, COLUMN_KEYWORDS_.date.exclude);
        const statusCol = findColumn_(headers, COLUMN_KEYWORDS_.status.keywords, COLUMN_KEYWORDS_.status.exclude);
        const taskCol = findColumn_(headers, COLUMN_KEYWORDS_.task.keywords, COLUMN_KEYWORDS_.task.exclude);
        const assignedCol = findColumn_(headers, COLUMN_KEYWORDS_.assigned.keywords, COLUMN_KEYWORDS_.assigned.exclude);

        entry.detectedColumns = {
          dateColumn: dateCol !== -1 ? headers[dateCol] : 'NOT FOUND',
          statusColumn: statusCol !== -1 ? headers[statusCol] : 'NOT FOUND',
          taskColumn: taskCol !== -1 ? headers[taskCol] : 'NOT FOUND',
          assignedToColumn: assignedCol !== -1 ? headers[assignedCol] : 'NOT FOUND'
        };

        const scanResult = scanSheetForPendingToday_(link, session.username, session.fullName);
        entry.pendingItemsFound = scanResult.length;
      } catch (e) {
        entry.error = 'Exception: ' + e.toString();
      }
      report.push(entry);
    });

    return sanitizeForClient_(report);
  } catch (e) {
    Logger.log('debugPendingScan top-level error: ' + e);
    return [{ error: 'debugPendingScan crashed before building the report: ' + e.toString() }];
  }
}

function debugNameLookup(token, nameToCheck) {
  const session = verifySession_(token);
  assertIsAdmin_(session);
  nameToCheck = (nameToCheck || '').toString().trim();
  if (!nameToCheck) throw new Error('Provide a name to check.');

  const report = { nameToCheck: nameToCheck, usersTabMatches: [], sheetLinksRows: [] };
  const normTarget = normalizeName_(nameToCheck);

  const usersSheet = getCentralSheet_(USERS_SHEET);
  if (usersSheet) {
    const usersData = usersSheet.getDataRange().getValues();
    usersData.forEach((row, i) => {
      if (i === 0) return;
      const username = (row[0] || '').toString();
      const fullName = (row[2] || '').toString();
      if (normalizeName_(username) === normTarget || normalizeName_(fullName) === normTarget || normalizeName_(username).includes(normTarget)) {
        report.usersTabMatches.push({
          row: i + 1, username: username, fullName: fullName,
          role: row[3], personKey: row[4],
          usernameCharCodes: Array.from(username).map(c => c.charCodeAt(0))
        });
      }
    });
  }

  const linksSheet = getCentralSheet_(SHEETLINKS_SHEET);
  if (linksSheet) {
    const linksData = linksSheet.getDataRange().getValues();
    linksData.forEach((row, i) => {
      if (i === 0) return;
      const owner = (row[0] || '').toString();
      const normOwner = normalizeName_(owner);
      report.sheetLinksRows.push({
        row: i + 1, owner: owner, label: row[1], url: row[2], tabName: row[3],
        normalizedOwner: normOwner,
        wouldMatchTarget: normOwner === normTarget,
        ownerCharCodes: Array.from(owner).map(c => c.charCodeAt(0))
      });
    });
  }

  return sanitizeForClient_(report);
}

function sanitizeForClient_(value) {
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (e) {
    Logger.log('sanitizeForClient_ failed: ' + e);
    return [{ error: 'Report was built but could not be serialized for the browser: ' + e.toString() }];
  }
}

// ============================================================
// HELPER: FORMAT DATE
// ============================================================
function formatDate(dateValue) {
  if (!dateValue) return '';
  try {
    let date;
    if (dateValue instanceof Date) date = dateValue;
    else if (typeof dateValue === 'string') date = new Date(dateValue);
    else if (typeof dateValue === 'number') date = new Date(Math.round((dateValue - 25569) * 86400 * 1000));
    else return dateValue;

    if (isNaN(date.getTime())) return dateValue;

    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return day + '/' + month + '/' + year + ' ' + hours + ':' + minutes;
  } catch (error) {
    return dateValue;
  }
}

function debugSayeedLookup() {
  const nameToCheck = 'Sayeed';

  const usersSheet = getCentralSheet_(USERS_SHEET);
  if (!usersSheet) { Logger.log('Users tab not found!'); return; }
  const usersData = usersSheet.getDataRange().getValues();
  Logger.log('--- USERS TAB ROWS CONTAINING "' + nameToCheck + '" ---');
  usersData.forEach((row, i) => {
    if (i === 0) return;
    const rowText = row.join(' | ').toLowerCase();
    if (rowText.includes(nameToCheck.toLowerCase())) {
      Logger.log('Row ' + (i+1) + ': Username="' + row[0] + '" FullName="' + row[2] + '" Role="' + row[3] + '" PersonKey="' + row[4] + '"');
    }
  });

  const linksSheet = getCentralSheet_(SHEETLINKS_SHEET);
  if (!linksSheet) { Logger.log('SheetLinks tab not found!'); return; }
  const linksData = linksSheet.getDataRange().getValues();
  Logger.log('--- ALL SheetLinks ROWS (so you can see EXACT spelling in column A) ---');
  linksData.forEach((row, i) => {
    if (i === 0) return;
    Logger.log('Row ' + (i+1) + ': Username="' + row[0] + '" Label="' + row[1] + '" URL="' + row[2] + '"');
  });

  Logger.log('--- Rows that WOULD match "' + nameToCheck + '" (case-insensitive exact) ---');
  const matches = linksData.filter((row, i) => i > 0 && (row[0]||'').toString().trim().toLowerCase() === nameToCheck.toLowerCase());
  Logger.log('Matched count: ' + matches.length);
}

// ============================================================
// LIVE DUE-TODAY / DUE-TOMORROW CHECK
// ============================================================
function checkDueSoonAndNotify(token) {
  const session = verifySession_(token);
  const cache = CacheService.getScriptCache();

  // This scans every task sheet, so run it at most once every 15 minutes per user type (it dedupes anyway).
  const runKey = 'dueSoonRun_' + (isAdminRole_(session.role) ? 'ADMIN' : (session.personKey || 'none'));
  if (cache.get(runKey)) return { checked: 0, notified: 0, skipped: true };
  cache.put(runKey, '1', 900);

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate() + 1);

  function scanAndNotify(personKey, notifyKey, namePrefix) {
    const tasks = getTasksByPersonRaw_(personKey);
    let notified = 0;

    tasks.forEach(function (t) {
      if (t.Status === 'Done' || !t.Deadline) return;
      const parts = String(t.Deadline).split('/');
      if (parts.length < 3) return;
      const d = new Date(parts[2].split(' ')[0], parts[1] - 1, parts[0]);
      d.setHours(0, 0, 0, 0);

      let label = null;
      if (d.getTime() === today.getTime()) label = 'due today';
      else if (d.getTime() === tomorrow.getTime()) label = 'due tomorrow';
      if (!label) return;

      const dedupeKey = 'dueSoon_' + notifyKey + '_' + personKey + '_' + t.rowNumber + '_' + t.Deadline;
      if (cache.get(dedupeKey)) return;

      const title = 'Task ' + label;
      const message = (namePrefix ? '[' + namePrefix + '] ' : '') +
        '"' + (t.TaskName || '(untitled)') + '"' + (t.Priority ? ' (' + t.Priority + ' priority)' : '');

      createNotification_(notifyKey, title, message);
      cache.put(dedupeKey, '1', 72000);
      notified++;
    });

    return { checked: tasks.length, notified: notified };
  }

  if (isAdminRole_(session.role)) {
    let totalChecked = 0, totalNotified = 0;
    Object.keys(SHEETS_CONFIG).forEach(function (personKey) {
      const config = SHEETS_CONFIG[personKey];
      const result = scanAndNotify(personKey, ADMIN_BROADCAST_KEY, config ? config.name : personKey);
      totalChecked += result.checked;
      totalNotified += result.notified;
    });
    return { checked: totalChecked, notified: totalNotified };
  }

  if (!session.personKey) return { checked: 0, notified: 0 };
  return scanAndNotify(session.personKey, session.personKey, null);
}

// ============================================================
// WHATSAPP REMINDERS via META WHATSAPP CLOUD API
// ============================================================
// SECURITY: the access token is NO LONGER hardcoded here. Set it in
// Apps Script > Project Settings > Script Properties:
//     WA_ACCESS_TOKEN      = <your NEW token from Meta Business Manager>
//     WA_PHONE_NUMBER_ID   = 1414203591773756        (WhatsApp Manager > Phone numbers; optional, has default)
//     WA_ADMIN_NUMBER      = 919355483400            (optional, has default)
//   NEW "Task Assigned" template (5 variables: task name, assigned to, description, deadline, status):
//     WA_ASSIGN_TEMPLATE   = <exact template name from WhatsApp Manager, e.g. new_task_assigned>
//     WA_ASSIGN_LANG       = en_US                   (optional, language code of that template)
//     WA_ASSIGN_PARAM_NAMES= task_name,assigned_to,description,deadline,status
//                            (ONLY if your template uses named variables like {{task_name}};
//                             leave empty if it uses {{1}} {{2}} {{3}} {{4}} {{5}})
// While WA_ASSIGN_TEMPLATE is empty, the old 3-variable template (task_reminder1) is used.
// Until WA_ACCESS_TOKEN is set, WhatsApp sends are skipped (nothing else breaks).
function getScriptProp_(key, fallback) {
  try {
    const v = PropertiesService.getScriptProperties().getProperty(key);
    return (v !== null && v !== undefined && v !== '') ? v : fallback;
  } catch (e) {
    return fallback;
  }
}
// Meta wants a LANGUAGE CODE ("en", "en_US", "hi"), not the word "English". This converts common names automatically.
function waLangCode_(lang) {
  const s = String(lang || '').trim();
  const k = s.toLowerCase().replace(/[\s()_-]+/g, '');
  const map = { english: 'en', englishus: 'en_US', enus: 'en_US', englishuk: 'en_GB', englishgb: 'en_GB', engb: 'en_GB', hindi: 'hi' };
  return map[k] || s || 'en';
}
function waAltLang_(code) { return code === 'en' ? 'en_US' : (code === 'en_US' ? 'en' : ''); }

// Short, plain-English reasons for the Meta error codes you are most likely to meet.
function waErrorHint_(code, msg) {
  const hints = {
    131030: 'Recipient number is not in the allowed list. A TEST number can only message numbers added in Meta > WhatsApp > API Setup > Manage phone number list. Use a real WhatsApp number for everyone else.',
    131042: 'Payment problem on the WhatsApp account. Add a valid payment method in Meta Business Settings > Billing.',
    132000: 'Number of variables sent does not match the template. Open the template in WhatsApp Manager: variables must be real numbered ones ({{1}} to {{5}}, added with the Add variable button), not typed {{}}.',
    132001: 'Template name or language does not match an APPROVED template. Run listWhatsAppTemplates and copy the exact name and language.',
    190:    'Access token is expired or wrong. Generate a new one and update WA_ACCESS_TOKEN.',
    100:    'Often a wrong WA_PHONE_NUMBER_ID, or the token has no access to that number. Run checkWhatsAppToken.'
  };
  return hints[code] ? hints[code] + ' [Meta: ' + msg + ']' : msg;
}

const WA_PHONE_NUMBER_ID = getScriptProp_('WA_PHONE_NUMBER_ID', '1414203591773756');   // phone number ID (NOT the template ID)
const WA_ACCESS_TOKEN    = getScriptProp_('WA_ACCESS_TOKEN', '');
const WA_API_VERSION     = 'v21.0';
const WA_TEMPLATE_NAME   = 'automatesystem';
const WA_TEMPLATE_LANG   = waLangCode_(getScriptProp_('WA_TEMPLATE_LANG', 'en'));
const WA_ADMIN_NUMBER    = getScriptProp_('WA_ADMIN_NUMBER', 'Number');
// New 5-variable "Task Assigned" template (see notes above)
const WA_ASSIGN_TEMPLATE    = getScriptProp_('WA_ASSIGN_TEMPLATE', 'automatesystem');   // your new 5-variable template
const WA_ASSIGN_LANG        = waLangCode_(getScriptProp_('WA_ASSIGN_LANG', 'en'));
const WA_ASSIGN_PARAM_NAMES = getScriptProp_('WA_ASSIGN_PARAM_NAMES', '');

function getEmployeeWhatsApp_(personKey) {
  const usersSheet = getCentralSheet_(USERS_SHEET);
  if (!usersSheet) return null;
  const data = usersSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][4] || '').toString().trim() === personKey) {
      const wa = (data[i][6] || '').toString().trim();
      return wa || null;
    }
  }
  return null;
}

function sendWhatsAppCloud_(toNumber, templateParams) {
  if (!toNumber) return false;
  if (!WA_ACCESS_TOKEN) {
    Logger.log('sendWhatsAppCloud_: WA_ACCESS_TOKEN is not set in Script Properties — skipping.');
    return false;
  }
  toNumber = toNumber.toString().replace(/[\s+]/g, '');

  const url = 'https://graph.facebook.com/' + WA_API_VERSION + '/' + WA_PHONE_NUMBER_ID + '/messages';

  const payload = {
    messaging_product: 'whatsapp',
    to: toNumber,
    type: 'template',
    template: {
      name: WA_TEMPLATE_NAME,
      language: { code: WA_TEMPLATE_LANG },
      components: [
        {
          type: 'body',
          parameters: templateParams.map(function (p) { return { type: 'text', text: String(p) }; })
        }
      ]
    }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + WA_ACCESS_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code === 200) {
      Logger.log('sendWhatsAppCloud_: sent to ' + toNumber);
      return true;
    }
    Logger.log('sendWhatsAppCloud_: FAILED for ' + toNumber + ' (' + code + '): ' + response.getContentText());
    return false;
  } catch (e) {
    Logger.log('sendWhatsAppCloud_: error for ' + toNumber + ': ' + e);
    return false;
  }
}

// Generic Cloud API sender for ANY template. params = array of strings in template order.
// paramNames (optional) = array of variable names, only for templates with NAMED variables.
// Returns { ok: true } or { ok: false, error: '...' }. Never throws.
function sendWhatsAppTemplateCloud_(toNumber, templateName, lang, params, paramNames) {
  const to = waNumber_(toNumber);                 // 10 digits -> adds 91 (helper from WhatsAppReminders file)
  if (!to) return { ok: false, error: 'No number' };
  if (!WA_ACCESS_TOKEN) return { ok: false, error: 'WA_ACCESS_TOKEN is not set in Script Properties' };
  const url = 'https://graph.facebook.com/' + WA_API_VERSION + '/' + WA_PHONE_NUMBER_ID + '/messages';
  const parameters = params.map(function (p, i) {
    const o = { type: 'text', text: waClean_(p) };          // no new lines / tabs, never empty, max 300 chars
    if (paramNames && paramNames[i]) o.parameter_name = paramNames[i];
    return o;
  });
  function attempt(code) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + WA_ACCESS_TOKEN },
      payload: JSON.stringify({
        messaging_product: 'whatsapp', to: to, type: 'template',
        template: { name: templateName, language: { code: code }, components: [{ type: 'body', parameters: parameters }] }
      }),
      muteHttpExceptions: true
    });
    const status = res.getResponseCode();
    let errCode = 0, errMsg = res.getContentText();
    if (status !== 200) {
      try {
        const j = JSON.parse(errMsg);
        errCode = (j.error && j.error.code) || 0;
        errMsg = (j.error && j.error.message) || errMsg;
        const det = j.error && j.error.error_data && j.error.error_data.details;
        if (det) errMsg += ' | DETAILS: ' + det;              // e.g. "number of localizable_params (5) does not match the expected number of params (0)"
      } catch (e) {}
    }
    return { status: status, errCode: errCode, errMsg: String(errMsg).slice(0, 400) };
  }
  try {
    const code = waLangCode_(lang);
    let r = attempt(code);
    if (r.status !== 200 && r.errCode === 132001) {        // template not found in that language: try the other English code
      const alt = waAltLang_(code);
      if (alt) {
        const r2 = attempt(alt);
        if (r2.status === 200) {
          Logger.log('NOTE: worked with language "' + alt + '". Set Script Property WA_ASSIGN_LANG = ' + alt + ' to skip the retry.');
          return { ok: true };
        }
      }
    }
    if (r.status === 200) return { ok: true };
    const hint = waErrorHint_(r.errCode, r.errMsg);
    Logger.log('sendWhatsAppTemplateCloud_ FAILED ' + to + ' (' + r.status + '): ' + hint);
    return { ok: false, error: hint };
  } catch (e) {
    Logger.log('sendWhatsAppTemplateCloud_ error ' + to + ': ' + e);
    return { ok: false, error: String(e) };
  }
}

// Sent to the employee (+ admin copy) when a task is added.
// New template (WA_ASSIGN_TEMPLATE set): Task name, Assigned To, Description, Deadline, Status.
// Otherwise: old 3-variable template, exactly as before.
function sendTaskAssignedWhatsAppCloud_(personKey, taskData) {
  try {
    const config = SHEETS_CONFIG[personKey];
    const personName = config ? config.name : (taskData.assignedTo || '');

    const targets = [];
    const employeeNumber = getEmployeeWhatsApp_(personKey);
    if (employeeNumber) targets.push(employeeNumber);
    if (WA_ADMIN_NUMBER) targets.push(WA_ADMIN_NUMBER);
    const seen = {};                                  // do not send twice to the same phone
    const unique = targets.filter(function (t) {
      const k = waNumber_(t);
      if (!k || seen[k]) return false;
      seen[k] = true;
      return true;
    });

    if (WA_ASSIGN_TEMPLATE) {
      const params = [
        taskData.taskName || '(untitled)',                       // {{1}} Task name
        personName || '-',                                       // {{2}} Assigned To
        taskData.description || '-',                             // {{3}} Description
        waDmy_(waDateKey_(taskData.deadline)),                   // {{4}} Deadline  dd/MM/yyyy (or "Not set")
        'Pending'                                                // {{5}} Status (every new task starts as Pending)
      ];
      const names = WA_ASSIGN_PARAM_NAMES
        ? WA_ASSIGN_PARAM_NAMES.split(',').map(function (x) { return x.trim(); })
        : null;
      unique.forEach(function (t) {
        const r = sendWhatsAppTemplateCloud_(t, WA_ASSIGN_TEMPLATE, WA_ASSIGN_LANG, params, names);
        if (!r.ok) Logger.log('Task-assigned WhatsApp not sent to ' + t + ': ' + r.error);
      });
      return;
    }

    // fallback: previous behaviour (3 variables, template task_reminder1)
    const oldParams = [taskData.taskName || '(untitled)', personName, taskData.deadline || 'Not set'];
    unique.forEach(function (t) { sendWhatsAppCloud_(t, oldParams); });
  } catch (e) {
    Logger.log('sendTaskAssignedWhatsAppCloud_ error: ' + e);
  }
}

// RUN FROM THE EDITOR: sends a sample "Task Assigned" message to WA_ADMIN_NUMBER and logs the result.
function testAssignTemplateSend() {
  if (!WA_ASSIGN_TEMPLATE) { Logger.log('Set WA_ASSIGN_TEMPLATE in Script Properties first.'); return; }
  const names = WA_ASSIGN_PARAM_NAMES ? WA_ASSIGN_PARAM_NAMES.split(',').map(function (x) { return x.trim(); }) : null;
  const r = sendWhatsAppTemplateCloud_(WA_ADMIN_NUMBER, WA_ASSIGN_TEMPLATE, WA_ASSIGN_LANG,
    ['Sample task', 'Test User', 'This is a test description', '25/11/2026', 'Pending'], names);
  Logger.log('testAssignTemplateSend: ' + (r.ok ? 'SUCCESS' : 'FAILED - ' + r.error));
}

// RUN FROM THE EDITOR (then View > Logs): is the token + phone number ID correct?
function checkWhatsAppToken() {
  if (!WA_ACCESS_TOKEN) { Logger.log('NO TOKEN: add Script Property WA_ACCESS_TOKEN'); return; }
  Logger.log('Token length: ' + WA_ACCESS_TOKEN.length + ' | starts with: ' + WA_ACCESS_TOKEN.slice(0, 3) + ' | has spaces: ' + /\s/.test(WA_ACCESS_TOKEN));
  Logger.log('Phone number ID used by the code: ' + WA_PHONE_NUMBER_ID + '  (must match WhatsApp Manager > Phone numbers)');
  const url = 'https://graph.facebook.com/' + WA_API_VERSION + '/' + WA_PHONE_NUMBER_ID + '?fields=display_phone_number,verified_name,quality_rating';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + WA_ACCESS_TOKEN }, muteHttpExceptions: true });
  Logger.log(res.getResponseCode() + ' ' + res.getContentText());
  Logger.log('Assign template: ' + WA_ASSIGN_TEMPLATE + ' | language: ' + WA_ASSIGN_LANG + ' | old/reminder template: ' + WA_TEMPLATE_NAME + ' | language: ' + WA_TEMPLATE_LANG);
}

// RUN FROM THE EDITOR: lists your templates with exact name, language, status and number of variables.
// Needs Script Property WA_WABA_ID = your WhatsApp Business Account ID (shown in WhatsApp Manager / Business Settings > Accounts > WhatsApp accounts).
function listWhatsAppTemplates() {
  const waba = getScriptProp_('WA_WABA_ID', '');
  if (!waba) { Logger.log('Add Script Property WA_WABA_ID (your WhatsApp Business Account ID) first.'); return; }
  const url = 'https://graph.facebook.com/' + WA_API_VERSION + '/' + waba + '/message_templates?limit=100&fields=name,language,status,components';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + WA_ACCESS_TOKEN }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) { Logger.log('FAILED (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 400)); return; }
  (JSON.parse(res.getContentText()).data || []).forEach(function (t) {
    const body = (t.components || []).filter(function (c) { return c.type === 'BODY'; })[0];
    const vars = body && body.text ? (body.text.match(/\{\{[^}]+\}\}/g) || []) : [];
    Logger.log(t.name + '  |  language: ' + t.language + '  |  status: ' + t.status + '  |  variables: ' + vars.length + ' ' + vars.join(' '));
  });
}

function testWhatsAppCloudSend() {
  const success = sendWhatsAppCloud_(WA_ADMIN_NUMBER, ['Sample Task', 'Test User', '20/09/2026']);
  Logger.log('testWhatsAppCloudSend result: ' + (success ? 'SUCCESS' : 'FAILED — check logs above'));
}

// ============================================================
// ATTENDANCE (Google Form responses sheet)
// ============================================================
// ============================================================
// ATTENDANCE (Google Form responses sheet)
// ============================================================
const ATTENDANCE_SHEET_ID   = '';  //pest the sheet id here to link the attandance sheet
const ATTENDANCE_SHEET_NAME = 'Form Responses 1';
const ATTENDANCE_TZ         = 'Asia/Kolkata';   // fixed so "today" is always India date

function getAttendanceTab_() {
  const ss = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
  const norm = function (s) { return String(s || '').toLowerCase().replace(/[\s_]+/g, ''); };
  const target = norm(ATTENDANCE_SHEET_NAME);
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) {
    if (norm(all[i].getName()) === target) return all[i];
  }
  throw new Error('Attendance tab "' + ATTENDANCE_SHEET_NAME + '" not found. Available tabs: ' +
    all.map(function (s) { return s.getName(); }).join(', '));
}

// Reads a cell value as a Date (handles real dates AND text like "9/30/2026 9:15:00")
function attToDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (v === '' || v === null || v === undefined) return null;
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    // Form sheets in India are usually M/d/yyyy; if first number > 12 it must be d/M/yyyy
    let a = Number(m[1]), b = Number(m[2]);
    const month = a > 12 ? b : a;
    const day   = a > 12 ? a : b;
    return new Date(Number(m[3]), month - 1, day, Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// Finds Timestamp / Name / Status columns by header (0-based). Falls back to A, B, E.
function getAttendanceCols_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 5);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h || '').trim().toLowerCase(); });

  let ts = headers.findIndex(function (h) { return h.indexOf('timestamp') !== -1; });
  let nm = headers.findIndex(function (h) { return h.indexOf('name') !== -1; });
  let st = headers.findIndex(function (h) {
    return h.indexOf('status') !== -1 || h.indexOf('present') !== -1 || h.indexOf('attendance') !== -1;
  });
  if (ts === -1) ts = 0;
  if (nm === -1) nm = 1;
  if (st === -1) st = 4;
  // "Logout Time" column (created automatically on the first counted logout). -1 = not created yet.
  const lo = headers.findIndex(function (h) { return h.indexOf('logout') !== -1; });
  return { ts: ts, nm: nm, st: st, lo: lo, width: Math.max(lastCol, st + 1, nm + 1, ts + 1) };
}

// === AUTO PUNCH ===
// Writes ONE row the first time a person logs in on a given day.
function punchAttendance_(personKey, fullName) {
  const cfg  = personKey ? SHEETS_CONFIG[personKey] : null;
  const name = String(cfg ? cfg.name : fullName || '').trim();
  if (!name) return { success: false, message: 'No name to punch' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh   = getAttendanceTab_();
    const cols = getAttendanceCols_(sh);
    const now  = new Date();
    const todayKey = Utilities.formatDate(now, ATTENDANCE_TZ, 'yyyy-MM-dd');

    // last row that has a name
    const maxRows = sh.getMaxRows();
    const nameVals = sh.getRange(1, cols.nm + 1, maxRows, 1).getValues();
    let last = 1;
    for (let i = nameVals.length - 1; i >= 1; i--) {
      if (String(nameVals[i][0]).trim() !== '') { last = i + 1; break; }
    }

    // already punched today? (check last 300 rows)
    if (last >= 2) {
      const from = Math.max(2, last - 300);
      const rows = sh.getRange(from, 1, last - from + 1, cols.width).getValues();
      for (let r = 0; r < rows.length; r++) {
        const ts = attToDate_(rows[r][cols.ts]);
        if (ts &&
            Utilities.formatDate(ts, ATTENDANCE_TZ, 'yyyy-MM-dd') === todayKey &&
            normalizeName_(rows[r][cols.nm]) === normalizeName_(name)) {
          return { success: true, already: true, message: 'Already marked today' };
        }
      }
    }

    // make sure the new row exists (this was the crash when the sheet was full)
    const row = last + 1;
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());

    sh.getRange(row, cols.ts + 1).setValue(now).setNumberFormat('M/d/yyyy H:mm:ss');
    sh.getRange(row, cols.nm + 1).setValue(name);
    sh.getRange(row, cols.st + 1).setValue('PRESENT');
    SpreadsheetApp.flush();

    return { success: true, already: false, message: 'Attendance punched in', row: row };
  } finally {
    lock.releaseLock();
  }
}

// Dashboard can call this after login; safe to call repeatedly.
function markAttendance(token) {
  const session = verifySession_(token);
  try {
    return punchAttendance_(session.personKey, session.fullName);
  } catch (e) {
    Logger.log('markAttendance error: ' + e);
    return { success: false, message: 'Attendance error: ' + e.message };
  }
}

// === AUTO LOGOUT TIME ===
// Rule: every time the person presses Logout, the time is saved on today's punch-in row.
// If they log out more than once in a day, the LAST logout is kept.

function punchLogout_(personKey, fullName) {
  const cfg  = personKey ? SHEETS_CONFIG[personKey] : null;
  const name = String(cfg ? cfg.name : fullName || '').trim();
  if (!name) return { success: false, message: 'No name to log out' };

  const now = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh   = getAttendanceTab_();
    const cols = getAttendanceCols_(sh);
    const todayKey = Utilities.formatDate(now, ATTENDANCE_TZ, 'yyyy-MM-dd');

    const last = sh.getLastRow();
    if (last < 2) return { success: true, counted: false, message: 'No punch-in found for today' };

    // find today's punch-in row of this person (look at the newest 300 rows)
    const from = Math.max(2, last - 300);
    const width = Math.max(cols.width, cols.lo + 1);
    const rows = sh.getRange(from, 1, last - from + 1, width).getValues();
    let target = -1;
    for (let r = rows.length - 1; r >= 0; r--) {
      const ts = attToDate_(rows[r][cols.ts]);
      if (!ts) continue;
      if (Utilities.formatDate(ts, ATTENDANCE_TZ, 'yyyy-MM-dd') !== todayKey) continue;
      if (normalizeName_(rows[r][cols.nm]) !== normalizeName_(name)) continue;
      if (attIsLeaveStatus_(rows[r][cols.st])) continue;
      target = from + r;
      break;
    }
    if (target === -1) return { success: true, counted: false, message: 'No punch-in found for today' };

    // create the "Logout Time" header the first time
    let lo = cols.lo;
    if (lo === -1) {
      lo = sh.getLastColumn();                       // 0-based index of the new last+1 column
      sh.getRange(1, lo + 1).setValue('Logout Time').setFontWeight('bold');
    }
    sh.getRange(target, lo + 1).setValue(now).setNumberFormat('M/d/yyyy H:mm:ss');   // the last logout wins
    SpreadsheetApp.flush();

    return { success: true, counted: true, message: 'Logout time saved', row: target };
  } finally {
    lock.releaseLock();
  }
}

// The app calls this when the user presses Logout. Safe to call at any time.
function markLogout(token) {
  const session = verifySession_(token);
  try {
    return punchLogout_(session.personKey, session.fullName);
  } catch (e) {
    Logger.log('markLogout error: ' + e);
    return { success: false, message: 'Logout error: ' + e.message };
  } finally {
    sessionEnd_(token);          // pressing Logout also ends the login on the server
  }
}

function testAttendanceLogout() {
  // change 'person1' to any key from SHEETS_CONFIG (saves the logout time on today's punch-in row)
  Logger.log(JSON.stringify(punchLogout_('person1', 'Test')));
}

function readAttendanceRecords_(sheet, dateStr) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const cols = getAttendanceCols_(sheet);
  const records = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ts = attToDate_(row[cols.ts]);
    if (!ts) continue;
    if (Utilities.formatDate(ts, ATTENDANCE_TZ, 'yyyy-MM-dd') !== dateStr) continue;
    records.push({
      tsRaw: ts.getTime(),
      timestamp: Utilities.formatDate(ts, ATTENDANCE_TZ, 'dd/MM/yyyy HH:mm'),
      name: String(row[cols.nm] || '').trim(),
      status: String(row[cols.st] || 'Present').trim() || 'Present'
    });
  }
  records.sort(function (a, b) { return b.tsRaw - a.tsRaw; });
  return records;
}

function getTodayAttendanceSummary(token) {
  const session = verifySession_(token);
  assertIsAdmin_(session);

  const todayStr = Utilities.formatDate(new Date(), ATTENDANCE_TZ, 'yyyy-MM-dd');
  const records  = readAttendanceRecords_(getAttendanceTab_(), todayStr);

  const anyRowNames = new Set(records.map(function (r) { return normalizeName_(r.name); }));
  const presentNames = new Set(records
    .filter(function (r) { return String(r.status || '').toLowerCase().indexOf('leave') === -1; })
    .map(function (r) { return normalizeName_(r.name); }));
  const allPeople = Object.values(SHEETS_CONFIG).map(function (c) { return c.name; });

  // people with no row today whose day is a holiday / weekly off (shown as "Day off", not Absent)
  const hol = attReadHolidays_(todayStr, todayStr)[todayStr] || '';
  const off = allPeople
    .filter(function (n) {
      return !anyRowNames.has(normalizeName_(n)) && (hol || attIsWeeklyOff_(n, todayStr));
    })
    .map(function (n) { return { name: n, reason: hol ? 'Holiday: ' + hol : 'Weekly off' }; });
  const offNames = new Set(off.map(function (o) { return normalizeName_(o.name); }));

  const present = allPeople.filter(function (n) { return presentNames.has(normalizeName_(n)); });
  const missing = allPeople.filter(function (n) {
    return !anyRowNames.has(normalizeName_(n)) && !offNames.has(normalizeName_(n));
  });

  return sanitizeForClient_({ date: todayStr, present: present, missing: missing, off: off, records: records });
}

function getAttendanceForDate(token, dateStr) {
  const session = verifySession_(token);
  assertIsAdmin_(session);
  if (!dateStr) throw new Error('Provide a date in yyyy-MM-dd format.');
  return sanitizeForClient_(readAttendanceRecords_(getAttendanceTab_(), dateStr));
}

// ------------------------------------------------------------
// RUN THESE FROM THE APPS SCRIPT EDITOR TO FIND THE PROBLEM
// ------------------------------------------------------------
function testAttendancePunch() {
  // change 'person1' to any key from SHEETS_CONFIG
  const res = punchAttendance_('person1', 'Test');
  Logger.log(JSON.stringify(res));
}

function debugAttendance() {
  try {
    const sh = getAttendanceTab_();
    const cols = getAttendanceCols_(sh);
    Logger.log('Tab: ' + sh.getName() + ' | maxRows=' + sh.getMaxRows() + ' | lastRow=' + sh.getLastRow());
    Logger.log('Header row: ' + JSON.stringify(sh.getRange(1, 1, 1, cols.width).getValues()[0]));
    Logger.log('Columns used (1-based) -> Timestamp:' + (cols.ts + 1) + ' Name:' + (cols.nm + 1) + ' Status:' + (cols.st + 1));
    Logger.log('Script timezone: ' + Session.getScriptTimeZone() + ' | using: ' + ATTENDANCE_TZ);
  } catch (e) {
    Logger.log('ATTENDANCE ERROR: ' + e);
  }
}
// ============================================================
// ATTENDANCE HISTORY + MONTHLY VIEW  (+ weekly offs, holidays, leave)
// Lives in the same project as the attendance code above.
// It reuses: getAttendanceTab_, getAttendanceCols_, attToDate_,
// normalizeName_, verifySession_, assertIsAdmin_, sanitizeForClient_,
// SHEETS_CONFIG, ATTENDANCE_SHEET_ID, ATTENDANCE_TZ
// ============================================================

// Weekly off for EVERYONE (0 = Sunday ... 6 = Saturday)
const ATT_WEEKLY_OFF = [0];

// Extra off ONLY for these people: the Saturdays listed below (2nd and 4th of every month)
const ATT_SAT_OFF_NAMES = ['sonu'];
const ATT_SAT_OFF_WEEKS = [2, 4];

// ONLY these admins can mark holidays and leave.
// Matched against the login username, the full name and the person name in SHEETS_CONFIG.
const ATT_MANAGER_NAMES = ['sonu'];

const HOLIDAYS_SHEET = 'Holidays';   // created automatically inside the attendance spreadsheet

// ---------- helpers ----------

function attIsAdmin_(session) {
  try { assertIsAdmin_(session); return true; } catch (e) { return false; }
}

function attSessionName_(session) {
  const cfg = session.personKey ? SHEETS_CONFIG[session.personKey] : null;
  return String(cfg ? cfg.name : session.fullName || '').trim();
}

// true only for an Admin/Manager whose username / name is in ATT_MANAGER_NAMES
function attCanManage_(session) {
  if (!attIsAdmin_(session)) return false;
  const ids = [session.username, session.fullName, attSessionName_(session)].map(normalizeName_);
  return ATT_MANAGER_NAMES.some(function (n) {
    const want = normalizeName_(n);
    return ids.some(function (id) { return id === want || id.split(' ').indexOf(want) !== -1; });
  });
}

function assertCanManageAttendance_(session) {
  if (!attCanManage_(session)) throw new Error('Only Sonu (admin) can mark holidays and leave.');
}

function attIsLeaveStatus_(s) {
  return String(s || '').toLowerCase().indexOf('leave') !== -1;
}

// 'yyyy-MM-dd' for a Date or a text date
function attKey_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, ATTENDANCE_TZ, 'yyyy-MM-dd');
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[1] + '-' + m[2] + '-' + m[3] : '';
}

function attCheckDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) throw new Error('Date must be in yyyy-MM-dd format.');
  return s;
}

// Sunday for everyone; 2nd and 4th Saturday for the people in ATT_SAT_OFF_NAMES
function attIsWeeklyOff_(name, dateKey) {
  const p = dateKey.split('-').map(Number);
  const dow = new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
  if (ATT_WEEKLY_OFF.indexOf(dow) !== -1) return true;
  if (dow === 6 && ATT_SAT_OFF_WEEKS.indexOf(Math.ceil(p[2] / 7)) !== -1) {
    const n = normalizeName_(name);
    return ATT_SAT_OFF_NAMES.some(function (x) { return normalizeName_(x) === n; });
  }
  return false;
}

// the real name as written in SHEETS_CONFIG (so a typo can't create a ghost person)
function attResolvePerson_(name) {
  const want = normalizeName_(name);
  const all = Object.values(SHEETS_CONFIG).map(function (c) { return c.name; });
  for (let i = 0; i < all.length; i++) if (normalizeName_(all[i]) === want) return all[i];
  throw new Error('Unknown employee: ' + name);
}

// ---------- holidays tab ----------

function attHolidaysSheet_(create) {
  const ss = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
  let sh = ss.getSheetByName(HOLIDAYS_SHEET);
  if (sh || !create) return sh;
  sh = ss.insertSheet(HOLIDAYS_SHEET);
  sh.appendRow(['Date', 'Title', 'AddedBy', 'AddedOn']);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  sh.getRange('A:A').setNumberFormat('@');
  return sh;
}

// { 'yyyy-MM-dd': 'Diwali', ... } for the date range
function attReadHolidays_(fromStr, toStr) {
  const out = {};
  const sh = attHolidaysSheet_(false);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    const k = attKey_(r[0]);
    if (k && k >= fromStr && k <= toStr) out[k] = String(r[1] || 'Holiday').trim() || 'Holiday';
  });
  return out;
}

// ---------- reading punches ----------

// Reads the sheet ONCE and returns every row between fromStr..toStr (yyyy-MM-dd, inclusive)
function readAttendanceRange_(sheet, fromStr, toStr) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const cols = getAttendanceCols_(sheet);
  const out = [];

  for (let i = 1; i < data.length; i++) {
    const ts = attToDate_(data[i][cols.ts]);
    if (!ts) continue;
    const dateKey = Utilities.formatDate(ts, ATTENDANCE_TZ, 'yyyy-MM-dd');
    if (dateKey < fromStr || dateKey > toStr) continue;   // yyyy-MM-dd compares correctly as text
    const loDate = cols.lo !== -1 ? attToDate_(data[i][cols.lo]) : null;
    out.push({
      tsRaw: ts.getTime(),
      date: dateKey,
      day: Utilities.formatDate(ts, ATTENDANCE_TZ, 'EEE'),
      time: Utilities.formatDate(ts, ATTENDANCE_TZ, 'hh:mm a'),
      logout: loDate ? Utilities.formatDate(loDate, ATTENDANCE_TZ, 'hh:mm a') : '',
      timestamp: Utilities.formatDate(ts, ATTENDANCE_TZ, 'dd/MM/yyyy HH:mm'),
      name: String(data[i][cols.nm] || '').trim(),
      status: String(data[i][cols.st] || 'Present').trim() || 'Present',
      rowNumber: i + 1
    });
  }
  out.sort(function (a, b) { return b.tsRaw - a.tsRaw; });   // newest first
  return out;
}

// ---------- 1) HISTORY (date range, punch times) ----------
// Admin: any person (personName optional = everyone). Employee: only their own.
function getAttendanceHistory(token, fromStr, toStr, personName) {
  const session = verifySession_(token);
  const isAdmin = attIsAdmin_(session);

  const today = Utilities.formatDate(new Date(), ATTENDANCE_TZ, 'yyyy-MM-dd');
  if (!toStr) toStr = today;
  if (!fromStr) {
    const d = new Date(); d.setDate(d.getDate() - 30);
    fromStr = Utilities.formatDate(d, ATTENDANCE_TZ, 'yyyy-MM-dd');
  }

  let records = readAttendanceRange_(getAttendanceTab_(), fromStr, toStr);

  // Non-admins can only ever see their own rows
  const filterName = isAdmin ? String(personName || '').trim() : attSessionName_(session);
  if (filterName) {
    const key = normalizeName_(filterName);
    records = records.filter(function (r) { return normalizeName_(r.name) === key; });
  }

  return sanitizeForClient_({ from: fromStr, to: toStr, person: filterName || 'All', count: records.length, records: records });
}

// ---------- 2) MONTH VIEW (one month, per person, day by day) ----------
// Day status, in this order:
//   LEAVE (marked by Sonu)  >  PRESENT (punched)  >  HOLIDAY  >  WEEKLY OFF  >  UPCOMING / ABSENT
function getMonthlyAttendance(token, monthStr, personName) {
  const session = verifySession_(token);
  const isAdmin = attIsAdmin_(session);

  const nowKey = Utilities.formatDate(new Date(), ATTENDANCE_TZ, 'yyyy-MM-dd');
  if (!monthStr) monthStr = nowKey.substring(0, 7);
  const parts = monthStr.split('-');
  const year = Number(parts[0]), month = Number(parts[1]);
  if (!year || !month || month < 1 || month > 12) throw new Error('Month must be in yyyy-MM format.');

  const daysInMonth = new Date(year, month, 0).getDate();
  const pad = function (n) { return n < 10 ? '0' + n : '' + n; };
  const firstKey = monthStr + '-01';
  const lastKey  = monthStr + '-' + pad(daysInMonth);

  const records  = readAttendanceRange_(getAttendanceTab_(), firstKey, lastKey);
  const holidays = attReadHolidays_(firstKey, lastKey);

  // who to show
  let people;
  if (!isAdmin) {
    people = [attSessionName_(session)];
  } else if (personName) {
    people = [String(personName).trim()];
  } else {
    people = Object.values(SHEETS_CONFIG).map(function (c) { return c.name; });
  }

  // index: person -> date -> { leave, time, tsRaw }   (earliest punch wins; a LEAVE row wins over everything)
  const byPerson = {};
  records.forEach(function (r) {
    const k = normalizeName_(r.name);
    if (!byPerson[k]) byPerson[k] = {};
    const prev = byPerson[k][r.date];
    const leave = attIsLeaveStatus_(r.status);
    if (!prev || (leave && !prev.leave) || (!leave && !prev.leave && r.tsRaw < prev.tsRaw)) {
      byPerson[k][r.date] = { leave: leave, time: r.time, logout: r.logout || '', tsRaw: r.tsRaw };
    }
  });

  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const result = people.map(function (name) {
    const punches = byPerson[normalizeName_(name)] || {};
    const days = [];
    let present = 0, absent = 0, off = 0, hol = 0, leaves = 0;

    for (let d = 1; d <= daysInMonth; d++) {
      const dateKey = monthStr + '-' + pad(d);
      const dow = new Date(year, month - 1, d).getDay();
      const rec = punches[dateKey];
      const holTitle = holidays[dateKey] || '';
      let status, time = '', logout = '', note = holTitle;

      if (rec && rec.leave) { status = 'LEAVE'; leaves++; }
      else if (rec) { status = 'PRESENT'; time = rec.time; logout = rec.logout || ''; present++; }
      else if (holTitle) { status = 'HOLIDAY'; hol++; }
      else if (attIsWeeklyOff_(name, dateKey)) { status = 'WEEKLY OFF'; off++; }
      else if (dateKey > nowKey) { status = 'UPCOMING'; }
      else { status = 'ABSENT'; absent++; }

      days.push({ date: dateKey, day: dayNames[dow], status: status, time: time, logout: logout, note: note });
    }
    return {
      name: name, presentDays: present, absentDays: absent, weeklyOffDays: off,
      holidayDays: hol, leaveDays: leaves, days: days
    };
  });

  return sanitizeForClient_({
    month: monthStr,
    daysInMonth: daysInMonth,
    canManage: attCanManage_(session),
    holidays: Object.keys(holidays).sort().map(function (k) { return { date: k, title: holidays[k] }; }),
    people: result
  });
}

// ---------- 3) HOLIDAYS (company-wide) — ONLY Sonu ----------

function markHoliday(token, dateStr, title) {
  const session = verifySession_(token);
  assertCanManageAttendance_(session);
  attCheckDate_(dateStr);
  title = String(title || '').trim();
  if (!title) throw new Error('Enter the holiday name (for example: Diwali).');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = attHolidaysSheet_(true);
    const by = String(session.fullName || session.username || '');
    if (sh.getLastRow() >= 2) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (attKey_(rows[i][0]) === dateStr) {          // already a holiday: just rename it
          sh.getRange(i + 2, 2).setValue(title);
          sh.getRange(i + 2, 3).setValue(by);
          return { success: true, message: 'Holiday updated: ' + title + ' (' + dateStr + ')' };
        }
      }
    }
    sh.appendRow([dateStr, title, by, new Date()]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('@').setValue(dateStr);
    return { success: true, message: 'Holiday added: ' + title + ' (' + dateStr + ')' };
  } finally {
    lock.releaseLock();
  }
}

function removeHoliday(token, dateStr) {
  const session = verifySession_(token);
  assertCanManageAttendance_(session);
  attCheckDate_(dateStr);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = attHolidaysSheet_(false);
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No holiday on ' + dateStr };
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (attKey_(rows[i][0]) === dateStr) { sh.deleteRow(i + 2); return { success: true, message: 'Holiday removed (' + dateStr + ')' }; }
    }
    return { success: false, message: 'No holiday on ' + dateStr };
  } finally {
    lock.releaseLock();
  }
}

// ---------- 4) LEAVE (per employee, one day or a date range) — ONLY Sonu ----------
// A leave is stored as a normal row in the attendance sheet with Status = LEAVE.
// Because that row exists, the auto-punch at login will NOT turn the day into PRESENT.
// Days that are already weekly off / holiday are skipped. Days with a real punch are NOT overwritten.
function markLeave(token, personName, fromStr, toStr) {
  const session = verifySession_(token);
  assertCanManageAttendance_(session);
  const name = attResolvePerson_(personName);
  attCheckDate_(fromStr);
  if (!toStr) toStr = fromStr;
  attCheckDate_(toStr);
  if (toStr < fromStr) throw new Error('"To" date must be on or after the "From" date.');

  const f = fromStr.split('-').map(Number), t = toStr.split('-').map(Number);
  const span = Math.round((Date.UTC(t[0], t[1] - 1, t[2]) - Date.UTC(f[0], f[1] - 1, f[2])) / 86400000) + 1;
  if (span > 31) throw new Error('Mark at most 31 days at a time.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getAttendanceTab_();
    const cols = getAttendanceCols_(sh);
    const holidays = attReadHolidays_(fromStr, toStr);
    const mine = readAttendanceRange_(sh, fromStr, toStr).filter(function (r) {
      return normalizeName_(r.name) === normalizeName_(name);
    });
    const have = {};
    mine.forEach(function (r) { have[r.date] = attIsLeaveStatus_(r.status) ? 'leave' : 'present'; });

    const toAdd = [];
    let skippedOff = 0, skippedPresent = 0, already = 0;
    for (let i = 0; i < span; i++) {
      const key = Utilities.formatDate(new Date(Date.UTC(f[0], f[1] - 1, f[2] + i)), 'UTC', 'yyyy-MM-dd');
      if (holidays[key] || attIsWeeklyOff_(name, key)) { skippedOff++; continue; }
      if (have[key] === 'leave') { already++; continue; }
      if (have[key] === 'present') { skippedPresent++; continue; }
      toAdd.push(key);
    }

    if (toAdd.length) {
      // first empty row after the last row that has a name
      const nameVals = sh.getRange(1, cols.nm + 1, sh.getMaxRows(), 1).getValues();
      let last = 1;
      for (let i = nameVals.length - 1; i >= 1; i--) {
        if (String(nameVals[i][0]).trim() !== '') { last = i + 1; break; }
      }
      const needRows = last + toAdd.length;
      if (needRows > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());
      toAdd.forEach(function (key, idx) {
        const row = last + 1 + idx;
        const ts = Utilities.parseDate(key + ' 09:00', ATTENDANCE_TZ, 'yyyy-MM-dd HH:mm');
        sh.getRange(row, cols.ts + 1).setValue(ts).setNumberFormat('M/d/yyyy H:mm:ss');
        sh.getRange(row, cols.nm + 1).setValue(name);
        sh.getRange(row, cols.st + 1).setValue('LEAVE');
      });
      SpreadsheetApp.flush();
    }

    let msg = name + ': leave marked for ' + toAdd.length + ' day' + (toAdd.length === 1 ? '' : 's');
    const extra = [];
    if (already) extra.push(already + ' already on leave');
    if (skippedOff) extra.push(skippedOff + ' off/holiday day(s) skipped');
    if (skippedPresent) extra.push(skippedPresent + ' day(s) skipped because ' + name + ' already punched in');
    if (extra.length) msg += ' (' + extra.join(', ') + ')';
    return { success: toAdd.length > 0 || already > 0, added: toAdd.length, message: msg + '.' };
  } finally {
    lock.releaseLock();
  }
}

function removeLeave(token, personName, dateStr) {
  const session = verifySession_(token);
  assertCanManageAttendance_(session);
  const name = attResolvePerson_(personName);
  attCheckDate_(dateStr);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getAttendanceTab_();
    const rows = readAttendanceRange_(sh, dateStr, dateStr).filter(function (r) {
      return normalizeName_(r.name) === normalizeName_(name) && attIsLeaveStatus_(r.status);
    });
    if (!rows.length) return { success: false, message: 'No leave found for ' + name + ' on ' + dateStr + '.' };
    rows.map(function (r) { return r.rowNumber; })
      .sort(function (a, b) { return b - a; })               // delete from the bottom up
      .forEach(function (n) { sh.deleteRow(n); });
    return { success: true, message: 'Leave removed for ' + name + ' (' + dateStr + ').' };
  } finally {
    lock.releaseLock();
  }
}

// ---------- test from the Apps Script editor ----------
function testAttendanceHistory() {
  const res = readAttendanceRange_(getAttendanceTab_(), '2026-10-01', '2026-10-31');
  Logger.log('Rows in October: ' + res.length);
  Logger.log(JSON.stringify(res.slice(0, 5)));
}

// Sonu's Saturdays: should log 10 and 24 for October 2026 (2nd and 4th Saturday)
function testSonuSaturdays() {
  const offs = [];
  for (let d = 1; d <= 31; d++) {
    const key = '2026-10-' + (d < 10 ? '0' + d : d);
    if (attIsWeeklyOff_('Sonu', key) && new Date(2026, 9, d).getDay() === 6) offs.push(d);
  }
  Logger.log('Sonu Saturday offs in Oct 2026: ' + offs.join(', '));
  Logger.log('Mukesh Sat 10 Oct off? ' + attIsWeeklyOff_('Mukesh', '2026-10-10') + ' (should be false)');
}

// ============================================================
// MIS SCORE — Delegation / Checklist / FMS / Combined
// ------------------------------------------------------------
// Sources:
//   Delegation = the Task Ledger tasks (SHEETS_CONFIG sheets)
//   Checklist  = each person's personal sheets listed in the SheetLinks tab (any label WITHOUT "FMS")
//   FMS        = personal sheets whose label contains "FMS" (multi-stage sheets: every stage is scored)
// Only items whose planned date / deadline is inside the chosen period are counted, and only if they are
// done, or already overdue. Open items due today or later are not counted yet.
//   Planned          = done on time + done late + not done
//   % Work not done  = Not done / Planned x 100
//   % Not on time    = (Done late + Not done) / Planned x 100
//   MIS score        = -(% Work not done + % Not on time)      -> 0 is perfect, lower is worse
// "Done late" needs an actual / done date; when that is empty the item counts as on time.
// ============================================================
const MIS_FMS_LABEL = /fms/i;          // label OR tab name contains FMS  -> scored as FMS
const MIS_CHECKLIST_LABEL = /checklist/i;   // label contains Checklist (Ultimate Checklist, Bond Checklist...) -> scored as Checklist
// every other personal sheet (Help Ticket, Sales, IMS, Tank tracker...) is NOT part of MIS
const MIS_CACHE_SECONDS = 600;         // Refresh in the dashboard bypasses this

function misKey_(v) {
  if (!v) return '';
  if (v instanceof Date) return attKey_(v);
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return attKey_(s);
}

// A sheet cell -> 'yyyy-MM-dd' ('' when it is not a real date)
function misCellKey_(v) {
  if (v === '' || v === null || v === undefined) return '';
  let k = '';
  if (v instanceof Date) k = attKey_(v);
  else {
    const s = String(v).trim();
    const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);     // Indian style d/m/yyyy
    if (m) k = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    else { const d = new Date(s); k = isNaN(d.getTime()) ? '' : attKey_(d); }
  }
  return k && k >= '2000-01-01' ? k : '';
}

function misRound_(n) { return Math.round(n * 10) / 10; }

function misIsDone_(raw) {
  const s = String(raw == null ? '' : raw).trim().toLowerCase();
  const notDone = /\bnot\b/.test(s) || s === 'pending' || s === 'no' || s.indexOf('incomplete') !== -1;
  return !notDone && DONE_KEYWORDS.some(function (k) { return s.indexOf(k) !== -1; });
}

function misNewCounts_() { return { onTime: 0, late: 0, notDone: 0, items: [] }; }

// one item: planned date + done? + actual date -> counted into c (does nothing for items not yet due)
function misCellMs_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  let s;
  if (v instanceof Date) s = Utilities.formatDate(v, ATTENDANCE_TZ, 'dd/MM/yyyy HH:mm:ss');
  else s = String(v).trim();
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return 0;
  return Date.UTC(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
}

// plannedMs/actualMs are optional (FMS uses them: its TAT is in hours, so same-day delay counts as late)
function misCountItem_(c, plannedKey, done, actualKey, todayKey, fromStr, toStr, title, source, plannedMs, actualMs) {
  if (!plannedKey || plannedKey < fromStr || plannedKey > toStr) return;
  if (done) {
    const isLate = (plannedMs && actualMs) ? actualMs > plannedMs : (actualKey && actualKey > plannedKey);
    if (isLate) {
      c.late++; c.items.push({ task: title, deadline: plannedKey, doneDate: actualKey, kind: 'Done late', source: source });
    } else c.onTime++;
  } else if (plannedKey < todayKey) {
    c.notDone++; c.items.push({ task: title, deadline: plannedKey, doneDate: '', kind: 'Not done', source: source });
  }
}

// ---------- Delegation (Task Ledger) ----------
function misDelegationCounts_(personKey, fromStr, toStr, todayKey) {
  const c = misNewCounts_();
  let tasks = [];
  try { tasks = getTasksByPersonRaw_(personKey); }
  catch (e) { Logger.log('MIS delegation read error ' + personKey + ': ' + e); }
  tasks.forEach(function (t) {
    const done = DONE_KEYWORDS.indexOf(String(t.Status || '').trim().toLowerCase()) !== -1;
    misCountItem_(c, misKey_(t.Deadline), done, misKey_(t.DoneDate), todayKey, fromStr, toStr, t.TaskName, 'Delegation');
  });
  return c;
}

// ---------- Checklist / FMS (personal sheets from SheetLinks) ----------
function misUsers_() {
  const out = {};
  const sh = getCentralSheet_(USERS_SHEET);
  if (!sh) return out;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const pk = (data[i][4] || '').toString().trim();
    if (pk) out[pk] = { username: (data[i][0] || '').toString().trim(), fullName: (data[i][2] || '').toString().trim() };
  }
  return out;
}

function misAllLinks_() {
  const sh = getCentralSheet_(SHEETLINKS_SHEET);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const owner = (data[i][0] || '').toString().trim();
    if (owner) out.push({ owner: owner, label: String(data[i][1] || 'Untitled Sheet'), url: data[i][2] || '', tabName: data[i][3] || '' });
  }
  return out;
}

// stages of an FMS sheet: [{ statusCol, plannedCol, actualCol, label }]
function misStages_(data, headers, headerRowIndex) {
  const statusCols = findAllColumns_(headers, 'status');
  if (statusCols.length < 2) return null;
  const stages = statusCols.map(function (sc, idx) {
    const blockStart = idx === 0 ? 0 : statusCols[idx - 1] + 1;
    const labelEnd = idx === statusCols.length - 1 ? headers.length - 1 : statusCols[idx + 1] - 1;
    return {
      blockStart: blockStart, labelEnd: labelEnd, statusCol: sc,
      plannedCol: findColumnInRange_(headers, blockStart, sc, ['planned']),
      actualCol: findColumnInRange_(headers, blockStart, sc, ['actual'])
    };
  });
  const labelScanStart = Math.max(0, headerRowIndex - 8);
  stages.forEach(function (stage, idx) {
    let best = '';
    for (let r = labelScanStart; r < headerRowIndex; r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = stage.blockStart; c <= stage.labelEnd; c++) {
        const v = (row[c] || '').toString().trim();
        if (v && v.length > best.length) best = v;
      }
    }
    stage.label = best || ('Step ' + (idx + 1));
  });
  return stages;
}

// header row of a multi-stage FMS sheet = first row (top 25) with 2+ cells that say exactly "status"
function misFindStageHeader_(data) {
  const lim = Math.min(25, data.length);
  for (let r = 0; r < lim; r++) {
    let n = 0;
    (data[r] || []).forEach(function (v) { if (String(v || '').trim().toLowerCase() === 'status') n++; });
    if (n >= 2) return r;
  }
  return -1;
}

function misTitleOf_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ATTENDANCE_TZ, 'dd/MM/yyyy');
  return String(v == null ? '' : v).trim();
}

// does this person own the row? (Name column may be "Mukesh", "Brijpal", "Brijpal Ji", "Mukesh Kumar" ...)
function misIsMine_(cell, user) {
  const a = normalizeName_(cell);
  if (!a) return false;
  const nu = normalizeName_(user.username), nf = normalizeName_(user.fullName);
  if (a === nu || a === nf) return true;
  if (nu && (a.indexOf(nu) !== -1 || (a.length >= 3 && nu.indexOf(a) !== -1))) return true;
  if (nf && (a.indexOf(nf) !== -1 || (a.length >= 3 && nf.indexOf(a) !== -1))) return true;
  const ft = nu.split(' ')[0];
  return ft.length >= 4 && a.split(' ')[0] === ft;
}

// Reads ONE personal sheet and adds its items to c. Returns '' when fine, or a short reason it was skipped.
function misScanLink_(link, user, fromStr, toStr, todayKey, c, kind) {
  const sheetId = extractSheetId_(link.url);
  if (!sheetId) return 'no valid link';
  const source = (kind === 'fms' ? 'FMS: ' : 'Checklist: ') + link.label;
  const ss = SpreadsheetApp.openById(sheetId);
  const tab = link.tabName ? getTabFuzzy_(ss, link.tabName) : pickDefaultTab_(ss);
  if (!tab) return 'tab not found';
  const data = tab.getDataRange().getValues();
  if (data.length < 1) return '';

  // ----- 1) multi-stage FMS layout: every stage counts, stop at the first unfinished stage -----
  const sh = misFindStageHeader_(data);
  if (sh !== -1) {
    const headers = data[sh];
    const stages = misStages_(data, headers, sh);
    const firstAnchor = stages[0].plannedCol !== -1 ? stages[0].plannedCol : stages[0].statusCol;
    const preEnd = Math.max(0, firstAnchor - 1);

    for (let i = sh + 1; i < data.length; i++) {
      const row = data[i];
      // identity = first non-empty cell before the first stage (Vessel / Order no / Timestamp / Qty...)
      let identity = '';
      for (let k = 0; k <= preEnd; k++) { const t = misTitleOf_(row[k]); if (t) { identity = t; break; } }
      let anyStage = false;
      stages.forEach(function (st) {
        if (misTitleOf_(row[st.statusCol]) || (st.actualCol !== -1 && misTitleOf_(row[st.actualCol]))) anyStage = true;
      });
      if (!identity && !anyStage) continue;                  // empty row
      if (!identity) identity = 'Row ' + (i + 1);
      // prefer a descriptive name (Vessel / Order no.) when it is filled
      const nameCol = findColumnInRange_(headers, 0, preEnd, ['vessel', 'order', 'party', 'supplier', 'client', 'name']);
      if (nameCol !== -1 && misTitleOf_(row[nameCol])) identity = misTitleOf_(row[nameCol]);

      for (let s = 0; s < stages.length; s++) {
        const st = stages[s];
        const done = misIsDone_(row[st.statusCol]);
        const plannedRaw = st.plannedCol !== -1 ? row[st.plannedCol] : '';
        const actualRaw = st.actualCol !== -1 ? row[st.actualCol] : '';
        let planned = misCellKey_(plannedRaw);
        const actual = misCellKey_(actualRaw);
        const title = identity + ' (' + st.label + ')';
        if (done) {
          // a finished stage without a planned date counts as on time, on its actual date
          misCountItem_(c, planned || actual, true, actual, todayKey, fromStr, toStr, title, source,
            planned ? misCellMs_(plannedRaw) : 0, planned ? misCellMs_(actualRaw) : 0);
          continue;
        }
        if (planned) misCountItem_(c, planned, false, '', todayKey, fromStr, toStr, title, source);
        break;                                               // later stages have not started
      }
    }
    return '';
  }

  // ----- 2) Ultimate Checklist (Master tab). No header row: A Name, C Dept, D Task ID, E Freq, F Task, G Planned, H Actual, I Status -----
  let dateRows = 0, sampled = 0;
  for (let i = 0; i < Math.min(data.length, 40); i++) {
    if (!data[i] || data[i].length < 9) continue;
    sampled++;
    if (misCellKey_(data[i][6])) dateRows++;
  }
  const looksUltimate = sampled > 0 && dateRows >= Math.max(1, Math.floor(sampled / 2)) &&
    (/ultimate/i.test(link.label) || /^master$/i.test(String(link.tabName || '').trim()) || data[0].length >= 9);
  if (looksUltimate) {
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const planned = misCellKey_(row[6]);
      if (!planned) continue;                                // header / blank / note rows
      if (!misIsMine_(row[0], user)) continue;
      misCountItem_(c, planned, misIsDone_(row[8]), misCellKey_(row[7]), todayKey, fromStr, toStr,
        String(row[5] || 'Untitled item'), source);
    }
    return '';
  }

  // ----- 3) any other checklist with a normal header row -----
  const hi = findHeaderRow_(data, 15);
  const headers = hi.headers;
  const first = hi.headerRowIndex + 1;
  let dateCol = findColumn_(headers, ['planned', 'plan date', 'due date', 'deadline', 'target date'], []);
  if (dateCol === -1) dateCol = findColumn_(headers, COLUMN_KEYWORDS_.date.keywords, COLUMN_KEYWORDS_.date.exclude);
  const statusCol = findColumn_(headers, COLUMN_KEYWORDS_.status.keywords, COLUMN_KEYWORDS_.status.exclude);
  const taskCol = findColumn_(headers, COLUMN_KEYWORDS_.task.keywords, COLUMN_KEYWORDS_.task.exclude);
  const assignedCol = findColumn_(headers, COLUMN_KEYWORDS_.assigned.keywords, COLUMN_KEYWORDS_.assigned.exclude);
  const actualCol = findColumn_(headers, ['actual date', 'actual', 'done date', 'completed on', 'completed date', 'date of completion'], []);
  if (dateCol === -1) return 'no date / deadline column found';
  if (statusCol === -1) return 'no status column found';

  for (let i = first; i < data.length; i++) {
    const row = data[i];
    const planned = misCellKey_(row[dateCol]);
    if (!planned) continue;
    if (assignedCol !== -1 && !misIsMine_(row[assignedCol], user)) continue;
    const title = taskCol !== -1 ? String(row[taskCol] || 'Untitled item') : 'Untitled item';
    misCountItem_(c, planned, misIsDone_(row[statusCol]), actualCol !== -1 ? misCellKey_(row[actualCol]) : '',
      todayKey, fromStr, toStr, title, source);
  }
  return '';
}

function misPersonalCounts_(personKey, kind, fromStr, toStr, todayKey, ctx) {
  const c = misNewCounts_();
  const user = ctx.users[personKey];
  if (!user) return c;
  const nu = normalizeName_(user.username), nf = normalizeName_(user.fullName);
  ctx.links.forEach(function (link) {
    const o = normalizeName_(link.owner);
    if (!(o === nu || (nf && o === nf))) return;
    const isFms = MIS_FMS_LABEL.test(link.label) || MIS_FMS_LABEL.test(link.tabName || '');
    const isChk = !isFms && MIS_CHECKLIST_LABEL.test(link.label);
    if (kind === 'fms' ? !isFms : !isChk) return;
    try {
      const why = misScanLink_(link, user, fromStr, toStr, todayKey, c, kind);
      if (why) ctx.skipped.push(link.owner + ' - ' + link.label + ': ' + why);
    } catch (e) {
      Logger.log('MIS: error reading ' + link.label + ': ' + e);
      ctx.skipped.push(link.owner + ' - ' + link.label + ': could not be opened');
    }
  });
  return c;
}

// ---------- put it together ----------
function misSummary_(c) {
  const planned = c.onTime + c.late + c.notDone;
  const pctNotDone = planned ? misRound_(c.notDone / planned * 100) : 0;
  const pctNotOnTime = planned ? misRound_((c.late + c.notDone) / planned * 100) : 0;
  return {
    planned: planned, doneOnTime: c.onTime, doneLate: c.late, notDone: c.notDone,
    pctNotDone: pctNotDone, pctNotOnTime: pctNotOnTime,
    score: misRound_(0 - (pctNotDone + pctNotOnTime)), hasData: planned > 0
  };
}

function misMerge_(list) {
  const m = misNewCounts_();
  list.forEach(function (c) {
    m.onTime += c.onTime; m.late += c.late; m.notDone += c.notDone; m.items = m.items.concat(c.items);
  });
  m.items.sort(function (a, b) { return a.deadline < b.deadline ? -1 : a.deadline > b.deadline ? 1 : 0; });
  return m;
}

// type: 'delegation' (default) | 'checklist' | 'fms' | 'all'
// Admin/Manager: everyone (or one person via personKey). Employee: only their own row.
function getMisScores(token, fromStr, toStr, personKey, type, force) {
  const session = verifySession_(token);
  const todayKey = Utilities.formatDate(new Date(), ATTENDANCE_TZ, 'yyyy-MM-dd');
  if (!toStr) toStr = todayKey;
  if (!fromStr) fromStr = toStr.substring(0, 8) + '01';
  attCheckDate_(fromStr); attCheckDate_(toStr);
  if (toStr < fromStr) throw new Error('"To" date must be on or after the "From" date.');
  type = ['delegation', 'checklist', 'fms', 'all'].indexOf(type) === -1 ? 'delegation' : type;

  let keys;
  if (!isAdminRole_(session.role)) {
    if (!session.personKey) throw new Error('No task sheet is linked to your login.');
    keys = [session.personKey];
  } else {
    keys = personKey ? [personKey] : Object.keys(SHEETS_CONFIG);
  }

  const cache = CacheService.getScriptCache();
  const ck = 'mis_v3_' + [type, fromStr, toStr, keys.length === 1 ? keys[0] : 'ALL'].join('_');
  if (!force) {
    try { const hit = cache.get(ck); if (hit) { const o = JSON.parse(hit); o.cached = true; return o; } } catch (e) {}
  }

  const needD = type === 'delegation' || type === 'all';
  const needC = type === 'checklist' || type === 'all';
  const needF = type === 'fms' || type === 'all';
  const ctx = { users: {}, links: [], skipped: [] };
  if (needC || needF) { ctx.users = misUsers_(); ctx.links = misAllLinks_(); }

  const rows = keys.filter(function (k) { return SHEETS_CONFIG[k]; }).map(function (k) {
    const parts = {};
    if (needD) parts.delegation = misDelegationCounts_(k, fromStr, toStr, todayKey);
    if (needC) parts.checklist = misPersonalCounts_(k, 'checklist', fromStr, toStr, todayKey, ctx);
    if (needF) parts.fms = misPersonalCounts_(k, 'fms', fromStr, toStr, todayKey, ctx);
    const merged = misMerge_(Object.keys(parts).map(function (p) { return parts[p]; }));
    const row = misSummary_(merged);
    row.key = k;
    row.name = SHEETS_CONFIG[k].name;
    row.items = merged.items.slice(0, 40);
    if (type === 'all') {
      row.src = {};
      Object.keys(parts).forEach(function (p) {
        const s = misSummary_(parts[p]);
        row.src[p] = { planned: s.planned, score: s.score, hasData: s.hasData };
      });
    }
    return row;
  }).sort(function (a, b) {
    if (a.hasData !== b.hasData) return a.hasData ? -1 : 1;
    return b.score - a.score;
  });

  const withData = rows.filter(function (r) { return r.hasData; });
  const avg = withData.length ? misRound_(withData.reduce(function (s, r) { return s + r.score; }, 0) / withData.length) : 0;
  const skipped = ctx.skipped.filter(function (x, i) { return ctx.skipped.indexOf(x) === i; }).slice(0, 12);
  const out = sanitizeForClient_({ type: type, from: fromStr, to: toStr, today: todayKey, avgScore: avg, rows: rows, skipped: skipped });
  try {
    const json = JSON.stringify(out);
    if (json.length < 90000) cache.put(ck, json, MIS_CACHE_SECONDS);
  } catch (e) {}
  return out;
}

// RUN FROM THE EDITOR to see what the MIS engine finds for each person (no dashboard needed)
function testMisSources() {
  const month = Utilities.formatDate(new Date(), ATTENDANCE_TZ, 'yyyy-MM');
  const ctx = { users: misUsers_(), links: misAllLinks_(), skipped: [] };
  const today = Utilities.formatDate(new Date(), ATTENDANCE_TZ, 'yyyy-MM-dd');
  Object.keys(SHEETS_CONFIG).forEach(function (k) {
    const d = misSummary_(misDelegationCounts_(k, month + '-01', today, today));
    const c = misSummary_(misPersonalCounts_(k, 'checklist', month + '-01', today, today, ctx));
    const f = misSummary_(misPersonalCounts_(k, 'fms', month + '-01', today, today, ctx));
    Logger.log(SHEETS_CONFIG[k].name + ' | Delegation ' + d.score + ' (' + d.planned + ') | Checklist ' + c.score + ' (' + c.planned + ') | FMS ' + f.score + ' (' + f.planned + ')');
  });
  Logger.log('Skipped sheets: ' + (ctx.skipped.join(' || ') || 'none'));
}
