// ============================================================
// RECURRING TASKS + FOLLOW-UPS + WEEKLY STALE DIGEST
// ------------------------------------------------------------
// NEW FILE in your Apps Script project (File > + > Script > name it "RecurringTasks").
// Needs the updated Code.gs (appendTaskRow_, 13-column tasks) and whatsapp_reminders.gs.
//
// Tab "Recurring" is created automatically in the central "User login" sheet:
// A RuleId | B TaskName | C PersonKey | D AssignedTo | E Description | F Priority | G Repeat | H Day |
// I StartDate | J EndDate | K DueInDays | L Reminder (JSON) | M Active | N LastCreated | O Created
//
// Repeat / Day:
//   Daily     -> Day empty
//   Weekdays  -> Mon to Sat, Day empty (Sunday skipped)
//   Weekly    -> Day = Mon, Tue, Wed, Thu, Fri, Sat or Sun
//   Monthly   -> Day = 1..31 or Last   (31 in a short month runs on the last day)
//   Yearly    -> Day = MM-dd  (e.g. 03-15)
//
// RUN ONCE from the editor:  createDailyTaskEngineTrigger()
// ============================================================

const RECURRING_SHEET = 'Recurring';
const REC_HEADERS = ['RuleId', 'TaskName', 'PersonKey', 'AssignedTo', 'Description', 'Priority', 'Repeat', 'Day',
  'StartDate', 'EndDate', 'DueInDays', 'Reminder', 'Active', 'LastCreated', 'Created'];
const REC_REPEATS = ['Daily', 'Weekdays', 'Weekly', 'Monthly', 'Yearly'];
const REC_DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const STALE_DAYS = 7;          // an open task with no update for this many days is "stale"

// ---------- sheet ----------

function recSheet_(create) {
  let sh = getCentralSheet_(RECURRING_SHEET);
  if (sh || !create) return sh;
  sh = SpreadsheetApp.openById(SHEET_ID).insertSheet(RECURRING_SHEET);
  sh.appendRow(REC_HEADERS);
  sh.getRange(1, 1, 1, REC_HEADERS.length).setFontWeight('bold');
  sh.getRange('H:J').setNumberFormat('@');     // Day / Start / End stay plain text
  sh.getRange('N:N').setNumberFormat('@');
  return sh;
}

function recText_(v) {                         // cell -> plain text ('yyyy-MM-dd' for real dates)
  if (v instanceof Date) return Utilities.formatDate(v, ATTENDANCE_TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}

function recRowToRule_(r, rowNo) {
  let reminder = null;
  try { reminder = r[11] ? JSON.parse(r[11]) : null; } catch (e) { reminder = null; }
  return {
    id: String(r[0]), row: rowNo, taskName: String(r[1] || ''), personKey: String(r[2] || ''),
    assignedTo: String(r[3] || ''), description: String(r[4] || ''), priority: String(r[5] || 'Medium'),
    repeat: String(r[6] || 'Daily'), day: recText_(r[7]), startDate: recText_(r[8]), endDate: recText_(r[9]),
    dueInDays: Number(r[10]) || 0, reminder: reminder, active: String(r[12] || 'Yes').toLowerCase() !== 'no',
    lastCreated: recText_(r[13])
  };
}

function recReadAll_() {
  const sh = recSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, REC_HEADERS.length).getValues();
  const out = [];
  rows.forEach(function (r, i) { if (r[0]) out.push(recRowToRule_(r, i + 2)); });
  return out;
}

// ---------- date maths (all 'yyyy-MM-dd' keys, India calendar) ----------

function recAddDays_(key, n) {
  const p = key.split('-').map(Number);
  return Utilities.formatDate(new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)), 'UTC', 'yyyy-MM-dd');
}

function recDow_(key) {
  const p = key.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
}

function recDaysInMonth_(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }   // m = 1..12

// Does this rule fall on this day?
function recIsDueOn_(rule, key) {
  if (rule.startDate && key < rule.startDate) return false;
  if (rule.endDate && key > rule.endDate) return false;
  const p = key.split('-').map(Number), dom = p[2], dow = recDow_(key);
  switch (rule.repeat) {
    case 'Daily': return true;
    case 'Weekdays': return dow !== 0;
    case 'Weekly': return REC_DOW[dow].toLowerCase() === String(rule.day || '').toLowerCase().slice(0, 3);
    case 'Monthly': {
      const dim = recDaysInMonth_(p[0], p[1]);
      if (String(rule.day).toLowerCase() === 'last') return dom === dim;
      return dom === Math.min(Number(rule.day) || 1, dim);
    }
    case 'Yearly': {
      const m = String(rule.day || '').match(/^(\d{1,2})-(\d{1,2})$/);
      if (!m) return false;
      const mm = +m[1], dd = +m[2];
      if (p[1] !== mm) return false;
      return dom === Math.min(dd, recDaysInMonth_(p[0], mm));
    }
  }
  return false;
}

// next dates this rule will create a task (for the dashboard preview)
function recNextDates_(rule, fromKey, count) {
  const out = [];
  let k = fromKey;
  for (let i = 0; i < 800 && out.length < count; i++) {
    if (recIsDueOn_(rule, k)) out.push(k);
    k = recAddDays_(k, 1);
  }
  return out;
}

// ---------- validate + save (admin only) ----------

function recClean_(raw) {
  raw = raw || {};
  const rule = {
    taskName: String(raw.taskName || '').trim(),
    personKey: String(raw.personKey || '').trim(),
    assignedTo: String(raw.assignedTo || '').trim(),
    description: String(raw.description || '').trim(),
    priority: ['High', 'Medium', 'Low'].indexOf(raw.priority) === -1 ? 'Medium' : raw.priority,
    repeat: REC_REPEATS.indexOf(raw.repeat) === -1 ? '' : raw.repeat,
    day: String(raw.day == null ? '' : raw.day).trim(),
    startDate: waDateKey_(raw.startDate),
    endDate: waDateKey_(raw.endDate),
    dueInDays: Math.max(0, Math.min(365, Number(raw.dueInDays) || 0))
  };
  if (!rule.taskName) throw new Error('Task name is required.');
  if (!SHEETS_CONFIG[rule.personKey]) throw new Error('Pick the employee this task is for.');
  if (!rule.repeat) throw new Error('Pick how often it repeats.');
  if (!rule.startDate) throw new Error('Start date is required.');
  if (rule.endDate && rule.endDate < rule.startDate) throw new Error('End date is before the start date.');
  if (!rule.assignedTo) rule.assignedTo = SHEETS_CONFIG[rule.personKey].name;
  if (rule.repeat === 'Weekly') {
    const d = REC_DOW.filter(function (x) { return x.toLowerCase() === rule.day.toLowerCase().slice(0, 3); })[0];
    if (!d) throw new Error('Pick the weekday.');
    rule.day = d;
  } else if (rule.repeat === 'Monthly') {
    if (rule.day.toLowerCase() === 'last') rule.day = 'Last';
    else {
      const n = Number(rule.day);
      if (!(n >= 1 && n <= 31)) throw new Error('Month day must be 1 to 31, or Last.');
      rule.day = String(Math.floor(n));
    }
  } else if (rule.repeat === 'Yearly') {
    const m = rule.day.match(/^(\d{1,2})-(\d{1,2})$/);
    if (!m || +m[1] < 1 || +m[1] > 12 || +m[2] < 1 || +m[2] > 31) throw new Error('Yearly day must look like 03-15 (month-day).');
    rule.day = ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  } else rule.day = '';
  const plan = waNormalizePlan_(raw.reminder);
  rule.reminderJson = plan ? JSON.stringify(plan) : '';
  return rule;
}

function getRecurringRules(token) {
  assertIsAdmin_(verifySession_(token));
  const today = waTodayKey_();
  const rules = recReadAll_().map(function (r) {
    r.person = SHEETS_CONFIG[r.personKey] ? SHEETS_CONFIG[r.personKey].name : r.personKey;
    r.next = r.active ? recNextDates_(r, today, 3) : [];
    return r;
  });
  return sanitizeForClient_(rules);
}

function saveRecurringRule(token, raw) {
  assertIsAdmin_(verifySession_(token));
  const rule = recClean_(raw);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = recSheet_(true);
    const vals = [rule.taskName, rule.personKey, rule.assignedTo, rule.description, rule.priority, rule.repeat, rule.day,
      rule.startDate, rule.endDate, rule.dueInDays, rule.reminderJson];
    const id = raw && raw.id ? String(raw.id) : '';
    let rowNo = 0;
    if (id) {
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
      for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) { rowNo = i + 2; break; }
      if (!rowNo) throw new Error('That rule no longer exists.');
      sh.getRange(rowNo, 2, 1, 11).setValues([vals]);
    } else {
      const newId = 'R' + Date.now().toString(36).toUpperCase();
      sh.appendRow([newId].concat(vals).concat(['Yes', '', new Date()]));
      rowNo = sh.getLastRow();
    }
    sh.getRange(rowNo, 8, 1, 3).setNumberFormat('@').setValues([[rule.day, rule.startDate, rule.endDate]]);
  } finally { lock.releaseLock(); }
  return { success: true, message: id_(raw) ? 'Repeat rule updated.' : 'Repeat rule saved.' };
}

function id_(raw) { return !!(raw && raw.id); }

function recFindRow_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function setRecurringActive(token, id, active) {
  assertIsAdmin_(verifySession_(token));
  const sh = recSheet_(false);
  const row = sh ? recFindRow_(sh, id) : 0;
  if (!row) throw new Error('Rule not found.');
  sh.getRange(row, 13).setValue(active ? 'Yes' : 'No');
  return { success: true, message: active ? 'Rule resumed.' : 'Rule paused.' };
}

function deleteRecurringRule(token, id) {
  assertIsAdmin_(verifySession_(token));
  const sh = recSheet_(false);
  const row = sh ? recFindRow_(sh, id) : 0;
  if (!row) throw new Error('Rule not found.');
  sh.deleteRow(row);
  return { success: true, message: 'Rule deleted. Tasks already created stay as they are.' };
}

// ---------- the engine ----------

// Creates today's tasks for every active rule. Returns the number of tasks created.
function recRunRules_() {
  const sh = recSheet_(false);
  if (!sh) return 0;
  const today = waTodayKey_();
  let created = 0;
  recReadAll_().forEach(function (rule) {
    if (!rule.active) return;
    try {
      // look back a few days in case the trigger was missed; create ONE task for the latest due day
      let dueDay = '';
      for (let back = 0; back <= 3; back++) {
        const k = recAddDays_(today, -back);
        if (rule.lastCreated && k <= rule.lastCreated) break;   // already created for this day
        if (recIsDueOn_(rule, k)) { dueDay = k; break; }
      }
      if (!dueDay) return;
      const config = SHEETS_CONFIG[rule.personKey];
      if (!config) return;
      const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheetName);
      if (!sheet) return;

      const now = new Date();
      const taskData = {
        taskName: rule.taskName, assignedTo: rule.assignedTo,
        description: (rule.description ? rule.description + '\n' : '') + '(Repeat task - ' + rule.repeat + ')',
        deadline: recAddDays_(dueDay, rule.dueInDays), priority: rule.priority,
        reminder: rule.reminder, notes: 'Auto-created from repeat rule ' + rule.id
      };
      appendTaskRow_(sheet, now, taskData);
      saveReminderPlan_(rule.personKey, taskData, now.getTime());
      sh.getRange(rule.row, 14).setNumberFormat('@').setValue(dueDay);
      created++;

      try { sendTaskAssignedWhatsAppCloud_(rule.personKey, taskData); } catch (e) { Logger.log('rec WA: ' + e); }
      try { createNotification_(rule.personKey, 'New repeat task', buildTaskAssignedNotifText_(taskData)); } catch (e) { Logger.log('rec notif: ' + e); }
    } catch (e) {
      Logger.log('Recurring rule ' + rule.id + ' failed: ' + e);
    }
  });
  if (created) cacheRemoveBig_(TASKS_CACHE_KEY);
  return created;
}

// ---------- follow-ups + stale tasks ----------

function recAgeDays_(t) {
  const base = waDateKey_(t.LastUpdate) || waDateKey_(t.Timestamp);
  return base ? waDayDiff_(waTodayKey_(), base) : 0;
}

// Sends WhatsApp for follow-ups that are due today. Returns the number of follow-ups handled.
function recRunFollowUps_() {
  const today = waTodayKey_();
  let n = 0;
  Object.keys(SHEETS_CONFIG).forEach(function (pk) {
    let tasks = [];
    try { tasks = getTasksByPersonRaw_(pk); } catch (e) { Logger.log('followup read ' + pk + ': ' + e); }
    tasks.forEach(function (t) {
      if (waIsDone_(t.Status) || !t.FollowUp) return;
      if (waDateKey_(t.FollowUp) !== today) return;
      const res = waSendTaskAlert_(pk, t.TaskName, waDateKey_(t.Deadline) || today, 'FOLLOW-UP today', true);
      createNotification_(pk, 'Follow-up today', '"' + t.TaskName + '" needs a follow-up today.');
      n++;
      if (!res.sent) Logger.log('Follow-up WA not delivered for ' + t.TaskName + ': ' + res.errors.join(' | '));
    });
  });
  return n;
}

// One WhatsApp + in-app note to the admin number with no employee copy
function recAdminNote_(text) {
  const custom = getScriptProp_('WA_REMINDER_TEMPLATE', '');
  const tpl = custom || WA_TEMPLATE_NAME;
  const params = custom ? ['WEEKLY DIGEST', text, 'All team', waDmy_(waTodayKey_())] : [text, 'All team', waDmy_(waTodayKey_())];
  return WA_ADMIN_NUMBER ? waSendTemplate_(WA_ADMIN_NUMBER, tpl, WA_TEMPLATE_LANG, params) : { ok: false };
}

// Monday digest: how many open tasks are stale / overdue, and the oldest few
function recWeeklyDigest_() {
  const stale = [];
  let overdue = 0, open = 0;
  Object.keys(SHEETS_CONFIG).forEach(function (pk) {
    let tasks = [];
    try { tasks = getTasksByPersonRaw_(pk); } catch (e) { return; }
    tasks.forEach(function (t) {
      if (waIsDone_(t.Status)) return;
      open++;
      if (waOverdueInfo_(t)) overdue++;
      const age = recAgeDays_(t);
      if (age >= STALE_DAYS) stale.push({ owner: t.Owner, name: t.TaskName, age: age });
    });
  });
  stale.sort(function (a, b) { return b.age - a.age; });
  const top = stale.slice(0, 3).map(function (s) { return s.name + ' (' + s.owner + ', ' + s.age + 'd)'; }).join('; ');
  const msg = open + ' open, ' + overdue + ' overdue, ' + stale.length + ' not updated for ' + STALE_DAYS + '+ days.' + (top ? ' Oldest: ' + top : '');
  createNotification_(ADMIN_BROADCAST_KEY, 'Weekly digest', msg);
  recAdminNote_(msg);
  return msg;
}

// ---------- one trigger runs everything ----------

function runDailyTaskEngine() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const made = recRunRules_();
    const fu = recRunFollowUps_();
    let digest = '';
    if (recDow_(waTodayKey_()) === 1) digest = recWeeklyDigest_();       // Monday
    Logger.log('Daily engine: ' + made + ' repeat task(s) created, ' + fu + ' follow-up(s) sent. ' + digest);
  } finally { lock.releaseLock(); }
}

// RUN THIS ONCE from the editor: creates the daily trigger (about 7 AM India time)
function createDailyTaskEngineTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runDailyTaskEngine') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyTaskEngine').timeBased().everyDays(1).atHour(7).inTimezone(ATTENDANCE_TZ).create();
  Logger.log('Created: runDailyTaskEngine daily at 7 AM.');
}

// Button in the dashboard: create today's repeat tasks right now (safe to press twice - no duplicates)
function runRecurringNow(token) {
  assertIsAdmin_(verifySession_(token));
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Another run is in progress. Try again in a minute.');
  try {
    const made = recRunRules_();
    return { success: true, message: made ? made + ' repeat task(s) created.' : 'Nothing due right now - no new tasks created.' };
  } finally { lock.releaseLock(); }
}

// ---------- test from the editor ----------
function testRecurringDates() {
  const rule = { repeat: 'Monthly', day: '31', startDate: '2026-01-01', endDate: '' };
  Logger.log(JSON.stringify(recNextDates_(rule, '2026-10-08', 4)));   // expect 10-31, 11-30, 12-31, 01-31
  Logger.log(JSON.stringify(recNextDates_({ repeat: 'Weekly', day: 'Mon', startDate: '2026-01-01', endDate: '' }, '2026-10-08', 3)));
  Logger.log(JSON.stringify(recNextDates_({ repeat: 'Yearly', day: '02-29', startDate: '2026-01-01', endDate: '' }, '2026-10-08', 2)));
}