// ============================================================
// WHATSAPP REMINDERS + OVERDUE ALERTS
// ------------------------------------------------------------
// NEW FILE in your Apps Script project (File > + > Script > name it "WhatsAppReminders").
// Keep your existing code as it is. Only ONE line is added in addTaskToPerson()
// (see the setup steps). This file reuses what you already have:
//   SHEET_ID, getCentralSheet_, SHEETS_CONFIG, DATA_START_ROW, DONE_KEYWORDS,
//   getScriptProp_, getEmployeeWhatsApp_, getTasksByPersonRaw_, verifySession_,
//   assertIsAdmin_, sanitizeForClient_, WA_PHONE_NUMBER_ID, WA_API_VERSION,
//   WA_ADMIN_NUMBER, WA_TEMPLATE_NAME, WA_TEMPLATE_LANG, ATTENDANCE_TZ
// ============================================================

const REMINDERS_SHEET = 'Reminders';
// Columns of the Reminders tab (1-based):
// A Id | B Created | C PersonKey | D TaskName | E CreatedMs | F Deadline (yyyy-MM-dd) | G Plan (JSON) | H SentLog | I State
// State: ACTIVE, FINISHED (all reminders sent), DONE (task completed), STOPPED (task deleted / not found)

// ---------- small helpers ----------

function waToken_() { return getScriptProp_('WA_ACCESS_TOKEN', ''); }

function waIsDone_(status) {
  return DONE_KEYWORDS.indexOf(String(status || '').trim().toLowerCase()) !== -1;
}

// Any date-ish value -> 'yyyy-MM-dd' (India date). Accepts Date, 'yyyy-MM-dd', 'dd/MM/yyyy[ HH:mm]'.
function waDateKey_(v) {
  if (!v) return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, ATTENDANCE_TZ, 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function waDmy_(key) {            // 'yyyy-MM-dd' -> 'dd/MM/yyyy'
  return key ? key.split('-').reverse().join('/') : 'Not set';
}

// whole days between two 'yyyy-MM-dd' keys (a - b)
function waDayDiff_(a, b) {
  const pa = a.split('-').map(Number), pb = b.split('-').map(Number);
  return Math.round((Date.UTC(pa[0], pa[1] - 1, pa[2]) - Date.UTC(pb[0], pb[1] - 1, pb[2])) / 86400000);
}

function waTodayKey_() { return Utilities.formatDate(new Date(), ATTENDANCE_TZ, 'yyyy-MM-dd'); }

function waLabel_(diff) {         // diff = today - deadline (days)
  if (diff < 0) return 'Reminder: due in ' + (-diff) + ' day' + (diff === -1 ? '' : 's');
  if (diff === 0) return 'DUE TODAY';
  return 'OVERDUE by ' + diff + ' day' + (diff === 1 ? '' : 's');
}

function waClean_(s) {            // WhatsApp template params: no new lines / tabs / long spaces
  return String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').replace(/ {4,}/g, '   ').trim().slice(0, 300) || '-';
}

function waNumber_(n) {
  n = String(n || '').replace(/[^\d]/g, '');
  if (n.length === 10) n = '91' + n;            // add India code when only 10 digits are saved
  return n;
}

// ---------- sending ----------

function waSendTemplate_(to, templateName, lang, params) {
  const token = waToken_();
  to = waNumber_(to);
  if (!to) return { ok: false, error: 'No number' };
  if (!token) return { ok: false, error: 'WA_ACCESS_TOKEN not set in Script Properties' };
  const url = 'https://graph.facebook.com/' + WA_API_VERSION + '/' + WA_PHONE_NUMBER_ID + '/messages';
  const payload = {
    messaging_product: 'whatsapp',
    to: to,
    type: 'template',
    template: {
      name: templateName,
      language: { code: lang },
      components: [{ type: 'body', parameters: params.map(function (p) { return { type: 'text', text: waClean_(p) }; }) }]
    }
  };
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) return { ok: true };
    Logger.log('WA FAILED ' + to + ' (' + res.getResponseCode() + '): ' + res.getContentText());
    return { ok: false, error: res.getContentText().slice(0, 300) };
  } catch (e) {
    Logger.log('WA error ' + to + ': ' + e);
    return { ok: false, error: String(e) };
  }
}

// Sends one task alert to the employee (+ admin copy). Returns { sent: n, errors: [..] }
function waSendTaskAlert_(personKey, taskName, deadlineKey, label, alsoAdmin) {
  const config = SHEETS_CONFIG[personKey];
  const person = config ? config.name : personKey;
  const dl = waDmy_(deadlineKey);

  // Best: your own approved 4-variable template (set WA_REMINDER_TEMPLATE in Script Properties).
  // Fallback: your existing template task_reminder1 (3 variables), so it works from day one.
  const custom = getScriptProp_('WA_REMINDER_TEMPLATE', '');
  const tpl = custom || WA_TEMPLATE_NAME;
  const params = custom
    ? [label, taskName, person, dl]
    : [taskName + ' - ' + label, person, dl];

  const targets = [];
  const emp = getEmployeeWhatsApp_(personKey);
  if (emp) targets.push(emp);
  if (alsoAdmin && WA_ADMIN_NUMBER) targets.push(WA_ADMIN_NUMBER);

  let sent = 0; const errors = [];
  const seen = {};
  targets.forEach(function (t) {
    const k = waNumber_(t);
    if (!k || seen[k]) return;
    seen[k] = true;
    const r = waSendTemplate_(k, tpl, WA_TEMPLATE_LANG, params);
    if (r.ok) sent++; else errors.push(r.error);
  });
  if (!emp) errors.push('No WhatsApp number for ' + person + ' in Users sheet (column G)');
  return { sent: sent, errors: errors };
}

// ---------- finding the task row again ----------

function waFindTask_(personKey, taskName, createdMs) {
  const config = SHEETS_CONFIG[personKey];
  if (!config) return null;
  const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheetName);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return null;
  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 10).getValues();
  const want = String(taskName || '').trim();
  let byName = [];
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1] || '').trim() !== want) continue;
    const a = data[i][0];
    const ms = a instanceof Date ? a.getTime() : new Date(a).getTime();
    const hit = { rowNumber: i + DATA_START_ROW, status: data[i][6], deadlineKey: waDateKey_(data[i][4]), priority: data[i][5] };
    if (createdMs && Math.abs(ms - createdMs) < 5000) return hit;
    byName.push(hit);
  }
  const open = byName.filter(function (h) { return !waIsDone_(h.status); });
  return open.length === 1 ? open[0] : null;     // fallback only when it is unambiguous
}

// ---------- plan: validate, save, build time slots ----------

function waNormalizePlan_(raw) {
  if (!raw || raw.enabled === false) return null;
  const timeOk = /^([01]\d|2[0-3]):[0-5]\d$/;
  let times = (raw.times || []).map(String).filter(function (t) { return timeOk.test(t); });
  times = times.filter(function (t, i) { return times.indexOf(t) === i; }).sort().slice(0, 4);
  if (!times.length) times = ['09:00'];
  let before = (raw.before || []).map(Number).filter(function (n) { return n >= 1 && n <= 30; });
  before = before.filter(function (n, i) { return before.indexOf(n) === i; }).sort(function (a, b) { return b - a; }).slice(0, 5);
  const plan = {
    before: before,
    onDay: raw.onDay !== false,
    overdueDays: Math.max(0, Math.min(14, Number(raw.overdueDays) || 0)),
    times: times,
    notifyAdmin: raw.notifyAdmin !== false
  };
  if (!plan.before.length && !plan.onDay && !plan.overdueDays) return null;
  return plan;
}

// Every planned send time as { key:'yyyy-MM-dd HH:mm', ms }
function waBuildSlots_(deadlineKey, plan, afterMs) {
  const p = deadlineKey.split('-').map(Number);
  const offsets = [];
  plan.before.forEach(function (n) { offsets.push(-n); });
  if (plan.onDay) offsets.push(0);
  for (let d = 1; d <= plan.overdueDays; d++) offsets.push(d);
  const slots = [];
  offsets.forEach(function (off) {
    const day = Utilities.formatDate(new Date(Date.UTC(p[0], p[1] - 1, p[2] + off)), 'UTC', 'yyyy-MM-dd');
    plan.times.forEach(function (t) {
      const ms = Utilities.parseDate(day + ' ' + t, ATTENDANCE_TZ, 'yyyy-MM-dd HH:mm').getTime();
      if (!afterMs || ms > afterMs) slots.push({ key: day + ' ' + t, ms: ms });
    });
  });
  slots.sort(function (a, b) { return a.ms - b.ms; });
  return slots.slice(0, 40);
}

function waGetRemindersSheet_(create) {
  let sh = getCentralSheet_(REMINDERS_SHEET);
  if (sh || !create) return sh;
  sh = SpreadsheetApp.openById(SHEET_ID).insertSheet(REMINDERS_SHEET);
  sh.appendRow(['Id', 'Created', 'PersonKey', 'TaskName', 'CreatedMs', 'Deadline', 'Plan', 'SentLog', 'State']);
  sh.getRange(1, 1, 1, 9).setFontWeight('bold');
  sh.getRange('F:F').setNumberFormat('@');     // keep the deadline as plain text
  sh.getRange('H:H').setNumberFormat('@');
  return sh;
}

// >>> CALLED FROM addTaskToPerson() — saves the reminder plan chosen while adding the task
function saveReminderPlan_(personKey, taskData, createdMs) {
  try {
    const plan = waNormalizePlan_(taskData && taskData.reminder);
    if (!plan) return;
    const deadlineKey = waDateKey_(taskData.deadline);
    if (!deadlineKey) return;                    // reminders need a deadline
    const sh = waGetRemindersSheet_(true);
    const id = Date.now() + '_' + Math.floor(Math.random() * 1000);
    sh.appendRow([id, new Date(), personKey, taskData.taskName || '', createdMs, deadlineKey, JSON.stringify(plan), '', 'ACTIVE']);
    const r = sh.getLastRow();
    sh.getRange(r, 6).setNumberFormat('@').setValue(deadlineKey);
  } catch (e) {
    Logger.log('saveReminderPlan_ error: ' + e);
  }
}

// ---------- the engine (runs every 15 minutes by trigger) ----------

function runReminderEngine() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const sh = waGetRemindersSheet_(false);
    if (!sh || sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    const now = Date.now();
    const today = waTodayKey_();

    rows.forEach(function (r, i) {
      const rowNo = i + 2;
      if (String(r[8]) !== 'ACTIVE') return;
      try {
        const personKey = String(r[2]), taskName = String(r[3]), createdMs = Number(r[4]);
        const plan = JSON.parse(r[6]);
        const sent = r[7] ? String(r[7]).split(',') : [];

        // 1) Cheap check using only this tab: is anything due right now?
        let deadlineKey = waDateKey_(r[5]);
        let slots = waBuildSlots_(deadlineKey, plan, createdMs);
        let remaining = slots.filter(function (s) { return sent.indexOf(s.key) === -1; });
        if (!remaining.length) { sh.getRange(rowNo, 9).setValue('FINISHED'); return; }
        if (!remaining.some(function (s) { return s.ms <= now; })) return;

        // 2) Something is due -> open the task sheet (handles Done / deleted / changed deadline)
        const task = waFindTask_(personKey, taskName, createdMs);
        if (!task) { sh.getRange(rowNo, 9).setValue('STOPPED'); return; }
        if (waIsDone_(task.status)) { sh.getRange(rowNo, 9).setValue('DONE'); return; }
        if (task.deadlineKey && task.deadlineKey !== deadlineKey) {          // deadline was edited in the sheet
          deadlineKey = task.deadlineKey;
          sh.getRange(rowNo, 6).setNumberFormat('@').setValue(deadlineKey);
          slots = waBuildSlots_(deadlineKey, plan, createdMs);
          remaining = slots.filter(function (s) { return sent.indexOf(s.key) === -1; });
          if (!remaining.some(function (s) { return s.ms <= now; })) return;
        }

        // 3) Send ONE message even if several slots piled up, then mark them all as sent
        const dueSlots = remaining.filter(function (s) { return s.ms <= now; });
        const label = waLabel_(waDayDiff_(today, deadlineKey));
        const res = waSendTaskAlert_(personKey, taskName, deadlineKey, label, plan.notifyAdmin);
        if (res.sent > 0) {
          dueSlots.forEach(function (s) { sent.push(s.key); });
          sh.getRange(rowNo, 8).setNumberFormat('@').setValue(sent.join(','));
        } else {
          Logger.log('Reminder not delivered for "' + taskName + '": ' + res.errors.join(' | '));
        }
      } catch (e) {
        Logger.log('runReminderEngine row ' + rowNo + ' error: ' + e);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

// RUN THIS ONCE from the editor to start the engine
function createReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runReminderEngine') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runReminderEngine').timeBased().everyMinutes(15).create();
  Logger.log('Reminder trigger created: runReminderEngine every 15 minutes.');
}

// ---------- manual: send overdue tasks on WhatsApp (buttons in the dashboard) ----------

function waOverdueInfo_(t) {                    // t from getTasksByPersonRaw_
  if (!t || waIsDone_(t.Status) || !t.Deadline) return null;
  const key = waDateKey_(t.Deadline);
  if (!key) return null;
  const diff = waDayDiff_(waTodayKey_(), key);
  return diff > 0 ? { key: key, diff: diff } : null;
}

function sendOverdueWhatsApp(token, personKey, rowNumber) {
  const session = verifySession_(token);
  assertIsAdmin_(session);
  if (!waToken_()) return { success: false, message: 'WhatsApp is not set up yet: WA_ACCESS_TOKEN is missing in Script Properties.' };

  const t = getTasksByPersonRaw_(personKey).filter(function (x) { return x.rowNumber === Number(rowNumber); })[0];
  if (!t) return { success: false, message: 'Task not found (the sheet may have changed). Refresh and try again.' };
  const info = waOverdueInfo_(t);
  if (!info) return { success: false, message: 'This task is not overdue.' };

  const cache = CacheService.getScriptCache();
  const dk = 'manualWa_' + personKey + '_' + rowNumber;
  if (cache.get(dk)) return { success: false, message: 'Already sent in the last 10 minutes.' };

  const res = waSendTaskAlert_(personKey, t.TaskName, info.key, waLabel_(info.diff), true);
  if (res.sent > 0) {
    cache.put(dk, '1', 600);
    return { success: true, message: 'WhatsApp sent (' + res.sent + ' number' + (res.sent > 1 ? 's' : '') + ').' };
  }
  return { success: false, message: 'Not sent. ' + (res.errors[0] || 'Unknown error') };
}

// personKey optional. Without it: every employee.
function sendAllOverdueWhatsApp(token, personKey) {
  const session = verifySession_(token);
  assertIsAdmin_(session);
  if (!waToken_()) return { success: false, message: 'WhatsApp is not set up yet: WA_ACCESS_TOKEN is missing in Script Properties.' };

  const cache = CacheService.getScriptCache();
  const keys = personKey ? [personKey] : Object.keys(SHEETS_CONFIG);
  let tasks = 0, sentMsgs = 0, skipped = 0, failed = 0;

  keys.forEach(function (pk) {
    getTasksByPersonRaw_(pk).forEach(function (t) {
      const info = waOverdueInfo_(t);
      if (!info) return;
      if (tasks >= 40) return;                    // safety limit per click
      const dk = 'manualWa_' + pk + '_' + t.rowNumber;
      if (cache.get(dk)) { skipped++; return; }
      tasks++;
      const res = waSendTaskAlert_(pk, t.TaskName, info.key, waLabel_(info.diff), true);
      if (res.sent > 0) { sentMsgs += res.sent; cache.put(dk, '1', 600); } else failed++;
    });
  });

  return {
    success: failed === 0,
    message: tasks + ' overdue task(s) processed, ' + sentMsgs + ' message(s) sent' +
      (skipped ? ', ' + skipped + ' skipped (sent in the last 10 min)' : '') +
      (failed ? ', ' + failed + ' failed (check Executions log / WhatsApp numbers)' : '') + '.'
  };
}

// ---------- tests / debug (run from the Apps Script editor) ----------

function testReminderSlots() {
  const plan = waNormalizePlan_({ before: [3, 1], onDay: true, overdueDays: 2, times: ['09:00', '16:30'], notifyAdmin: true });
  Logger.log(JSON.stringify(plan));
  waBuildSlots_('2026-10-10', plan, 0).forEach(function (s) { Logger.log(s.key); });
}

function testOverdueWhatsApp() {      // sends a sample alert to the admin number
  const r = waSendTemplate_(WA_ADMIN_NUMBER,
    getScriptProp_('WA_REMINDER_TEMPLATE', '') || WA_TEMPLATE_NAME, WA_TEMPLATE_LANG,
    getScriptProp_('WA_REMINDER_TEMPLATE', '') ? ['OVERDUE by 2 days', 'Sample task', 'Test User', '04/10/2026']
                                                : ['Sample task - OVERDUE by 2 days', 'Test User', '04/10/2026']);
  Logger.log(JSON.stringify(r));
}

function debugReminders() {
  const sh = waGetRemindersSheet_(false);
  if (!sh) { Logger.log('No Reminders tab yet. It is created when the first task with reminders is added.'); return; }
  Logger.log('Rows: ' + Math.max(0, sh.getLastRow() - 1));
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    Logger.log(r[2] + ' | ' + r[3] + ' | deadline ' + r[5] + ' | state ' + r[8] + ' | sent ' + r[7]);
  });
  Logger.log('Triggers: ' + ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }).join(', '));
  Logger.log('Token set: ' + (waToken_() ? 'YES' : 'NO'));
}