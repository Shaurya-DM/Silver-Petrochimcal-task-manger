// ============================================================
// AI QUICK — instant answers WITHOUT Gemini (1-2 seconds instead of 10-30)
// ------------------------------------------------------------
// NEW FILE in your Apps Script project (File > + > Script > name it "AIQuick").
// Reuses (from AIAssistant / Code): verifySession_, aiAllTasks_, aiTaskRow_, aiToolReport_,
//         waIsDone_, waTodayKey_
//
// The buttons Overview / Overdue / MIS / Stale / "Kaun kya kar raha hai" and the new search box
// call aiQuick(token, kind, q). The numbers come straight from the (already cached) task list,
// so there is no AI wait and no chance of the AI making up a number.
//
// kind = overview | overdue | stale | whodoes | mis | find   (find uses q = the search words)
// Employees only ever get their own tasks (aiAllTasks_ already limits that).
// ============================================================

function aiQuick(token, kind, q) {
  const t0 = Date.now();
  const session = verifySession_(token);
  let reply;
  try {
    reply = aiQuickBuild_(token, session, String(kind || ''), String(q || ''));
  } catch (e) {
    reply = '⚠️ ' + (e.message || e);
  }
  return { reply: reply, ms: Date.now() - t0 };
}

function aiQItems_(token, session) {
  return aiAllTasks_(token, session).map(function (t) {
    const r = aiTaskRow_(t);
    return {
      r: r,
      done: waIsDone_(t.Status),
      hay: [t.Owner, t.TaskName, t.Description, t.Remark, t.Notes, t.Status, t.Priority].join(' ').toLowerCase()
    };
  });
}

function aiQLine_(r) {
  return '- ' + r.task + ' · due ' + (r.deadline || '-') + ' · ' + r.priority + ' · ' + r.status +
    (r.progress ? ' ' + r.progress + '%' : '') +
    (r.overdue_days > 0 ? ' · **' + r.overdue_days + ' din late**' : '');
}

function aiQLate_(r) {
  return '- **' + r.person + '** — ' + r.task + ' · due ' + (r.deadline || '-') + ' · ' + r.priority +
    (r.overdue_days > 0 ? ' · **' + r.overdue_days + ' din late**' : '') +
    (r.days_since_update !== null && r.days_since_update >= 7 ? ' · ' + r.days_since_update + ' din se update nahi' : '');
}

function aiQuickBuild_(token, session, kind, q) {
  if (kind === 'mis') return aiQMis_(token, session);

  const items = aiQItems_(token, session);
  const open = items.filter(function (x) { return !x.done; });

  if (kind === 'overview') {
    const per = {};
    items.forEach(function (x) {
      const o = per[x.r.person] || (per[x.r.person] = { open: 0, done: 0, overdue: 0, stale: 0 });
      if (x.done) { o.done++; return; }
      o.open++;
      if (x.r.overdue_days > 0) o.overdue++;
      if (x.r.days_since_update !== null && x.r.days_since_update >= 7) o.stale++;
    });
    const names = Object.keys(per).sort(function (a, b) { return per[b].overdue - per[a].overdue || per[b].open - per[a].open; });
    const sum = names.reduce(function (s, n) { s.open += per[n].open; s.done += per[n].done; s.overdue += per[n].overdue; s.stale += per[n].stale; return s; }, { open: 0, done: 0, overdue: 0, stale: 0 });
    const lines = ['**Overview — ' + waTodayKey_() + '**',
      '- Total **' + items.length + '** · Open **' + sum.open + '** · Done **' + sum.done + '** · Overdue **' + sum.overdue + '** · Stale (7+ din) **' + sum.stale + '**',
      '', '**Har employee:**'];
    names.forEach(function (n) {
      const o = per[n];
      lines.push('- **' + n + '**: open ' + o.open + ', overdue ' + o.overdue + ', stale ' + o.stale + ', done ' + o.done);
    });
    const worst = open.filter(function (x) { return x.r.overdue_days > 0; })
      .sort(function (a, b) { return b.r.overdue_days - a.r.overdue_days; }).slice(0, 5);
    if (worst.length) { lines.push('', '**Sabse zyada late:**'); worst.forEach(function (x) { lines.push(aiQLate_(x.r)); }); }
    return lines.join('\n');
  }

  if (kind === 'overdue') {
    const rows = open.filter(function (x) { return x.r.overdue_days > 0; })
      .sort(function (a, b) { return b.r.overdue_days - a.r.overdue_days; });
    if (!rows.length) return '✅ Koi overdue task nahi hai.';
    const lines = ['**Overdue tasks: ' + rows.length + '** (sabse late pehle)'];
    rows.slice(0, 40).forEach(function (x) { lines.push(aiQLate_(x.r)); });
    if (rows.length > 40) lines.push('… aur ' + (rows.length - 40) + ' tasks.');
    return lines.join('\n');
  }

  if (kind === 'stale') {
    const rows = open.filter(function (x) { return x.r.days_since_update !== null && x.r.days_since_update >= 7; })
      .sort(function (a, b) { return b.r.days_since_update - a.r.days_since_update; });
    if (!rows.length) return '✅ Koi open task 7+ din se pending-update nahi hai.';
    const lines = ['**Stale tasks (7+ din se update nahi): ' + rows.length + '**'];
    rows.slice(0, 40).forEach(function (x) {
      lines.push('- **' + x.r.person + '** — ' + x.r.task + ' · ' + x.r.days_since_update + ' din se update nahi · ' + x.r.status + ' · due ' + (x.r.deadline || '-'));
    });
    return lines.join('\n');
  }

  if (kind === 'whodoes') {
    if (!open.length) return '✅ Koi open task nahi hai.';
    const by = {};
    open.forEach(function (x) { (by[x.r.person] || (by[x.r.person] = [])).push(x.r); });
    const names = Object.keys(by).sort(function (a, b) { return by[b].length - by[a].length; });
    const lines = ['**Kaun kya kar raha hai** — ' + open.length + ' open tasks'];
    names.forEach(function (n) {
      const list = by[n].sort(function (a, b) { return b.overdue_days - a.overdue_days; });
      lines.push('', '**' + n + '** (' + list.length + ')');
      list.slice(0, 8).forEach(function (r) { lines.push(aiQLine_(r)); });
      if (list.length > 8) lines.push('… aur ' + (list.length - 8) + ' tasks. Poora dekhne ke liye naam search karo.');
    });
    return lines.join('\n');
  }

  if (kind === 'find') return aiQFind_(items, q);

  throw new Error('Unknown quick report: ' + kind);
}

// Search: every word you type must appear in the employee name, task name, details, remark or notes.
// Examples: "ledger" · "mukesh" · "mukesh ledger" · "kandla container" · "pending high"
function aiQFind_(items, q) {
  const words = String(q || '').toLowerCase().split(/[\s,]+/).filter(Boolean);
  if (!words.length) return 'Search me kuch likho, jaise: ledger, Mukesh, Kandla container.';
  const hits = items.filter(function (x) {
    return words.every(function (w) { return x.hay.indexOf(w) !== -1; });
  });
  if (!hits.length) return 'Is search me koi task nahi mila: **' + q + '**. Chhota ya alag word try karo (naam, party, product).';

  const by = {};
  hits.forEach(function (x) { (by[x.r.person] || (by[x.r.person] = [])).push(x); });
  const names = Object.keys(by).sort(function (a, b) { return by[b].length - by[a].length; });
  const lines = ['**"' + q + '"** — ' + hits.length + ' tasks, ' + names.length + ' log'];
  names.forEach(function (n) {
    const list = by[n].sort(function (a, b) {                       // open first, most late first
      return (a.done - b.done) || (b.r.overdue_days - a.r.overdue_days);
    });
    const openN = list.filter(function (x) { return !x.done; }).length;
    lines.push('', '**' + n + '** — ' + openN + ' open, ' + (list.length - openN) + ' done');
    list.slice(0, 8).forEach(function (x) { lines.push(x.done ? '- ✅ ' + x.r.task + ' · done' : aiQLine_(x.r)); });
    if (list.length > 8) lines.push('… aur ' + (list.length - 8) + ' tasks.');
  });
  return lines.join('\n');
}

function aiQMis_(token, session) {
  const r = aiToolReport_(token, session, { kind: 'mis', mis_type: 'delegation' });
  const people = (r.people || []).slice().sort(function (a, b) { return (a.score || 0) - (b.score || 0); });
  if (!people.length) return 'Is period me MIS data nahi mila.';
  const lines = ['**MIS score (' + r.type + ') ' + r.from + ' → ' + r.to + '** · average **' + r.average_score + '**',
    '0 = perfect, jitna neeche utna kharab.'];
  people.forEach(function (p) {
    lines.push('- **' + p.name + '**: score **' + p.score + '** · planned ' + p.planned + ', time par ' + p.done_on_time + ', late ' + p.done_late + ', baaki ' + p.not_done);
  });
  return lines.join('\n');
}

// test from the editor (use any valid token)
function testAiQuick() {
  Logger.log(JSON.stringify(aiQuickBuild_ ? 'aiQuickBuild_ exists' : 'missing'));
}