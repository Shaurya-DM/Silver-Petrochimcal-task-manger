// ============================================================
// AI RAG — semantic search over task history (task name, details, remarks, notes)
// ------------------------------------------------------------
// NEW FILE in your Apps Script project (File > + > Script > name it "AIRag").
// Works together with AIAssistant. Reuses: aiKey_, aiAllTasks_, aiTaskRow_, aiResolvePerson_,
//         aiDmy_, verifySession_, assertIsAdmin_, getScriptProp_
//
// HOW IT WORKS
//  1. Each task's text (name + description + remark + notes) is turned into an embedding
//     (a list of 256 numbers that captures meaning) by Gemini and saved in a hidden sheet "AI_Index".
//  2. When the user asks something like "what happened with the Kandla container?", the question is
//     embedded the same way, the closest tasks are found, and Gemini answers from those tasks only.
//  3. The index fills itself: every search first embeds up to 100 new/changed tasks.
//     For the very first build, an Admin can run aiRagReindex(token) once (see bottom).
//
// SAFETY: results are joined with the LIVE task list the user is allowed to see (aiAllTasks_),
// so an employee can never get another person's task through search, even though the index is shared.
//
// OPTIONAL Script Property:
//   AI_INDEX_SHEET_ID = id of the spreadsheet that should hold the AI_Index tab
//   (if empty, the spreadsheet this script is attached to is used)
// ============================================================

const RAG_SHEET_NAME = 'AI_Index';
const RAG_EMBED_MODEL = 'gemini-embedding-001';
const RAG_DIMS = 256;

function ragSheet_() {
  const id = getScriptProp_('AI_INDEX_SHEET_ID', '');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RAG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(RAG_SHEET_NAME);
    sh.appendRow(['key', 'personKey', 'hash', 'vector']);
    try { sh.hideSheet(); } catch (e) { /* ignore */ }
  }
  return sh;
}

// ---------- embeddings ----------

function ragNorm_(v) {
  let s = 0;
  for (let i = 0; i < v.length; i++) s += v[i] * v[i];
  s = Math.sqrt(s) || 1;
  return v.map(function (x) { return Math.round((x / s) * 10000) / 10000; });
}

// taskType: 'RETRIEVAL_DOCUMENT' for tasks, 'RETRIEVAL_QUERY' for the user's question
function ragEmbed_(texts, taskType) {
  const key = aiKey_();
  if (!key) throw new Error('AI is not set up yet: add GEMINI_API_KEY in Script Properties.');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + RAG_EMBED_MODEL + ':batchEmbedContents';
  const out = [];
  for (let i = 0; i < texts.length; i += 100) {
    const chunk = texts.slice(i, i + 100);
    const body = {
      requests: chunk.map(function (t) {
        return { model: 'models/' + RAG_EMBED_MODEL, content: { parts: [{ text: t }] }, taskType: taskType, outputDimensionality: RAG_DIMS };
      })
    };
    let res = null;
    for (let attempt = 1; attempt <= 3; attempt++) {
      res = UrlFetchApp.fetch(url, {
        method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key },
        payload: JSON.stringify(body), muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code === 200) break;
      if ([429, 500, 503, 504].indexOf(code) !== -1 && attempt < 3) { Utilities.sleep(1000 * attempt); continue; }
      throw new Error('Embedding error ' + code + ': ' + res.getContentText().slice(0, 200));
    }
    JSON.parse(res.getContentText()).embeddings.forEach(function (e) { out.push(ragNorm_(e.values)); });
  }
  return out;
}

// ---------- what gets indexed ----------

function ragText_(t) {
  return [
    'Task: ' + (t.TaskName || ''),
    t.Description ? 'Details: ' + t.Description : '',     // remove this line if your sheet has no Description column
    t.Remark ? 'Remark: ' + t.Remark : '',
    t.Notes ? 'Notes: ' + t.Notes : ''
  ].filter(Boolean).join('\n').slice(0, 2000);
}

function ragHash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8));
}

// Embeds new or changed tasks (max maxNew per call) and saves them in the AI_Index sheet.
function aiRagSync_(token, session, maxNew) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { added: 0, remaining: 0 };
  try {
    const sh = ragSheet_();
    const last = sh.getLastRow();
    const existing = {};
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (r, i) { existing[r[0]] = { row: i + 2, hash: r[2] }; });
    }
    const todo = [];
    aiAllTasks_(token, session).forEach(function (t) {
      const text = ragText_(t);
      const key = t.OwnerKey + '|' + t.rowNumber;
      const h = ragHash_(text);
      if (existing[key] && existing[key].hash === h) return;
      todo.push({ key: key, personKey: t.OwnerKey, hash: h, text: text, row: existing[key] ? existing[key].row : 0 });
    });
    const batch = todo.slice(0, Math.max(1, maxNew || 100));
    if (!batch.length) return { added: 0, remaining: 0 };

    const vecs = ragEmbed_(batch.map(function (b) { return b.text; }), 'RETRIEVAL_DOCUMENT');
    const append = [];
    batch.forEach(function (b, i) {
      const row = [b.key, b.personKey, b.hash, JSON.stringify(vecs[i])];
      if (b.row) sh.getRange(b.row, 1, 1, 4).setValues([row]);
      else append.push(row);
    });
    if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, 4).setValues(append);
    return { added: batch.length, remaining: todo.length - batch.length };
  } finally {
    lock.releaseLock();
  }
}

// ---------- the tool Gemini can call ----------

function aiToolSearchHistory_(token, session, a) {
  const q = String(a.query || '').trim();
  if (!q) throw new Error('query is required.');
  let sync = { added: 0, remaining: 0 };
  try { sync = aiRagSync_(token, session, 100); } catch (e) { Logger.log('aiRagSync_: ' + e); }

  const pk = a.person ? aiResolvePerson_(a.person, session).key : '';
  const live = {};                                          // only tasks this user is allowed to see
  aiAllTasks_(token, session).forEach(function (t) {
    if (!pk || t.OwnerKey === pk) live[t.OwnerKey + '|' + t.rowNumber] = t;
  });

  const qv = ragEmbed_([q], 'RETRIEVAL_QUERY')[0];
  const sh = ragSheet_();
  const last = sh.getLastRow();
  const scored = [];
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) {
      const t = live[r[0]];
      if (!t) return;
      let v;
      try { v = JSON.parse(r[3]); } catch (e) { return; }
      let s = 0;
      for (let i = 0; i < v.length && i < qv.length; i++) s += v[i] * qv[i];
      scored.push({ s: s, t: t });
    });
  }
  scored.sort(function (x, y) { return y.s - x.s; });
  const k = Math.max(1, Math.min(15, Number(a.top_k) || 8));
  return {
    note: 'Matches are ranked by meaning. A low match_score means a weak match; say so if nothing fits well.' +
      (sync.remaining > 0 ? ' Index is still being built (' + sync.remaining + ' tasks not searchable yet); results may be incomplete.' : ''),
    results: scored.slice(0, k).map(function (x) {
      const row = aiTaskRow_(x.t);
      row.match_score = Math.round(x.s * 100) / 100;
      row.all_notes = String(x.t.Notes || '').slice(-600);
      return row;
    })
  };
}

// ---------- first-time build (run once, as Admin) ----------
// Call from the app or from a small test function that passes an Admin token.
function aiRagReindex(token) {
  const session = verifySession_(token);
  assertIsAdmin_(session);
  const end = Date.now() + 240000;                          // stay inside the 6-minute limit
  let total = 0, r = { added: 0, remaining: 0 };
  do {
    r = aiRagSync_(token, session, 100);
    total += r.added;
  } while (r.remaining > 0 && r.added > 0 && Date.now() < end);
  return { success: true, added: total, remaining: r.remaining };
}

// ============================================================
// 3 SMALL CHANGES TO AIAssistant (copy these in)
// ============================================================
//
// (1) In aiTools_(): add this object inside the `read` array, after get_report:
//
//    ,{
//      name: 'search_task_history',
//      description: 'Search task names, details, remarks and notes BY MEANING, including completed tasks. Use it for "what happened with...", "any earlier task about...", "who handled... before", "find tasks similar to...".',
//      parameters: { type: 'object', properties: {
//        query: { type: 'string', description: 'What to look for, in plain words.' },
//        person: { type: 'string', description: 'Optional: only this employee.' },
//        top_k: { type: 'integer', description: 'How many results, default 8, max 15.' }
//      }, required: ['query'] }
//    }
//
// (2) In aiChat(), inside the tool dispatch, right after the get_report line, add:
//
//        else if (name === 'search_task_history') out = aiToolSearchHistory_(token, session, a);
//
// (3) In aiSystemText_(), add this line to the array (before `snapshot || ''`):
//
//    'For questions about past work, notes, remarks or similar earlier tasks, call search_task_history, then answer only from the results and mention the person and deadline. Use get_tasks / get_report for counts, lists and filters.',