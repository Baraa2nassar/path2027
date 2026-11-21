/**
 * Gym Log + My 2027 Path: Google Sheet backend.
 * Paste this whole file into Extensions → Apps Script of your Google Sheet.
 * Setup steps are in google-sheet/README.md.
 *
 * @OnlyCurrentDoc  Limits the script to this one spreadsheet, so Google never
 *                  grants it access to any other files in the account.
 */

const SHEET_NAME = 'Sets';
const HEADERS = ['id', 'when', 'day', 'exercise', 'weight_lb', 'weight_kg', 'reps',
  'bar_lb', 'plates_per_side', 'pr', 'est_date', 'exercise_id', 'type'];

// Run this once from the Apps Script editor. It creates the "Sets" tab and
// prints your secret key in the Execution log.
function setup() {
  sheet_();
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('KEY', key);
  }
  Logger.log('Your secret key (paste it into the app): ' + key);
}

// Health check, so opening the web app URL in a browser shows something.
function doGet() {
  return json_({ ok: true, app: 'gym-log' });
}

// One request does everything: { key, add: [rows], del: [ids], config, list: true }
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad request' });
  }
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('KEY');
  if (!key || req.key !== key) return json_({ ok: false, error: 'wrong key' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (req.app === 'path') return json_(pathPost_(req)); // My 2027 Path page
    const sh = sheet_();
    if (req.add && req.add.length) addRows_(sh, req.add);
    if (req.del && req.del.length) deleteRows_(sh, req.del);
    if (req.config) props.setProperty('CONFIG', JSON.stringify(req.config));
    const res = { ok: true };
    if (req.list) {
      res.entries = readRows_(sh);
      res.config = JSON.parse(props.getProperty('CONFIG') || 'null');
    }
    return json_(res);
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Map header name -> column index, so columns can be reordered in the sheet.
function columns_(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = {};
  head.forEach((h, i) => { col[String(h).trim()] = i; });
  return col;
}

function addRows_(sh, entries) {
  const col = columns_(sh);
  const width = sh.getLastColumn();
  const have = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, col.id + 1, sh.getLastRow() - 1, 1).getValues()
      .forEach(r => { have[String(r[0])] = true; });
  }
  const rows = [];
  entries.forEach(e => {
    if (!e || !e.id || have[e.id]) return;
    have[e.id] = true;
    const row = new Array(width).fill('');
    const put = (name, v) => { if (name in col) row[col[name]] = v; };
    put('id', e.id);
    put('when', new Date(e.ts));
    put('day', e.day || '');
    put('exercise', e.name || e.ex);
    put('weight_lb', e.w);
    put('weight_kg', Math.round(e.w * 0.45359237 * 10) / 10);
    put('reps', e.reps);
    put('bar_lb', e.d ? e.d.bar : '');
    put('plates_per_side', e.d ? e.d.plates.join('+') : '');
    put('pr', e.pr ? 'yes' : '');
    put('est_date', e.est ? 'yes' : '');
    put('exercise_id', e.ex);
    put('type', e.type || '');
    rows.push(row);
  });
  if (!rows.length) return;
  rows.sort((a, b) => a[col.when] - b[col.when]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, width).setValues(rows);
}

function deleteRows_(sh, ids) {
  if (sh.getLastRow() < 2) return;
  const col = columns_(sh);
  const kill = {};
  ids.forEach(id => { kill[String(id)] = true; });
  const vals = sh.getRange(2, col.id + 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (kill[String(vals[i][0])]) sh.deleteRow(i + 2);
  }
}

function readRows_(sh) {
  if (sh.getLastRow() < 2) return [];
  const col = columns_(sh);
  const range = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn());
  const vals = range.getValues();
  const out = [];
  vals.forEach((r, i) => {
    const ex = String(r[col.exercise_id] || '').trim();
    if (!ex) return;
    let id = String(r[col.id] || '').trim();
    if (!id) {
      // A row typed in by hand: give it an id so the app can delete it later.
      id = 'sheet-' + Utilities.getUuid().slice(0, 8);
      sh.getRange(i + 2, col.id + 1).setValue(id);
    }
    const when = r[col.when];
    const ts = when instanceof Date ? when.getTime() : Date.parse(when);
    if (!ts) return;
    const e = { id: id, ts: ts, ex: ex, w: Number(r[col.weight_lb]) || 0, reps: Number(r[col.reps]) || 0 };
    const bar = r[col.bar_lb];
    if (bar !== '' && bar !== undefined) {
      e.d = {
        bar: Number(bar) || 0,
        plates: String(r[col.plates_per_side]).split('+').map(Number).filter(n => n > 0),
      };
    }
    if (String(r[col.pr]).toLowerCase() === 'yes') e.pr = true;
    if (String(r[col.est_date]).toLowerCase() === 'yes') e.est = 1;
    out.push(e);
  });
  return out;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}


// ---------------------------------------------------------------------------
// MY 2027 PATH
// The 2027 Path page sends { key, app: 'path', put: [{ id, row }], list: true }.
// Each week is one row in the "2027 Path" tab, found by its Monday date in column A.
// The reply also carries your gym days from the "Sets" tab, so Push/Pull/Legs fill in.
// ---------------------------------------------------------------------------
const PATH_SHEET = '2027 Path';
const PATH_HEADERS = ['week', 'task', 'task_done', 'surah', 'from_ayah', 'to_ayah', 'lessons',
  'push', 'pull', 'legs', 'week_done', 'score'];

function pathPost_(req) {
  const sh = pathSheet_();
  if (req.put && req.put.length) pathPut_(sh, req.put);
  const res = { ok: true };
  if (req.list) {
    res.weeks = pathRead_(sh);
    res.gym = gymDays_();
    res.plan = planRead_();
    res.months = monthsRead_();
    res.dreams = dreamsRead_();
  }
  return res;
}

// Rows the app renders as the weekly plan: [["2026-11-23","Check the course..."], ...]
function planRead_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plan');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .map(r => [dayKey_(r[0]), String(r[1] || '').trim()])
    .filter(r => /^\d{4}-\d{2}-\d{2}$/.test(r[0]) && r[1]);
}

// Months tab, columns: month (YYYY-MM), quran, course, film, body, family, money
// Returned as { "2026-11": { q, c, f, b, fa, m }, ... }
function monthsRead_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Months');
  if (!sh || sh.getLastRow() < 2) return {};
  const keys = ['q', 'c', 'f', 'b', 'fa', 'm'];
  const out = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().forEach(r => {
    const m = String(r[0] || '').trim().slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(m)) return;
    const o = {};
    keys.forEach((k, i) => { o[k] = String(r[i + 1] || '').trim(); });
    out[m] = o;
  });
  return out;
}

// Dreams tab, columns: group, icon, title, text, main (yes/empty)
// Returned as [[group, [[icon,title,text,cls], ...]], ...]
function dreamsRead_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dreams');
  if (!sh || sh.getLastRow() < 2) return [];
  const groups = {};
  const order = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(r => {
    const g = String(r[0] || '').trim();
    const ic = String(r[1] || '').trim();
    const h = String(r[2] || '').trim();
    const p = String(r[3] || '').trim();
    if (!g || !ic || !h) return;
    if (!groups[g]) { groups[g] = []; order.push(g); }
    const cls = String(r[4] || '').toLowerCase() === 'yes' ? 'main' : '';
    groups[g].push([ic, h, p, cls]);
  });
  return order.map(g => [g, groups[g]]);
}

function pathSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PATH_SHEET);
  if (!sh) sh = ss.insertSheet(PATH_SHEET);
  if (sh.getLastRow() === 0 || String(sh.getRange(1, 1).getValue()).trim() === '') {
    sh.getRange(1, 1, 1, PATH_HEADERS.length).setValues([PATH_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// A date cell or a "2026-11-23" text cell both become "2026-11-23".
function dayKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim().slice(0, 10);
}

function pathPut_(sh, puts) {
  const width = PATH_HEADERS.length;
  const rowOf = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
      .forEach((r, i) => { rowOf[dayKey_(r[0])] = i + 2; });
  }
  puts.forEach(p => {
    if (!p || !p.id || !p.row) return;
    const row = p.row.slice(0, width);
    while (row.length < width) row.push('');
    row[0] = p.id;
    let r = rowOf[p.id];
    if (!r) { r = sh.getLastRow() + 1; rowOf[p.id] = r; }
    sh.getRange(r, 1).setNumberFormat('@'); // keep the week as plain text
    sh.getRange(r, 1, 1, width).setValues([row]);
  });
}

function pathRead_(sh) {
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, PATH_HEADERS.length).getValues()
    .map(r => { r[0] = dayKey_(r[0]); return r; })
    .filter(r => /^\d{4}-\d{2}-\d{2}$/.test(r[0]));
}

// [["2026-11-24", "Push"], ...] from the Gym Log "Sets" tab.
function gymDays_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return [];
  const col = columns_(sh);
  if (!('when' in col) || !('day' in col)) return [];
  const seen = {};
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(r => {
    const d = dayKey_(r[col.when]);
    const day = String(r[col.day] || '').trim();
    if (!d || !day || seen[d + day]) return;
    seen[d + day] = true;
    out.push([d, day]);
  });
  return out;
}
