// Code.gs v1.1.4 — EcoOils Project HSSE Recognition backend (Google Apps Script web app)
// Deploy: Extensions > Apps Script (on iggifernandes@gmail.com) > paste this file > run setup() once >
//   Project Settings > Script properties: add GEMINI_API_KEY (free Google AI Studio key) >
//   Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.

var SHEET_NAME = 'Nominations';
var COLS = ['id','createdAt','week','status','certNo','name','nick','role','ic','location','eventDate',
  'observation','action','outcome','category','nomName','nomRole','nomIC','nomPhone',
  'beforePhotos','afterPhotos','aiScore','aiVerdict','aiJson','certLang','certDate','certPara','hseNote','updatedAt','source','level'];
var EDITABLE = ['status','certNo','name','nick','role','ic','location','eventDate','observation','action','outcome',
  'category','certLang','certDate','certPara','hseNote','level','aiJson'];
var IC_LIST = ['EcoOils','GSB','WEN','KKS','PBS'];
var VERSION = 'v1.1.4';

/* ---------- Rewards: voucher stock and reward distribution (admin only) ---------- */
var RW_NAME = 'Rewards';
var RW_COLS = ['id','createdAt','date','ic','name','category','items','amount','remarks','nominationId','serials','photos','status','issuedBy','updatedAt','alloc'];
var ST_NAME = 'Stock';
var ST_COLS = ['id','createdAt','date','type','denom','qty','custodian','ref','expiry','serials','photos','remarks','updatedAt'];
var DATE_COLS = ['date','expiry'];

function tab_(name, cols) {
  var ss = SpreadsheetApp.openById(prop_('SHEET_ID')), sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 || sh.getLastColumn() < cols.length) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold'); sh.setFrozenRows(1);
    cols.forEach(function (k, i) { if (DATE_COLS.indexOf(k) >= 0 || k === 'serials' || k === 'items') sh.getRange(2, i + 1, 2000, 1).setNumberFormat('@'); });
  }
  return sh;
}
function readTab_(name, cols) {
  var sh = tab_(name, cols), n = sh.getLastRow();
  if (n < 2) return [];
  var tz = SpreadsheetApp.openById(prop_('SHEET_ID')).getSpreadsheetTimeZone();
  return sh.getRange(2, 1, n - 1, cols.length).getValues().map(function (v) {
    var o = {}; cols.forEach(function (k, i) {
      var x = v[i];
      if (x instanceof Date) x = DATE_COLS.indexOf(k) >= 0 ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : x.toISOString();
      o[k] = x;
    });
    return o;
  }).filter(function (o) { return o.id; });
}
function findRow_(sh, id) {
  var n = sh.getLastRow(); if (n < 2) return -1;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function newId_(prefix) { return prefix + '-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase(); }
function isoDate_(v) { var t = s_(v, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) throw new Error('Bad date'); return t; }
function itemsClean_(items) {
  var o = {}, total = 0;
  Object.keys(items || {}).forEach(function (k) {
    var n = Math.floor(Number(items[k]));
    if (/^\d{4,7}$/.test(k) && n > 0 && n <= 1000) { o[k] = n; total += Number(k) * n; }
  });
  return { items: o, amount: total };
}
function rwList_() { return { ok: true, rewards: readTab_(RW_NAME, RW_COLS), stock: readTab_(ST_NAME, ST_COLS) }; }
function folderPhotos_(uris, prefix) { return uris && uris.length ? savePhotos_(uris, DriveApp.getFolderById(prop_('FOLDER_ID')), prefix) : []; }
function trash_(ids) { ids.forEach(function (f) { f = String(f).trim(); if (f) { try { DriveApp.getFileById(f).setTrashed(true); } catch (e) {} } }); }

function allocClean_(a, total) {
  if (!a || !a.length) return '';
  if (a.length > 10) throw new Error('Too many companies');
  var out = [], sum = 0;
  a.forEach(function (x) { var ic = s_(x.ic, 40), amt = Math.max(0, Math.floor(Number(x.amount) || 0)); if (!ic) throw new Error('Company missing'); out.push({ ic: ic, amount: amt }); sum += amt; });
  if (sum !== total) throw new Error('Company amounts must add up to the voucher total');
  return JSON.stringify(out);
}
function rwSave_(r) {
  var f = r.reward || {}, ic = itemsClean_(f.items);
  if (!s_(f.name) || !s_(f.category)) throw new Error('Missing name or category');
  if (ic.amount <= 0) throw new Error('Choose at least one voucher');
  var date = isoDate_(f.date);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = tab_(RW_NAME, RW_COLS), now = new Date().toISOString(), row = f.id ? findRow_(sh, f.id) : -1, rec = {};
    if (f.id && row < 0) throw new Error('Not found');
    var id = f.id || newId_('RW');
    if (row > 0) { var cur = sh.getRange(row, 1, 1, RW_COLS.length).getValues()[0]; RW_COLS.forEach(function (k, i) { rec[k] = cur[i]; }); }
    var keep = (f.keepPhotos || []).map(String), old = String(rec.photos || '').split(',').filter(Boolean);
    trash_(old.filter(function (x) { return keep.indexOf(x) < 0; }));
    var kept = old.filter(function (x) { return keep.indexOf(x) >= 0; });
    var added = folderPhotos_((f.newPhotos || []).slice(0, Math.max(0, 3 - kept.length)), id + '-reward');
    var o = {
      id: id, createdAt: rec.createdAt || now, date: date, ic: s_(f.ic, 120), name: s_(f.name, 160), category: s_(f.category, 120), alloc: allocClean_(f.alloc, ic.amount),
      items: JSON.stringify(ic.items), amount: ic.amount, remarks: s_(f.remarks, 600), nominationId: s_(f.nominationId, 40),
      serials: s_(f.serials, 3000), photos: kept.concat(added).join(','), status: 'Issued', issuedBy: rec.issuedBy || s_(r.email || '', 80), updatedAt: now
    };
    var vals = RW_COLS.map(function (k) { return o[k] == null ? '' : o[k]; });
    if (row > 0) sh.getRange(row, 1, 1, RW_COLS.length).setValues([vals]); else sh.appendRow(vals);
    return { ok: true, reward: o };
  } finally { lock.releaseLock(); }
}

function rwSaveMany_(r) {
  var rows = r.rows || [], sh = r.shared || {};
  if (!rows.length || rows.length > 30) throw new Error('Send 1 to 30 people');
  rows.forEach(function (x, i) {
    if (!s_(x.name) || !s_(x.ic) || !s_(x.category) || itemsClean_(x.items).amount <= 0) throw new Error('Person ' + (i + 1) + ': name, company, reason and vouchers are required');
  });
  var saved = [];
  rows.forEach(function (x, i) {
    var f = { date: sh.date, remarks: sh.remarks, serials: sh.serials, name: x.name, ic: x.ic, category: x.category, items: x.items };
    if (i === 0) { f.newPhotos = sh.newPhotos; f.keepPhotos = []; }
    saved.push(rwSave_({ reward: f, email: r.email }).reward);
  });
  return { ok: true, rewards: saved };
}
function stSave_(r) {
  var f = r.row || {}, type = s_(f.type, 6);
  if (['OPEN', 'IN', 'OUT', 'ADJ'].indexOf(type) < 0) throw new Error('Bad type');
  var denom = Math.floor(Number(f.denom)), qty = Math.floor(Number(f.qty));
  if (!(denom >= 1000 && denom <= 9999999)) throw new Error('Bad denomination');
  if (!qty || (type !== 'ADJ' && qty < 0) || Math.abs(qty) > 100000) throw new Error('Bad quantity');
  var date = isoDate_(f.date), expiry = s_(f.expiry, 10); if (expiry) isoDate_(expiry);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = tab_(ST_NAME, ST_COLS), now = new Date().toISOString(), row = f.id ? findRow_(sh, f.id) : -1, rec = {};
    if (f.id && row < 0) throw new Error('Not found');
    var id = f.id || newId_('ST');
    if (row > 0) { var cur = sh.getRange(row, 1, 1, ST_COLS.length).getValues()[0]; ST_COLS.forEach(function (k, i) { rec[k] = cur[i]; }); }
    var keep = (f.keepPhotos || []).map(String), old = String(rec.photos || '').split(',').filter(Boolean);
    trash_(old.filter(function (x) { return keep.indexOf(x) < 0; }));
    var kept = old.filter(function (x) { return keep.indexOf(x) >= 0; });
    var added = folderPhotos_((f.newPhotos || []).slice(0, Math.max(0, 3 - kept.length)), id + '-stock');
    var o = { id: id, createdAt: rec.createdAt || now, date: date, type: type, denom: denom, qty: qty, custodian: s_(f.custodian, 80), ref: s_(f.ref, 160),
      expiry: expiry, serials: s_(f.serials, 3000), photos: kept.concat(added).join(','), remarks: s_(f.remarks, 600), updatedAt: now };
    var vals = ST_COLS.map(function (k) { return o[k] == null ? '' : o[k]; });
    if (row > 0) sh.getRange(row, 1, 1, ST_COLS.length).setValues([vals]); else sh.appendRow(vals);
    return { ok: true, row: o };
  } finally { lock.releaseLock(); }
}

function rowsDelete_(name, cols, ids, photoCols) {
  ids = (ids || []).map(String); if (!ids.length || ids.length > 300) throw new Error('Nothing to delete');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = tab_(name, cols), n = sh.getLastRow(), done = 0; if (n < 2) return { ok: true, deleted: 0 };
    var data = sh.getRange(2, 1, n - 1, cols.length).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (ids.indexOf(String(data[i][0])) < 0) continue;
      photoCols.forEach(function (pc) { trash_(String(data[i][cols.indexOf(pc)] || '').split(',')); });
      sh.deleteRow(i + 2); done++;
    }
    return { ok: true, deleted: done };
  } finally { lock.releaseLock(); }
}

// bulk import of past reward records (date, ic, name, amount, remarks). Skips rows already present.
function rwImport_(r) {
  var rows = r.rows || []; if (!rows.length || rows.length > 400) throw new Error('Send 1 to 400 rows');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = tab_(RW_NAME, RW_COLS), have = {}, added = 0, skipped = 0, now = new Date().toISOString(), out = [];
    readTab_(RW_NAME, RW_COLS).forEach(function (x) { have[[x.date, x.ic, x.name, x.amount, x.remarks].join('|')] = 1; });
    rows.forEach(function (f) {
      var date = s_(f.date, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) { skipped++; return; }
      var amount = Math.max(0, Math.floor(Number(f.amount) || 0)), items = {}, left = amount;
      [100000, 50000].forEach(function (d) { var n = Math.floor(left / d); if (n) { items[d] = n; left -= n * d; } });
      var o = { id: newId_('RW'), createdAt: now, date: date, ic: s_(f.ic, 40), name: s_(f.name, 160), category: s_(f.category || 'Imported', 120),
        items: JSON.stringify(items), amount: amount - left, remarks: s_(f.remarks, 600), nominationId: '', serials: '', photos: '', status: 'Issued', issuedBy: 'import', updatedAt: now };
      var key = [o.date, o.ic, o.name, o.amount, o.remarks].join('|');
      if (have[key]) { skipped++; return; }
      have[key] = 1; out.push(RW_COLS.map(function (k) { return o[k] == null ? '' : o[k]; })); added++;
    });
    if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, RW_COLS.length).setValues(out);
    return { ok: true, added: added, skipped: skipped };
  } finally { lock.releaseLock(); }
}

/* ---------- one-time setup ---------- */
function setup() {
  var p = PropertiesService.getScriptProperties();
  var ss = p.getProperty('SHEET_ID') ? SpreadsheetApp.openById(p.getProperty('SHEET_ID'))
    : SpreadsheetApp.create('EcoOils Project - HSSE Recognition');
  p.setProperty('SHEET_ID', ss.getId());
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  tab_(RW_NAME, RW_COLS); tab_(ST_NAME, ST_COLS);
  if (!p.getProperty('FOLDER_ID')) {
    p.setProperty('FOLDER_ID', DriveApp.createFolder('EcoOils Project - Recognition photos').getId());
  }
  if (!p.getProperty('ADMIN_EMAILS')) p.setProperty('ADMIN_EMAILS', 'Ignatius.Ambrosio@ecooils.com.my');
  if (!p.getProperty('PIN_SALT')) p.setProperty('PIN_SALT', Utilities.getUuid());
  Logger.log('Sheet: ' + ss.getUrl());
  Logger.log('Now add GEMINI_API_KEY under Project Settings > Script properties, then deploy as a Web app.');
}

/* Run this from the editor (select testAI > Run) to check the Gemini key. Read the result under View > Logs. */
function testAI() {
  Logger.log('GEMINI_API_KEY present: ' + !!prop_('GEMINI_API_KEY'));
  var r = aiReview_({ client: 'test', fields: { name: 'Test', role: 'Technician', ic: 'WEN', location: 'DTP',
    observation: 'Scrap cable and tools were left around the scaffold base after termination work, a trip hazard.',
    action: 'He stopped, collected all offcuts and tools and cleaned the area before leaving.',
    outcome: 'Walkway clear before the next crew arrived.' }, before: [], after: [] });
  Logger.log(r.ok ? 'AI OK (model ' + r.model + '). Highlight (EN): ' + r.result.improved.en.highlight : 'AI FAILED: ' + r.error);
}

/* ---------- web app entry points ---------- */
function doGet(e) {
  return out_({ ok: true, version: VERSION, aiReady: !!prop_('GEMINI_API_KEY'), models: models_('full'), quickModels: models_('quick') });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'Bad request' }); }
  try {
    switch (req.action) {
      case 'ping': return out_({ ok: true, version: VERSION, aiReady: !!prop_('GEMINI_API_KEY') });
      case 'ai': return out_(aiReview_(req));
      case 'submit': return out_(submit_(req));
      case 'requestReset': return out_(requestReset_(req));
      case 'setPin': return out_(setPin_(req));
      case 'login': return out_(login_(req));
      case 'logout': CacheService.getScriptCache().remove('sess_' + req.token); return out_({ ok: true });
      case 'adminCreate': requireAdmin_(req); return out_(adminCreate_(req));
      case 'verify': return out_(verify_(req));
      case 'list': requireAdmin_(req); return out_({ ok: true, rows: readAll_() });
      case 'photo': requireAdmin_(req); return out_(photo_(req));
      case 'update': requireAdmin_(req); return out_(update_(req));
      case 'delete': requireAdmin_(req); return out_(delete_(req));
      case 'rwList': requireAdmin_(req); return out_(rwList_());
      case 'rwSave': req.email = requireAdmin_(req); return out_(rwSave_(req));
      case 'rwSaveMany': req.email = requireAdmin_(req); return out_(rwSaveMany_(req));
      case 'rwDelete': requireAdmin_(req); return out_(rowsDelete_(RW_NAME, RW_COLS, req.ids, ['photos']));
      case 'rwImport': requireAdmin_(req); return out_(rwImport_(req));
      case 'stSave': requireAdmin_(req); return out_(stSave_(req));
      case 'stDelete': requireAdmin_(req); return out_(rowsDelete_(ST_NAME, ST_COLS, req.ids, ['photos']));
      default: return out_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function sheet_() {
  var sh = SpreadsheetApp.openById(prop_('SHEET_ID')).getSheetByName(SHEET_NAME);
  // upgrade older sheets: add any new column headers (e.g. 'level') without touching existing data
  if (sh && sh.getLastColumn() < COLS.length) sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  return sh;
}
function s_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max || 2000); }

/* ---------- submit ---------- */
function weekKey_(d) {
  var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  var day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  var y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  var w = Math.ceil(((t - y0) / 86400000 + 1) / 7);
  return t.getUTCFullYear() + '-W' + (w < 10 ? '0' : '') + w;
}

function savePhotos_(arr, folder, prefix) {
  var ids = [];
  (arr || []).slice(0, 3).forEach(function (uri, i) {
    var m = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+\/=]+)$/.exec(String(uri || ''));
    if (!m) throw new Error('Invalid photo');
    var bytes = Utilities.base64Decode(m[2]);
    if (bytes.length > 3 * 1024 * 1024) throw new Error('Photo too large');
    var ext = m[1].split('/')[1];
    var blob = Utilities.newBlob(bytes, m[1], prefix + '-' + (i + 1) + '.' + ext);
    ids.push(folder.createFile(blob).getId());
  });
  return ids;
}

function submit_(r) {
  if (r.website) return { ok: true, id: 'OK' }; // honeypot
  var f = r.fields || {};
  var need = ['name','role','ic','location','eventDate','observation','action','outcome','nomName','nomRole'];
  for (var i = 0; i < need.length; i++) if (!s_(f[need[i]])) throw new Error('Missing: ' + need[i]);
  if (IC_LIST.indexOf(s_(f.ic)) < 0 && s_(f.ic).length < 2) throw new Error('Invalid IC');
  if (s_(f.observation).length < 40 || s_(f.action).length < 40 || s_(f.outcome).length < 25) throw new Error('Too short');
  if (!r.before || !r.before.length || !r.after || !r.after.length) throw new Error('Before and After photos are required');

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var now = new Date();
    var id = 'NOM-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
    var folder = DriveApp.getFolderById(prop_('FOLDER_ID'));
    var b = savePhotos_(r.before, folder, id + '-before');
    var a = savePhotos_(r.after, folder, id + '-after');
    var ai = r.ai || {};
    var row = {
      id: id, createdAt: now.toISOString(), week: weekKey_(now), status: 'Pending', certNo: '',
      name: s_(f.name, 120), nick: s_(f.nick, 60), role: s_(f.role, 120), ic: s_(f.ic, 40), location: s_(f.location, 120),
      eventDate: s_(f.eventDate, 20), observation: s_(f.observation), action: s_(f.action), outcome: s_(f.outcome),
      category: s_(f.category, 60), nomName: s_(f.nomName, 120), nomRole: s_(f.nomRole, 120), nomIC: s_(f.nomIC, 40),
      nomPhone: s_(f.nomPhone, 40), beforePhotos: b.join(','), afterPhotos: a.join(','),
      aiScore: typeof ai.score === 'number' ? ai.score : '', aiVerdict: s_(ai.verdict, 20),
      aiJson: ai.improved ? JSON.stringify(ai.improved).slice(0, 45000) : '',
      level: (['crew','sup','hsse'].indexOf(s_(f.level)) >= 0 ? s_(f.level) : 'crew'), certLang: 'id', certDate: '', certPara: '', hseNote: '', updatedAt: now.toISOString(), source: s_(r.source || 'web', 20)
    };
    sheet_().appendRow(COLS.map(function (c) { return row[c] == null ? '' : row[c]; }));
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

/* ---------- Gemini wording helper (OPTIONAL: the nomination quality check runs in the app, no AI needed) ---------- */
var AI_SYSTEM = [
  'You are a senior HSSE communications writer for the EcoOils Project site (Sei Mangkei, North Sumatra, Indonesia). You turn a nominator\'s rough notes into bulletin-quality writing for the EcoOils HSSE Bulletin and recognition certificates.',
  'You REWRITE, you do not just correct. Improve structure, flow and impact: strong active verbs, concrete safety terms (hazard, unsafe condition, trip hazard, line of fire, barricade, isolate, stop work, housekeeping and similar) where they fit what was written, and a respectful, positive tone that honours the person.',
  'You may explain why the situation mattered in general terms (for example that a trip hazard can cause a fall injury), and you may reorder and expand the nominator\'s points into complete, well-built sentences.',
  'You must NOT invent specific facts: no names, numbers, times, equipment or tag numbers, quantities, injuries or events that the nominator did not state. If a detail a reader would want is missing, write the sentence without it and add a short question for the nominator to "issues".',
  'You never look at photos, never give a score and never decide whether the nomination is good. Set generic=true only for generic praise ("strong safety commitment").',
  'For each language give: observation (what was found and why it mattered), action (what the person did, step by step, active voice), outcome (the result and the benefit to site safety), highlight (one sentence, max 25 words, naming the intervention), story (3 to 4 sentences, third person, bulletin-ready, starting with the person\'s first name and recognising the intervention).',
  'Provide English (en) and Bahasa Indonesia (id). The Bahasa Indonesia must read naturally to Indonesian crew, not like a literal translation. Keep names, tag numbers and locations exactly as written.',
  'Reply ONLY with JSON matching the schema.'
].join('\n');

function schema_() {
  var tri = { type: 'OBJECT', properties: { observation: { type: 'STRING' }, action: { type: 'STRING' }, outcome: { type: 'STRING' }, highlight: { type: 'STRING' }, story: { type: 'STRING' } },
    required: ['observation', 'action', 'outcome', 'highlight', 'story'] };
  return { type: 'OBJECT', properties: {
    generic: { type: 'BOOLEAN' },
    issues: { type: 'ARRAY', items: { type: 'STRING' } },
    improved: { type: 'OBJECT', properties: { en: tri, id: tri }, required: ['en', 'id'] }
  }, required: ['generic', 'issues', 'improved'] };
}

/* AI models, in order of preference. To change them without touching code: Project Settings > Script properties >
   add GEMINI_MODELS = model1,model2,model3 (comma separated). Models that were busy in the last 2 minutes are tried last. */
var DEFAULT_MODELS = ['gemini-2.5-flash', 'gemini-2.5-flash-lite', 'gemini-flash-latest', 'gemini-flash-lite-latest'];
var QUICK_MODELS = ['gemini-2.5-flash-lite', 'gemini-2.5-flash', 'gemini-flash-lite-latest', 'gemini-flash-latest']; // lightest first = fastest
function models_(mode) {
  var quick = mode === 'quick', cache = CacheService.getScriptCache();
  var p = quick ? prop_('GEMINI_QUICK_MODELS') : (prop_('GEMINI_MODELS') || prop_('GEMINI_MODEL'));
  var list = p ? String(p).split(',').map(function (x) { return x.trim(); }).filter(Boolean) : (quick ? QUICK_MODELS : DEFAULT_MODELS), ok = [], busy = [];
  list.forEach(function (m) { (cache.get('busy_' + m) ? busy : ok).push(m); });
  return ok.concat(busy);
}

/* QUICK REPHRASE: only the three text boxes, same language as written, short output = a few seconds */
function trSchema_() {
  return { type: 'OBJECT', properties: { observation: { type: 'STRING' }, action: { type: 'STRING' }, outcome: { type: 'STRING' } }, required: ['observation', 'action', 'outcome'] };
}
var AI_TRANS = [
  'You are a professional translator for HSSE documents at a construction site in Indonesia.',
  'Translate the three notes (observation, action, outcome) into {LANG}. Keep the meaning exactly. Use clear, correct, professional wording.',
  'Do not add, remove or invent any facts, names, numbers or tags. Keep acronyms (PPE, PTW, JHA, LOTO, DTP, etc.) and proper names unchanged. If a note is already in {LANG}, only fix grammar.',
  'Reply ONLY with JSON matching the schema.'
].join('\n');
var AI_QUICK = [
  'You are an HSSE copy editor for a construction site in Indonesia.',
  'Rephrase the nominator\'s three notes into clear, correct, professional sentences in active voice. Write in the SAME language the nominator used (English, Bahasa Indonesia or Malay). Keep each answer about as long as the original.',
  'Do not invent facts: no names, numbers, tags, times or events that were not written. If an important detail is missing, add a short question to "issues" (at most 2).',
  'Reply ONLY with JSON matching the schema.'
].join('\n');
function quickSchema_() {
  return { type: 'OBJECT', properties: { observation: { type: 'STRING' }, action: { type: 'STRING' }, outcome: { type: 'STRING' }, issues: { type: 'ARRAY', items: { type: 'STRING' } } },
    required: ['observation', 'action', 'outcome', 'issues'] };
}

function aiReview_(r) {
  var key = prop_('GEMINI_API_KEY');
  if (!key) return { ok: false, error: 'AI not configured: add script property GEMINI_API_KEY (Project Settings > Script properties)' };
  var f = r.fields || {}, tr = r.mode === 'translate', quick = r.mode === 'quick' || tr;
  var target = { en: 'English', id: 'Bahasa Indonesia', ms: 'Bahasa Melayu (Malay)' }[s_(r.target, 4)] || 'English';
  var cache = CacheService.getScriptCache(), rk = 'airate_' + (s_(r.client, 40) || 'x');
  var n = Number(cache.get(rk) || 0);
  if (n >= 20) return { ok: false, error: 'Too many AI requests. Try again in a few minutes.' };
  cache.put(rk, String(n + 1), 600);

  var text = 'Nominee: ' + s_(f.name, 120) + ' (' + s_(f.role, 120) + ', ' + s_(f.ic, 40) + ')\nLocation: ' + s_(f.location, 120) +
    '\nObservation (hazard found): ' + s_(f.observation) + '\nAction taken: ' + s_(f.action) + '\nOutcome: ' + s_(f.outcome);
  function body_(noThink) {
    var g = { responseMimeType: 'application/json', responseSchema: tr ? trSchema_() : (quick ? quickSchema_() : schema_()), temperature: quick ? 0.3 : 0.5, maxOutputTokens: tr ? 2000 : (quick ? 900 : 3500) };
    if (noThink) g.thinkingConfig = { thinkingBudget: 0 }; // no hidden "thinking" time: much faster
    return JSON.stringify({ systemInstruction: { parts: [{ text: tr ? AI_TRANS.replace('{LANG}', target) : (quick ? AI_QUICK : AI_SYSTEM) }] }, contents: [{ role: 'user', parts: [{ text: text }] }], generationConfig: g });
  }
  function call_(model, payload) {
    return UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': key }, payload: payload });
  }
  // go down the model list; a model that says "high demand" (503), "rate limit" (429), server error (500) or unknown (404) is skipped for a while
  var models = models_(quick ? 'quick' : 'full'), res = null, code = 0, used = '', t0 = new Date().getTime(), cache2 = CacheService.getScriptCache();
  for (var mi = 0; mi < models.length; mi++) {
    if (mi > 0 && new Date().getTime() - t0 > 30000) break; // stay inside the app's 45 s limit
    used = models[mi];
    res = call_(used, body_(true)); code = res.getResponseCode();
    if (code === 400) { res = call_(used, body_(false)); code = res.getResponseCode(); } // model does not accept thinkingConfig
    if ([404, 429, 500, 503].indexOf(code) < 0) break;
    cache2.put('busy_' + used, '1', code === 404 ? 21600 : 120);
    if (code === 503 || code === 429) Utilities.sleep(500);
  }
  if (code !== 200) {
    var why = ''; try { why = JSON.parse(res.getContentText()).error.message; } catch (e) {}
    return { ok: false, error: 'AI service error (' + code + ')' + (why ? ': ' + String(why).slice(0, 160) : '') };
  }
  try {
    var body = JSON.parse(res.getContentText());
    var j = JSON.parse(body.candidates[0].content.parts[0].text);
    if (!quick) j.generic = !!j.generic;
    return { ok: true, result: j, model: used, mode: tr ? 'translate' : (quick ? 'quick' : 'full') };
  } catch (e) { return { ok: false, error: 'AI returned an unreadable answer' }; }
}

/* ---------- admin auth: email-code PIN reset, no stored passwords ---------- */
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, prop_('PIN_SALT') + '|' + s)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function isAdminEmail_(e) {
  return (prop_('ADMIN_EMAILS') || '').toLowerCase().split(',').map(function (x) { return x.trim(); }).indexOf(String(e).toLowerCase().trim()) >= 0;
}

function requestReset_(r) {
  var email = s_(r.email, 120).toLowerCase(), cache = CacheService.getScriptCache();
  var rk = 'resetrate_' + email, n = Number(cache.get(rk) || 0);
  if (n >= 5) return { ok: true };
  cache.put(rk, String(n + 1), 3600);
  if (isAdminEmail_(email)) {
    var code = ('' + Math.floor(100000 + Math.random() * 900000));
    cache.put('reset_' + email, hash_('code|' + email + '|' + code), 900);
    MailApp.sendEmail(email, 'EcoOils Project Recognition - PIN reset code',
      'Your one-time code is ' + code + '. It expires in 15 minutes.\nIf you did not ask for it, ignore this email.');
  }
  return { ok: true }; // same answer either way
}

function setPin_(r) {
  var email = s_(r.email, 120).toLowerCase(), cache = CacheService.getScriptCache();
  if (!/^\d{6,12}$/.test(String(r.pin || ''))) throw new Error('PIN must be 6 to 12 digits');
  var fk = 'fail_' + email, fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('Too many attempts. Wait 15 minutes.');
  var want = cache.get('reset_' + email);
  if (!want || want !== hash_('code|' + email + '|' + s_(r.code, 10))) {
    cache.put(fk, String(fails + 1), 900); throw new Error('Wrong or expired code');
  }
  cache.remove('reset_' + email);
  PropertiesService.getScriptProperties().setProperty('PIN_' + email, hash_('pin|' + email + '|' + r.pin));
  return { ok: true };
}

function login_(r) {
  var email = s_(r.email, 120).toLowerCase(), cache = CacheService.getScriptCache();
  var fk = 'fail_' + email, fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('Too many attempts. Wait 15 minutes.');
  var stored = prop_('PIN_' + email);
  if (!isAdminEmail_(email) || !stored) { cache.put(fk, String(fails + 1), 900); throw new Error('No PIN set yet or wrong details. Use "Forgot PIN / set PIN".'); }
  if (stored !== hash_('pin|' + email + '|' + r.pin)) { cache.put(fk, String(fails + 1), 900); throw new Error('Wrong PIN'); }
  cache.remove(fk);
  var token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('sess_' + token, email, 21600);
  return { ok: true, token: token, email: email };
}

function requireAdmin_(r) {
  var e = r.token && CacheService.getScriptCache().get('sess_' + r.token);
  if (!e) throw new Error('SESSION');
  return e;
}

/* ---------- admin data ---------- */
/* Admin-issued appreciation certificate (e.g. crew leaving) - no nomination or photos; issued straight away */
function adminCreate_(r) {
  var f = r.fields || {};
  var name = s_(f.name, 120), role = s_(f.role, 120), ic = s_(f.ic, 40), to = s_(f.to, 20), from = s_(f.from, 20);
  if (!name || !role || !ic || !/^\d{4}-\d{2}-\d{2}$/.test(to)) throw new Error('Name, position, company and last day are required');
  if (from && !/^\d{4}-\d{2}-\d{2}$/.test(from)) throw new Error('Bad joined date');
  var lang = ['id', 'en', 'ms'].indexOf(s_(f.certLang)) >= 0 ? s_(f.certLang) : 'id';
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow(), now = new Date();
    var all = n < 2 ? 0 : sh.getRange(2, COLS.indexOf('certNo') + 1, n - 1, 1).getValues().filter(function (x) { return x[0]; }).length;
    var y = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy');
    var id = 'APP-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
    var row = {
      id: id, createdAt: now.toISOString(), week: weekKey_(now), status: 'Issued', certNo: 'EOSM-REC-' + y + '-' + ('000' + (all + 1)).slice(-4),
      name: name, role: role, ic: ic, location: 'Mentari Project', eventDate: to, observation: from ? 'Joined: ' + from : 'Leaving appreciation',
      action: 'Leaving appreciation', outcome: 'Leaving appreciation', category: 'Leaving appreciation', nomName: 'HSSE', nomRole: 'HSSE Manager',
      level: 'crew', certLang: lang, certDate: Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd'), updatedAt: now.toISOString(), source: 'admin'
    };
    sh.appendRow(COLS.map(function (c) { return row[c] == null ? '' : row[c]; }));
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}
/* QR verification: token is derived from the nomination id and the secret salt; public verify returns only non-sensitive fields */
function vtok_(id) { return hash_('verify|' + id).slice(0, 12); }
function verify_(req) {
  var v = String(req.v || '').split('.'), id = v[0], tok = v[1] || '';
  if (!id || tok.length !== 12 || tok !== vtok_(id)) throw new Error('Not found');
  var rows = readAll_().filter(function (o) { return o.id === id; });
  if (!rows.length || (rows[0].status !== 'Approved' && rows[0].status !== 'Issued')) throw new Error('Not found');
  var o = rows[0];
  return { ok: true, cert: { name: o.name, role: o.role, ic: o.ic, certNo: o.certNo, certDate: o.certDate, status: o.status } };
}
function readAll_() {
  var sh = sheet_(), n = sh.getLastRow();
  if (n < 2) return [];
  var tz = SpreadsheetApp.openById(prop_('SHEET_ID')).getSpreadsheetTimeZone();
  var vals = sh.getRange(2, 1, n - 1, COLS.length).getValues();
  return vals.map(function (v) {
    var o = {}; COLS.forEach(function (c, i) {
      var x = v[i];
      if (x instanceof Date) x = (c === 'eventDate' || c === 'certDate') ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : x.toISOString();
      o[c] = x;
    });
    o.vtok = vtok_(o.id);
    return o;
  }).filter(function (o) { return o.id; });
}

function photo_(r) {
  var f = DriveApp.getFileById(s_(r.fileId, 80));
  var it = f.getParents(), ok = false, fid = prop_('FOLDER_ID');
  while (it.hasNext()) if (it.next().getId() === fid) ok = true;
  if (!ok) throw new Error('Not allowed');
  var b = f.getBlob();
  return { ok: true, uri: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) };
}

/* Delete nominations (admin only). Rows are removed from the sheet; their photos go to the Drive trash (recoverable for 30 days). */
function delete_(r) {
  var ids = (r.ids || []).map(function (x) { return String(x); });
  if (!ids.length || ids.length > 200) throw new Error('Nothing to delete');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow(), done = 0;
    if (n < 2) return { ok: true, deleted: 0 };
    var data = sh.getRange(2, 1, n - 1, COLS.length).getValues(), ib = COLS.indexOf('beforePhotos'), ia = COLS.indexOf('afterPhotos');
    for (var i = data.length - 1; i >= 0; i--) {
      if (ids.indexOf(String(data[i][0])) < 0) continue;
      String(data[i][ib] + ',' + data[i][ia]).split(',').forEach(function (f) {
        f = f.trim(); if (f) { try { DriveApp.getFileById(f).setTrashed(true); } catch (e) {} }
      });
      sh.deleteRow(i + 2); done++;
    }
    return { ok: true, deleted: done };
  } finally { lock.releaseLock(); }
}

function update_(r) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), n = sh.getLastRow();
    if (n < 2) throw new Error('Not found');
    var ids = sh.getRange(2, 1, n - 1, 1).getValues();
    var row = -1;
    for (var i = 0; i < ids.length; i++) if (ids[i][0] === r.id) { row = i + 2; break; }
    if (row < 0) throw new Error('Not found');
    var cur = sh.getRange(row, 1, 1, COLS.length).getValues()[0], rec = {};
    COLS.forEach(function (c, i) { rec[c] = cur[i]; });
    var patch = r.patch || {};
    EDITABLE.forEach(function (k) { if (patch[k] !== undefined) rec[k] = s_(patch[k], k === 'aiJson' ? 45000 : (k === 'certPara' ? 4000 : 2000)); });
    if (['Pending', 'Approved', 'Issued', 'Rejected'].indexOf(rec.status) < 0) throw new Error('Bad status');
    if ((rec.status === 'Approved' || rec.status === 'Issued') && !rec.certNo) {
      var all = sh.getRange(2, COLS.indexOf('certNo') + 1, n - 1, 1).getValues().filter(function (x) { return x[0]; }).length;
      var y = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy');
      rec.certNo = 'EOSM-REC-' + y + '-' + ('000' + (all + 1)).slice(-4);
    }
    rec.updatedAt = new Date().toISOString();
    sh.getRange(row, 1, 1, COLS.length).setValues([COLS.map(function (c) { return rec[c] == null ? '' : rec[c]; })]);
    rec.vtok = vtok_(rec.id);
    return { ok: true, row: rec };
  } finally { lock.releaseLock(); }
}
