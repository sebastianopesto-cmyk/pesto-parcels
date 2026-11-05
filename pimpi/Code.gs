/**
 * PIMPI – office archive (Google Apps Script, free).
 * Tabs in the sheet: Parcels, Yachts, Deliveries, Settings. Columns are added automatically
 * when the app sends a new field, so app updates rarely need this script to change.
 * Label photos and signatures are saved as image files in a Drive folder "PIMPI foto".
 *
 * Setup (once):
 * 1. Create a Google Sheet, then Extensions > Apps Script, paste this file.
 * 2. Project Settings > Script properties > add TEAM_CODE = a short code you choose
 *    (colleagues type it once in the app; never post it in chat).
 * 3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 * 4. Copy the web app URL into the app's settings (or config.js).
 * Updating: Deploy > Manage deployments > Edit > Version: New version > Deploy (same URL).
 */
const TABLES = ['Parcels', 'Yachts', 'Deliveries', 'Settings'];
const PHOTO_FOLDER = 'PIMPI foto';

function ok_(code) {
  const want = PropertiesService.getScriptProperties().getProperty('TEAM_CODE');
  return want && String(code || '') === want;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(['id']); sh.setFrozenRows(1); }
  return sh;
}
function head_(sh) {
  return sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
}
function rows_(name) {
  const sh = sheet_(name), head = head_(sh), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, head.length).getValues().map(r => {
    const o = {};
    head.forEach((h, i) => { if (h) o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
    return o;
  }).filter(o => o.id !== '' && o.id != null);
}
function folder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

// GET ?code=...                -> {Parcels, Yachts, Deliveries, Settings}
// GET ?code=...&photo=<fileId> -> {data: 'data:image/jpeg;base64,...'}
function doGet(e) {
  if (!ok_(e.parameter.code)) return out_({ error: 'bad_code' });
  if (e.parameter.photo) {
    const f = DriveApp.getFileById(e.parameter.photo);
    if (f.getParents().next().getName() !== PHOTO_FOLDER) return out_({ error: 'not_found' });
    const b = f.getBlob();
    return out_({ data: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) });
  }
  const res = {};
  TABLES.forEach(t => res[t] = rows_(t));
  return out_(res);
}

// POST body (text/plain JSON): {code, op, ...}
//   op 'put'   {table, row}      insert or update by row.id -> {id}
//   op 'del'   {table, id}
//   op 'photo' {data, name}      data URL of a JPEG/PNG -> {id}
function doPost(e) {
  const b = JSON.parse(e.postData.contents || '{}');
  if (!ok_(b.code)) return out_({ error: 'bad_code' });
  if (b.op === 'photo') {
    const m = String(b.data || '').match(/^data:(image\/\w+);base64,(.+)$/);
    if (!m) return out_({ error: 'bad_photo' });
    const f = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (b.name || 'foto') + (m[1] === 'image/png' ? '.png' : '.jpg')));
    return out_({ id: f.getId() });
  }
  if (TABLES.indexOf(b.table) < 0) return out_({ error: 'bad_table' });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_(b.table);
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(x => String(x[0])) : [];
    if (b.op === 'del') {
      const i = ids.indexOf(String(b.id));
      if (i >= 0) sh.deleteRow(i + 2);
      return out_({ ok: true });
    }
    if (b.op !== 'put') return out_({ error: 'bad_op' });
    const row = b.row || {};
    if (!row.id) row.id = Utilities.getUuid();
    let head = head_(sh);
    const extra = Object.keys(row).filter(k => head.indexOf(k) < 0);
    if (extra.length) { sh.getRange(1, head.length + 1, 1, extra.length).setValues([extra]); head = head.concat(extra); }
    let i = ids.indexOf(String(row.id));
    const r = i < 0 ? sh.getLastRow() + 1 : i + 2;
    const range = sh.getRange(r, 1, 1, head.length);
    const cur = i < 0 ? head.map(() => '') : range.getValues()[0].map(v => v instanceof Date ? v.toISOString() : v);
    const vals = head.map((h, k) => (h in row) ? (row[h] == null ? '' : String(row[h])) : cur[k]);
    range.setNumberFormat('@').setValues([vals]); // keep everything as text: dates and tracking codes stay exact
    return out_({ id: row.id });
  } finally { lock.releaseLock(); }
}
