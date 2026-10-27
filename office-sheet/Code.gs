/**
 * Pesto Parcel Scanner – office inbox (Google Apps Script, free).
 * Every scan from a phone becomes a row in the "Parcels" sheet.
 *
 * Setup (once):
 * 1. Create a Google Sheet, then Extensions > Apps Script, paste this file.
 * 2. Project Settings > Script properties > add TEAM_CODE = a short code you choose
 *    (colleagues type it once in the app; never post it in chat).
 * 3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 * 4. Copy the web app URL into the app's settings (or config.js).
 */
const SHEET = 'Parcels';
const HEAD = ['id', 'at', 'courier', 'tracking', 'recipient', 'sender', 'by', 'done'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}
function ok_(code) {
  const want = PropertiesService.getScriptProperties().getProperty('TEAM_CODE');
  return want && String(code || '') === want;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// GET ?code=...  -> latest 500 parcels, newest first
function doGet(e) {
  if (!ok_(e.parameter.code)) return out_({ error: 'bad_code' });
  const values = sheet_().getDataRange().getValues();
  const rows = values.slice(1).map(r => Object.fromEntries(HEAD.map((h, i) => [h, r[i]])))
    .map(r => ({ ...r, at: r.at instanceof Date ? r.at.toISOString() : String(r.at), done: r.done === true || r.done === 'TRUE' }));
  rows.reverse();
  return out_({ rows: rows.slice(0, 500) });
}

// POST body (text/plain JSON): {code, op:'add'|'done'|'delete', ...}
function doPost(e) {
  const b = JSON.parse(e.postData.contents || '{}');
  if (!ok_(b.code)) return out_({ error: 'bad_code' });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_();
    if (b.op === 'add') {
      const id = Utilities.getUuid();
      const r = b.row || {};
      sh.appendRow([id, new Date(r.at || Date.now()), r.courier || '', "'" + (r.tracking || ''), r.recipient || '', r.sender || '', r.by || '', false]);
      return out_({ id });
    }
    const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(x => x[0]);
    const i = ids.indexOf(b.id);
    if (i < 0) return out_({ error: 'not_found' });
    if (b.op === 'done') sh.getRange(i + 2, 8).setValue(!!b.done);
    if (b.op === 'delete') sh.deleteRow(i + 2);
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}
