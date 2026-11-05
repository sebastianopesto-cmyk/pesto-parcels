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
 *
 * Offices: TEAM_CODE is the head-office (admin) code and sees every office.
 *   Each office gets its own access code in the "Offices" tab (the admin adds offices from the app,
 *   Tariffe > Uffici). An office sees and changes only its own parcels and its own rates; yachts are shared.
 *   Parcels saved before offices existed belong to the first office in the list.
 *
 * Optional paid add-ons (keys live only in Script properties, never in the app or a chat):
 *   ANTHROPIC_API_KEY  -> label photos are read by Claude (much more reliable than the free reader)
 *   AI_MODEL           -> optional, default claude-opus-5-5 (e.g. claude-haiku-5-5 is ~20x cheaper)
 *   TRACK17_KEY        -> live courier status via the 17TRACK API; add a time-driven trigger
 *                         (Triggers > Add trigger > refreshTracking > Time-driven > Every hour)
 */
const TABLES = ['Parcels', 'Yachts', 'Deliveries', 'Settings', 'Offices'];
const OFFICE_TABLES = ['Parcels', 'Deliveries']; // rows carry an "office" field
const PHOTO_FOLDER = 'PIMPI foto';

// Who is calling: {admin:true} for TEAM_CODE, {office, name, city} for an office code, null otherwise
function auth_(code) {
  code = String(code || '').trim();
  if (!code) return null;
  const team = prop_('TEAM_CODE');
  if (team && code === team) return { admin: true, office: '', name: 'Sede centrale' };
  const o = rows_('Offices').find(x => x.code && String(x.code) === code && String(x.active) !== 'no');
  return o ? { admin: false, office: String(o.id), name: String(o.name || o.id), city: String(o.city || '') } : null;
}
function firstOffice_() { const o = rows_('Offices')[0]; return o ? String(o.id) : ''; }
function officeOf_(row, first) { return String(row.office || '') || first; }
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
  const me = auth_(e.parameter.code);
  if (!me) return out_({ error: 'bad_code' });
  const first = firstOffice_();
  if (e.parameter.photo) {
    if (!me.admin && !rows_('Parcels').some(r => String(r.photo) === e.parameter.photo && officeOf_(r, first) === me.office)) return out_({ error: 'not_found' });
    const f = DriveApp.getFileById(e.parameter.photo);
    if (f.getParents().next().getName() !== PHOTO_FOLDER) return out_({ error: 'not_found' });
    const b = f.getBlob();
    return out_({ data: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) });
  }
  const res = {};
  TABLES.forEach(t => res[t] = rows_(t));
  OFFICE_TABLES.forEach(t => {
    res[t].forEach(r => r.office = officeOf_(r, first));
    if (!me.admin) res[t] = res[t].filter(r => r.office === me.office);
  });
  res.Offices = res.Offices.map(o => me.admin ? o : { id: o.id, name: o.name, city: o.city }).filter(o => me.admin || String(o.id) === me.office);
  if (!me.admin) res.Settings = res.Settings.filter(x => x.id === 'rates' || x.id === 'rates@' + me.office);
  res.me = me;
  res.features = { ai: !!prop_('ANTHROPIC_API_KEY'), track: !!prop_('TRACK17_KEY') };
  return out_(res);
}

// POST body (text/plain JSON): {code, op, ...}
//   op 'put'   {table, row}      insert or update by row.id -> {id}
//   op 'del'   {table, id}
//   op 'photo' {data, name}      data URL of a JPEG/PNG -> {id}
//   op 'read'  {data, yachts}    label photo -> {fields} read by Claude (needs ANTHROPIC_API_KEY)
//   op 'track' {}                refresh live courier status now (needs TRACK17_KEY)
function doPost(e) {
  const b = JSON.parse(e.postData.contents || '{}');
  const me = auth_(b.code);
  if (!me) return out_({ error: 'bad_code' });
  if (b.op === 'read') return out_(readLabel_(b.data, b.yachts));
  if (b.op === 'track') return out_(refreshTracking());
  if (b.op === 'photo') {
    const m = String(b.data || '').match(/^data:(image\/\w+);base64,(.+)$/);
    if (!m) return out_({ error: 'bad_photo' });
    const f = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (b.name || 'foto') + (m[1] === 'image/png' ? '.png' : '.jpg')));
    return out_({ id: f.getId() });
  }
  if (TABLES.indexOf(b.table) < 0) return out_({ error: 'bad_table' });
  if (b.table === 'Offices' && !me.admin) return out_({ error: 'admin_only' });
  if (b.table === 'Settings' && !me.admin) {
    if (b.op !== 'put' || !b.row || b.row.id !== 'rates@' + me.office) return out_({ error: 'admin_only' });
  }
  const first = firstOffice_();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_(b.table);
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(x => String(x[0])) : [];
    const mine = id => { // an office may only touch its own rows
      if (me.admin || OFFICE_TABLES.indexOf(b.table) < 0) return true;
      const i = ids.indexOf(String(id)); if (i < 0) return true;
      const col = head_(sh).indexOf('office');
      const o = col < 0 ? '' : String(sh.getRange(i + 2, col + 1).getValue() || '');
      return (o || first) === me.office;
    };
    if (b.op === 'del') {
      if (!mine(b.id)) return out_({ error: 'not_yours' });
      const i = ids.indexOf(String(b.id));
      if (i >= 0) sh.deleteRow(i + 2);
      return out_({ ok: true });
    }
    if (b.op !== 'put') return out_({ error: 'bad_op' });
    const row = b.row || {};
    if (!row.id) row.id = Utilities.getUuid();
    if (!mine(row.id)) return out_({ error: 'not_yours' });
    if (OFFICE_TABLES.indexOf(b.table) >= 0) {
      if (!me.admin) row.office = me.office;
      else if (ids.indexOf(String(row.id)) < 0 && !row.office) row.office = first;
    }
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

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }

/* ---------- AI label reading (Claude API, raw HTTP: Apps Script has no SDK) ---------- */
const COURIER_IDS = ['ups', 'amazon', 'dhl', 'tnt', 'fedex', 'gls', 'brt', 'dpd', 'poste', 'imile', 'inpost', 'other', ''];
const LABEL_SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['courier', 'tracking', 'yacht', 'recipient', 'sender', 'weight_kg', 'parcel_number', 'parcels_total', 'reference', 'notes'],
  properties: {
    courier: { type: 'string', enum: COURIER_IDS, description: 'Carrier that delivers the parcel. poste also covers SDA and Poste Delivery Business. Empty if unknown.' },
    tracking: { type: 'string', description: 'Main tracking / waybill / lettera di vettura number exactly as printed, without spaces. Empty if not readable.' },
    yacht: { type: 'string', description: 'Yacht name only, without M/Y, S/Y, C/O or crew role (e.g. "MMM", "Sakura"). Empty if none.' },
    recipient: { type: 'string', description: 'Recipient block, one line per printed line.' },
    sender: { type: 'string', description: 'Sender / mittente block, one line per printed line.' },
    weight_kg: { type: 'string', description: 'Weight of this parcel in kg as a plain number with a dot (e.g. "4.0"). Empty if not printed.' },
    parcel_number: { type: 'string', description: 'Which parcel this is, e.g. "1" for 1/2. Empty if not printed.' },
    parcels_total: { type: 'string', description: 'Total parcels in the shipment, e.g. "2" for 1/2. Empty if not printed.' },
    reference: { type: 'string', description: 'Order or shipment reference (Rif.) if printed, else empty.' },
    notes: { type: 'string', description: 'Anything the office should double-check, in Italian, short. Empty if all clear.' },
  },
};
function readLabel_(dataUrl, yachts) {
  const key = prop_('ANTHROPIC_API_KEY');
  if (!key) return { error: 'no_ai_key' };
  const m = String(dataUrl || '').match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  if (!m) return { error: 'bad_photo' };
  const model = prop_('AI_MODEL') || 'claude-opus-5-5';
  const body = {
    model: model, max_tokens: 4000,
    output_config: { effort: 'low', format: { type: 'json_schema', schema: LABEL_SCHEMA } },
    system: 'You read courier shipping labels photographed in the parcel office of Pesto Sea Group, a yacht agency in Genoa, Italy. ' +
      'Parcels are addressed to superyachts c/o Pesto Sea Group, Calata Molo Vecchio. Copy codes exactly as printed; never guess digits you cannot see. ' +
      'If a barcode has its number printed under it, that number is usually the tracking number.',
    messages: [{ role: 'user', content: [
      { type: 'image', source: { type: 'base64', media_type: m[1], data: m[2] } },
      { type: 'text', text: 'Read this label.' + (yachts && yachts.length ? ' Yachts we already know (use the same spelling if it is one of them): ' + yachts.slice(0, 200).join(', ') + '.' : '') },
    ] }],
  };
  const headers = { 'x-api-key': key, 'anthropic-version': '2026-10-21' };
  if (/^claude-(opus-5|opus-5-5|sonnet-5-5|fable-5-1)$/.test(model)) { body.fallbacks = 'default'; headers['anthropic-beta'] = 'server-side-fallback-2026-07-01'; }
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  let j; try { j = JSON.parse(res.getContentText()); } catch (e) { return { error: 'ai_http_' + res.getResponseCode() }; }
  if (res.getResponseCode() !== 200) return { error: 'ai_' + ((j.error && j.error.type) || res.getResponseCode()) };
  if (j.stop_reason === 'refusal') return { error: 'ai_refusal' };
  if (j.stop_reason === 'max_tokens') return { error: 'ai_truncated' };
  const text = (j.content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  try { return { fields: JSON.parse(text), model: j.model }; } catch (e) { return { error: 'ai_bad_json' }; }
}

/* ---------- Live tracking (17TRACK API v2.2) ---------- */
// Parcels with a tracking number that are expected or arrived in the last days get their status refreshed.
// Writes trackStatus (e.g. InTransit, OutForDelivery, Delivered), trackEvent and trackAt on the Parcels row.
function refreshTracking() {
  const key = prop_('TRACK17_KEY');
  if (!key) return { error: 'no_track_key' };
  const rows = rows_('Parcels').filter(r => r.tracking && (r.state === 'expected' || (r.state === 'in' && r.trackStatus !== 'Delivered')) && r.trackStatus !== 'Delivered');
  const list = rows.slice(0, 40);
  if (!list.length) return { ok: true, checked: 0 };
  const call = (path, payload) => {
    const res = UrlFetchApp.fetch('https://api.17track.net/track/v2.2/' + path, {
      method: 'post', contentType: 'application/json', headers: { '17token': key }, payload: JSON.stringify(payload), muteHttpExceptions: true,
    });
    try { return JSON.parse(res.getContentText()); } catch (e) { return { code: res.getResponseCode() }; }
  };
  const nums = list.map(r => ({ number: String(r.tracking) }));
  const fresh = list.filter(r => !r.trackAt).map(r => ({ number: String(r.tracking) }));
  if (fresh.length) call('register', fresh); // already-registered numbers are simply rejected, which is fine
  const info = call('gettrackinfo', nums);
  const got = ((info.data && info.data.accepted) || []);
  const sh = sheet_('Parcels');
  got.forEach(a => {
    const r = list.find(x => String(x.tracking) === String(a.number)); if (!r) return;
    const t = a.track_info || {}, st = (t.latest_status && t.latest_status.status) || '', ev = t.latest_event || {};
    putRow_(sh, { id: r.id, trackStatus: st, trackEvent: [ev.description, ev.location].filter(Boolean).join(' · '), trackAt: new Date().toISOString() });
  });
  return { ok: true, checked: list.length, updated: got.length };
}
function putRow_(sh, row) {
  let head = head_(sh);
  const extra = Object.keys(row).filter(k => head.indexOf(k) < 0);
  if (extra.length) { sh.getRange(1, head.length + 1, 1, extra.length).setValues([extra]); head = head.concat(extra); }
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(x => String(x[0]));
  const i = ids.indexOf(String(row.id)); if (i < 0) return;
  head.forEach((h, k) => { if (h in row) sh.getRange(i + 2, k + 1).setNumberFormat('@').setValue(row[h] == null ? '' : String(row[h])); });
}
