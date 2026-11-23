/**
 * Christmas 2026 dashboard: Google Sheets backend.
 * Paste into Extensions > Apps Script inside the "Christmas 2026 presents" Sheet.
 * Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 */

// The planner Sheet (menu, schedule, shopping). Presents is this Sheet itself.
const PLANNER_ID = 'PASTE_PLANNER_SHEET_ID_HERE';

function doGet() {
  return json_(readAll_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents);
    const name = String(req.name || '').trim().slice(0, 40);
    let result = { ok: true };
    if (req.action === 'claim') result = claim_(req, name);
    else if (req.action === 'unclaim') result = unclaim_(req, name);
    else if (req.action === 'suggest') result = suggest_(req, name);
    else if (req.action === 'santa') result = santa_(req, name);
    else result = { ok: false, error: 'Unknown action' };
    result.data = readAll_();
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Where each list lives: sheet, first data row, column holding the label, column holding the claimer.
function target_(list) {
  const presents = SpreadsheetApp.getActiveSpreadsheet();
  const planner = SpreadsheetApp.openById(PLANNER_ID);
  if (list === 'wishlist') return { sheet: presents.getSheetByName("Boys' wishlist"), labelCol: 2, nameCol: 5 };
  if (list === 'menu') return { sheet: planner.getSheetByName('Menu'), labelCol: 2, nameCol: 3 };
  if (list === 'bring') return { sheet: planner.getSheetByName('Shopping'), labelCol: 2, nameCol: 5 };
  throw new Error('Unknown list');
}

function isFree_(v) {
  v = String(v || '').trim();
  return v === '' || v.toUpperCase() === 'TBC';
}

function claim_(req, name) {
  if (!name) return { ok: false, error: 'Please add your name' };
  const t = target_(req.list);
  const row = Number(req.row);
  if (String(t.sheet.getRange(row, t.labelCol).getValue()) !== req.label) return { ok: false, error: 'The list changed, please refresh' };
  const cell = t.sheet.getRange(row, t.nameCol);
  if (!isFree_(cell.getValue())) return { ok: false, error: 'Someone has just taken this one' };
  cell.setValue(name);
  return { ok: true };
}

function unclaim_(req, name) {
  const t = target_(req.list);
  const row = Number(req.row);
  if (String(t.sheet.getRange(row, t.labelCol).getValue()) !== req.label) return { ok: false, error: 'The list changed, please refresh' };
  const cell = t.sheet.getRange(row, t.nameCol);
  if (String(cell.getValue()).trim().toLowerCase() !== name.toLowerCase()) return { ok: false, error: 'Only the person who took it can undo' };
  cell.setValue(req.list === 'wishlist' ? '' : 'TBC');
  return { ok: true };
}

function suggest_(req, name) {
  if (!name) return { ok: false, error: 'Please add your name' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Boys' wishlist");
  const item = String(req.item || '').trim().slice(0, 80);
  if (!item) return { ok: false, error: 'Please add the gift' };
  sh.appendRow(['Suggested', item, String(req.size || '').slice(0, 40), 'Suggested by ' + name + (req.notes ? '. ' + String(req.notes).slice(0, 200) : ''), '', '']);
  return { ok: true };
}

function santa_(req, name) {
  if (!name) return { ok: false, error: 'Please add your name' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Secret Santa');
  const values = sh.getRange(5, 1, Math.max(sh.getLastRow() - 4, 1), 3).getValues();
  const hints = String(req.hints || '').slice(0, 300);
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === name.toLowerCase()) {
      sh.getRange(5 + i, 2, 1, 2).setValues([[hints, 'Yes']]);
      return { ok: true };
    }
  }
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === '') {
      sh.getRange(5 + i, 1, 1, 3).setValues([[name, hints, 'Yes']]);
      return { ok: true };
    }
  }
  sh.appendRow([name, hints, 'Yes']);
  return { ok: true };
}

function readAll_() {
  const presents = SpreadsheetApp.getActiveSpreadsheet();
  const planner = SpreadsheetApp.openById(PLANNER_ID);
  const wishlist = rows_(presents.getSheetByName("Boys' wishlist"), 5, 6)
    .filter(r => r.v[1] && !String(r.v[1]).startsWith('Add your own'))
    .map(r => ({ row: r.row, category: r.v[0], item: r.v[1], size: r.v[2], notes: r.v[3], by: r.v[4] }));
  const santa = rows_(presents.getSheetByName('Secret Santa'), 5, 3)
    .filter(r => r.v[0]).map(r => ({ name: r.v[0] }));
  const menu = rows_(planner.getSheetByName('Menu'), 4, 5)
    .filter(r => r.v[1]).map(r => ({ row: r.row, course: r.v[0], dish: r.v[1], buys: r.v[2], cooks: r.v[3], notes: r.v[4] }));
  const schedule = rows_(planner.getSheetByName('Schedule'), 4, 4)
    .filter(r => r.v[0]).map(r => ({ day: r.v[0], when: r.v[1], what: r.v[2], notes: r.v[3] }));
  const bring = rows_(planner.getSheetByName('Shopping'), 5, 5)
    .filter(r => r.v[0] === 'Activities' || r.v[0] === 'Other')
    .map(r => ({ row: r.row, item: r.v[1], qty: r.v[2], who: r.v[4] }));
  return { wishlist, menu, schedule, bring, santa };
}

function rows_(sheet, firstRow, cols) {
  const n = sheet.getLastRow() - firstRow + 1;
  if (n < 1) return [];
  return sheet.getRange(firstRow, 1, n, cols).getDisplayValues().map((v, i) => ({ row: firstRow + i, v }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
