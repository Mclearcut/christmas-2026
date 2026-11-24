/**
 * Christmas 2026 dashboard: Google Sheets backend (v2).
 * Lives in Extensions > Apps Script of the "Christmas 2026 presents" Sheet.
 * Deploy > Manage deployments > edit > New version (keeps the same URL).
 * After pasting, run setup() once from the editor (select it, click Run).
 */

const PLANNER_ID = 'PASTE_PLANNER_SHEET_ID_HERE';
const PRESENTS_ID = 'PASTE_PRESENTS_SHEET_ID_HERE';

// Wishlist columns (1 based). Row 4 holds headers, items start on row 5.
const W = { cat: 1, item: 2, size: 3, notes: 4, who: 5, bought: 6, qty: 7, price: 8, link: 9, type: 10 };
const FIRST_ROW = 5;
const CLAIMS_TAB = 'Claims';
const CLAIM_HEADERS = ['When', 'List', 'Row', 'Item', 'Name', 'Amount £', 'Bought'];

function presents_() { return SpreadsheetApp.openById(PRESENTS_ID); }
function planner_() { return SpreadsheetApp.openById(PLANNER_ID); }

function tab_(ss, name) {
  const norm = s => s.toLowerCase().replace(/[‘’']/g, '').replace(/\s+/g, ' ').trim();
  const found = ss.getSheets().find(sh => norm(sh.getName()) === norm(name));
  if (!found) throw new Error('Tab "' + name + '" not found in "' + ss.getName() + '". Tabs: ' + ss.getSheets().map(sh => sh.getName()).join(', '));
  return found;
}

function claimsTab_() {
  const ss = presents_();
  let sh = ss.getSheets().find(s => s.getName() === CLAIMS_TAB);
  if (!sh) {
    sh = ss.insertSheet(CLAIMS_TAB);
    sh.getRange(1, 1, 1, CLAIM_HEADERS.length).setValues([CLAIM_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ---------- Web endpoints ----------

function doGet() {
  try { return json_(readAll_()); }
  catch (err) { return json_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents);
    const name = String(req.name || '').trim().slice(0, 40);
    let result;
    if (req.action === 'claim') result = claim_(req, name);
    else if (req.action === 'unclaim') result = unclaim_(req, name);
    else if (req.action === 'bought') result = bought_(req, name);
    else if (req.action === 'suggest') result = suggest_(req, name);
    else result = { ok: false, error: 'Unknown action' };
    result.data = readAll_();
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Actions ----------

function isFree_(v) {
  v = String(v || '').trim();
  return v === '' || v.toUpperCase() === 'TBC';
}

function claimRows_(list, row) {
  const sh = claimsTab_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, CLAIM_HEADERS.length).getValues()
    .map((v, i) => ({ sheetRow: i + 2, when: v[0], list: v[1], row: Number(v[2]), item: v[3], name: String(v[4]), amount: v[5], bought: v[6] === true || String(v[6]).toLowerCase() === 'yes' }))
    .filter(c => (!list || c.list === list) && (!row || c.row === row));
}

function logClaim_(list, row, item, name, amount) {
  claimsTab_().appendRow([new Date(), list, row, item, name, amount || '', false]);
}

function refreshWho_(row) {
  const sh = tab_(presents_(), "Boys' wishlist");
  const claims = claimRows_('wishlist', row);
  sh.getRange(row, W.who).setValue(claims.map(c => c.name).join(', '));
  sh.getRange(row, W.bought).setValue(claims.length && claims.every(c => c.bought) ? 'Yes' : '');
}

function checkLabel_(sheet, row, col, label) {
  return String(sheet.getRange(row, col).getValue()) === label;
}

function claim_(req, name) {
  if (!name) return { ok: false, error: 'Please add your name' };
  const row = Number(req.row);
  if (req.list === 'wishlist') {
    const sh = tab_(presents_(), "Boys' wishlist");
    if (!checkLabel_(sh, row, W.item, req.label)) return { ok: false, error: 'The list changed, please refresh' };
    const type = String(sh.getRange(row, W.type).getValue()).toLowerCase();
    const qty = Number(sh.getRange(row, W.qty).getValue()) || 1;
    const taken = claimRows_('wishlist', row).length;
    if (type !== 'money' && taken >= qty) return { ok: false, error: 'Someone has just taken the last one' };
    const amount = type === 'money' ? Math.max(0, Math.min(1000, Number(req.amount) || 0)) : '';
    logClaim_('wishlist', row, req.label, name, amount);
    refreshWho_(row);
    return { ok: true };
  }
  const t = singleTarget_(req.list);
  if (!checkLabel_(t.sheet, row, t.labelCol, req.label)) return { ok: false, error: 'The list changed, please refresh' };
  const cell = t.sheet.getRange(row, t.nameCol);
  if (!isFree_(cell.getValue())) return { ok: false, error: 'Someone has just taken this one' };
  cell.setValue(name);
  logClaim_(req.list, row, req.label, name, '');
  return { ok: true };
}

function unclaim_(req, name) {
  const row = Number(req.row);
  const mine = claimRows_(req.list, row).filter(c => c.name.toLowerCase() === name.toLowerCase());
  if (req.list === 'wishlist') {
    if (!mine.length) return { ok: false, error: 'Only the person who took it can undo' };
    claimsTab_().deleteRow(mine[mine.length - 1].sheetRow);
    refreshWho_(row);
    return { ok: true };
  }
  const t = singleTarget_(req.list);
  if (!checkLabel_(t.sheet, row, t.labelCol, req.label)) return { ok: false, error: 'The list changed, please refresh' };
  const cell = t.sheet.getRange(row, t.nameCol);
  if (String(cell.getValue()).trim().toLowerCase() !== name.toLowerCase()) return { ok: false, error: 'Only the person who took it can undo' };
  cell.setValue('TBC');
  if (mine.length) claimsTab_().deleteRow(mine[mine.length - 1].sheetRow);
  return { ok: true };
}

function bought_(req, name) {
  const row = Number(req.row);
  const mine = claimRows_('wishlist', row).filter(c => c.name.toLowerCase() === name.toLowerCase());
  if (!mine.length) return { ok: false, error: 'Only the person who took it can tick this' };
  const value = !mine.every(c => c.bought);
  mine.forEach(c => claimsTab_().getRange(c.sheetRow, 7).setValue(value));
  refreshWho_(row);
  return { ok: true };
}

function suggest_(req, name) {
  if (!name) return { ok: false, error: 'Please add your name' };
  const sh = tab_(presents_(), "Boys' wishlist");
  const item = String(req.item || '').trim().slice(0, 80);
  if (!item) return { ok: false, error: 'Please add the gift' };
  const link = /^https?:\/\//i.test(String(req.link || '')) ? String(req.link).slice(0, 300) : '';
  sh.appendRow(['Suggested', item, String(req.size || '').slice(0, 40), 'Suggested by ' + name + (req.notes ? '. ' + String(req.notes).slice(0, 200) : ''), '', '', 1, '', link, '']);
  return { ok: true };
}

function singleTarget_(list) {
  if (list === 'menu') return { sheet: tab_(planner_(), 'Menu'), labelCol: 2, nameCol: 3 };
  if (list === 'bring') return { sheet: tab_(planner_(), 'Shopping'), labelCol: 2, nameCol: 5 };
  throw new Error('Unknown list');
}

// ---------- Read ----------

function readAll_() {
  const presents = presents_();
  const planner = planner_();
  const claims = claimRows_('wishlist');
  const wishlist = rows_(tab_(presents, "Boys' wishlist"), FIRST_ROW, 10)
    .filter(r => r.v[1] && !String(r.v[1]).startsWith('Add your own'))
    .map(r => ({
      row: r.row, category: r.v[0], item: r.v[1], size: r.v[2], notes: r.v[3],
      qty: Number(r.v[6]) || 1, price: r.v[7], link: r.v[8], type: String(r.v[9]).toLowerCase(),
      // Names and bought ticks only. Money amounts are never sent to the page.
      claims: claims.filter(c => c.row === r.row).map(c => ({ name: c.name, bought: c.bought }))
    }));
  const menu = rows_(tab_(planner, 'Menu'), 4, 5)
    .filter(r => r.v[1]).map(r => ({ row: r.row, course: r.v[0], dish: r.v[1], buys: r.v[2], cooks: r.v[3], notes: r.v[4] }));
  const schedule = rows_(tab_(planner, 'Schedule'), 4, 4)
    .filter(r => r.v[0]).map(r => ({ day: r.v[0], when: r.v[1], what: r.v[2], notes: r.v[3] }));
  const bring = rows_(tab_(planner, 'Shopping'), 5, 5)
    .filter(r => r.v[0] === 'Activities' || r.v[0] === 'Other')
    .map(r => ({ row: r.row, item: r.v[1], qty: r.v[2], who: r.v[4] }));
  return { wishlist, menu, schedule, bring };
}

function rows_(sheet, firstRow, cols) {
  const n = sheet.getLastRow() - firstRow + 1;
  if (n < 1) return [];
  return sheet.getRange(firstRow, 1, n, cols).getDisplayValues().map((v, i) => ({ row: firstRow + i, v }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Daily email to the owner ----------

function dailySummary() {
  const since = Date.now() - 24 * 3600 * 1000;
  const recent = claimRows_().filter(c => c.when instanceof Date && c.when.getTime() >= since);
  if (!recent.length) return;
  const lines = recent.map(c => {
    const what = c.list === 'wishlist' ? 'gift' : c.list === 'menu' ? 'dish' : 'bring along';
    return '- ' + c.name + ': ' + c.item + ' (' + what + (c.amount ? ', £' + c.amount : '') + ')';
  });
  const data = readAll_();
  const giftsFree = data.wishlist.filter(w => w.type !== 'money' && w.claims.length < w.qty).length;
  const foodFree = data.menu.filter(m => isFree_(m.buys)).length + data.bring.filter(b => isFree_(b.who)).length;
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Christmas 2026: ' + recent.length + ' new today',
    'New in the last 24 hours:\n' + lines.join('\n') + '\n\nStill free: ' + giftsFree + ' gifts, ' + foodFree + ' food and bring along items.');
}

// ---------- One off setup: run once from the editor ----------

// Example options researched in November 2026. Prices are approximate and only fill empty cells.
const SEED = {
  'Jeans': [15, 'https://www.next.co.uk/style/su728807/841304'],
  'T shirts / shirts': [13, 'https://www.marksandspencer.com/3-pack-pure-cotton-plain-t-shirts-6-16-yrs-/p/clp61268487'],
  'Tracksuits (to run around in)': [35, 'https://www.next.co.uk/style/su413625/aw6205'],
  'Hoody jumpers': [18, 'https://www.next.co.uk/shop/boys/clothing/sweatshirts-and-hoodies'],
  'Casual trousers (chino style)': [12, 'https://www.marksandspencer.com/cotton-rich-chinos-6-16-yrs-/p/clp60580090'],
  'Wellies': ['', 'https://www.marksandspencer.com/kids-wellies-8-small-2-large-/p/clp23009542'],
  'Trainers for the park': [30, 'https://www.sportsdirect.com/vans-brooklyn-trainers-juniors-055046'],
  'Black school trainers': [26, 'https://www.next.co.uk/style/su588942/ap7791'],
  'Superhero slippers': [15, 'https://www.marksandspencer.com/kids-spider-man-slip-on-slippers-4-small-2-large-/p/clp61229605'],
  'Winter boots (smart / casual)': [30, 'https://www.next.co.uk/shop/gender-olderboys-gender-youngerboys-productaffiliation-footwear/category-boots'],
  'Black winter jacket for school': [34, 'https://www.marksandspencer.com/stormwear-padded-coat-6-16-yrs-/p/clp60618873'],
  'Winter jacket for weekends / park': [30, 'https://www.marksandspencer.com/stormwear-padded-hooded-coat-2-8-yrs-/p/clp60736669'],
  'Parkas / bombers with hoods': [16.5, 'https://tuclothing.sainsburys.co.uk/product/tuc140372640'],
  'Park bags': [8.49, 'https://www.decathlon.co.uk/p/kids-hiking-backpack-10-l-7-15-years-mh500-blue-green/308403/c261m8975746'],
  'Winter hat, scarf and gloves': ['', 'https://www.mountainwarehouse.com/p/mf87f0/thmo/kids-thinsulate-winter-hat-gloves-set/'],
  'Winter pyjamas': [24, 'https://www.marksandspencer.com/2pk-pure-cotton-checked-pyjama-sets-6-16-yrs-/p/clp60753297'],
  'Vests': [9.5, 'https://www.marksandspencer.com/5-pack-pure-cotton-plain-vests-2-14-yrs-/p/clp60736446'],
  'Socks': ['', 'https://www.marksandspencer.com/5pk-cotton-gamer-socks/p/clp60505082'],
  'Boxers': [13, 'https://www.marksandspencer.com/5pk-cotton-with-stretch-gamer-trunks-5-16-yrs-/p/clp60727948'],
  'Electric circuit /science game/toy': [36, 'https://cpc.farnell.com/snap-circuits/sc-110/snap-circuits-jr-100-experiments/dp/ED00458'],
  'Lego': [13, 'https://hmv.com/store/pop-culture/lego/hogwarts-castle-sorting-hat-ceremony-harry-potter'],
  'Puzzles, how to draw or colouring books': [10, 'https://www.foyles.co.uk/book/how-to-draw/dk/9780241457580'],
  'Telescope': [25, 'https://shop.sciencemuseum.org.uk/products/adventure-telescope'],
  'Maths practice books': [6.5, 'https://www.theportobellobookshop.com/9781847621900'],
  'Fun audio books': [10, 'https://www.mabelandfox.com/collections/playtime/products/yoto-yoto-card-little-leaders-exceptional-men-in-black-history'],
  'Board games': [14, 'https://www.very.co.uk/browse/toys-games/character--marvel'],
  'Character single bedding': [18, 'https://www.argos.co.uk/product/7615106'],
  'Phonics, reading and maths games': [18, 'https://www.argos.co.uk/product/9870554']
};
const TWO_OF = ['Clothes', 'Footwear', 'Outerwear', 'Sleepwear and underwear'];

function setup() {
  const sh = tab_(presents_(), "Boys' wishlist");
  sh.getRange(4, W.qty, 1, 4).setValues([['How many', 'Price £ (approx)', 'Example link', 'Type']]).setFontWeight('bold');
  const n = sh.getLastRow() - FIRST_ROW + 1;
  const vals = sh.getRange(FIRST_ROW, 1, n, 10).getValues();
  claimsTab_();
  vals.forEach((v, i) => {
    const row = FIRST_ROW + i, item = String(v[1]).trim();
    if (!item) return;
    if (v[6] === '' || v[6] === null) sh.getRange(row, W.qty).setValue(TWO_OF.indexOf(String(v[0])) >= 0 ? 2 : 1);
    if (/^money/i.test(item) && !v[9]) sh.getRange(row, W.type).setValue('money');
    const seed = SEED[item];
    if (seed) {
      if (v[7] === '' && seed[0] !== '') sh.getRange(row, W.price).setValue(seed[0]);
      if (v[8] === '') sh.getRange(row, W.link).setValue(seed[1]);
    }
    // Move any names typed straight into "Who's buying" into the Claims log.
    const who = String(v[4]).trim();
    if (who && !claimRows_('wishlist', row).length) {
      who.split(',').map(s => s.trim()).filter(Boolean).forEach(nm => logClaim_('wishlist', row, item, nm, ''));
    }
  });
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'dailySummary')) {
    ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(19).create();
  }
}
