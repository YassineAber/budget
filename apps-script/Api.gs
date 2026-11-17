/**
 * Budget App — data API for the phone app hosted on GitHub Pages.
 * Bound to the "Budget App" sheet (Extensions → Apps Script). Replaces Dashboard.gs / Dashboard.html.
 *
 * Setup:
 * 1. Paste this file and appsscript.json, then run createKey() once and copy the key from the log.
 * 2. Deploy → New deployment → Web app. Execute as: Me. Who has access: Anyone.
 * 3. Open the app's setup page and paste the web app URL and the key.
 *
 * Every request must carry the key, so the URL alone gives no access to the data.
 */

const EXCLUDED = /^Excluded|^Work: reimbursable/; // not counted as household spending
const UNCATEGORIZED = 'Needs review';
const LIMITS_HEADERS = ['Category', 'Monthly limit'];
const DEFAULT_LIMITS = [['Groceries', 900], ['Eating out and takeout', 600], ['Shopping', 1000]];

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function limitsTab_(ss) {
  let sh = ss.getSheetByName('Limits');
  if (!sh) {
    sh = ss.insertSheet('Limits');
    sh.getRange(1, 1, 1, 2).setValues([LIMITS_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, DEFAULT_LIMITS.length, 2).setValues(DEFAULT_LIMITS);
  }
  return sh;
}

function readLimits_(ss) {
  const sh = limitsTab_(ss);
  const n = sh.getLastRow() - 1;
  const out = {};
  if (n < 1) return out;
  sh.getRange(2, 1, n, 2).getValues().forEach(r => {
    const name = String(r[0]).trim();
    if (name && r[1] !== '' && !isNaN(Number(r[1]))) out[name] = Number(r[1]);
  });
  return out;
}

function readTransactions_(ss) {
  const sh = ss.getSheetByName('Transactions');
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const tz = ss.getSpreadsheetTimeZone();
  return sh.getRange(2, 1, n, 9).getValues().map((r, i) => ({
    row: i + 2,
    date: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "yyyy-MM-dd'T'HH:mm") : String(r[0]),
    amount: Number(r[1]) || 0,
    account: String(r[2]),
    person: String(r[3]),
    merchant: String(r[4]),
    category: String(r[5]) || UNCATEGORIZED,
    status: String(r[7]),
    id: String(r[8]),
  }));
}

function knownCategories_(ss, limits) {
  const set = new Set(Object.keys(limits));
  const rules = ss.getSheetByName('Rules');
  if (rules && rules.getLastRow() > 1) {
    rules.getRange(2, 2, rules.getLastRow() - 1, 1).getValues().forEach(r => r[0] && set.add(String(r[0]).trim()));
  }
  ['Guest suite', 'Travel', 'Other transfers', 'Excluded: card payment', 'Other'].forEach(c => set.add(c));
  return Array.from(set).sort();
}

/** Days between two ISO dates, inclusive. */
function daysBetween_(startIso, endIso) {
  return Math.round((Date.parse(endIso + 'T12:00:00Z') - Date.parse(startIso + 'T12:00:00Z')) / 864e5) + 1;
}

/**
 * How many monthly limits fit in the range: whole calendar months count exactly
 * (3 months = 3x the limit); any other range is pro-rated by days (a week = 7/30.4 of it).
 */
function limitFactor_(startIso, endIso) {
  const [sy, sm, sd] = startIso.split('-').map(Number);
  const [ey, em, ed] = endIso.split('-').map(Number);
  const lastDay = new Date(ey, em, 0).getDate();
  if (sd === 1 && ed === lastDay) return (ey - sy) * 12 + (em - sm) + 1;
  return daysBetween_(startIso, endIso) * 12 / 365.25;
}

/** Pure calculation, kept separate so it can be tested outside Apps Script. */
function summarize_(transactions, limits, startIso, endIso, todayIso) {
  const inRange = transactions.filter(t => t.date.slice(0, 10) >= startIso && t.date.slice(0, 10) <= endIso);
  const spent = {};
  inRange.forEach(t => {
    if (EXCLUDED.test(t.category)) return;
    spent[t.category] = (spent[t.category] || 0) + t.amount;
  });
  const days = daysBetween_(startIso, endIso);
  const isCurrent = todayIso >= startIso && todayIso <= endIso;
  const elapsed = isCurrent ? daysBetween_(startIso, todayIso) : (todayIso > endIso ? days : 0);
  const share = elapsed / days; // fraction of the period gone
  const factor = limitFactor_(startIso, endIso);
  const round = v => Math.round(v * 100) / 100;

  const names = new Set(Object.keys(limits).concat(Object.keys(spent)));
  const categories = Array.from(names).map(name => {
    const limit = name in limits ? round(limits[name] * factor) : null;
    const s = round(spent[name] || 0);
    let state = 'none';
    if (limit !== null) {
      if (s > limit) state = 'over';
      else if (isCurrent && s > limit * Math.min(1, share + 0.1)) state = 'ahead'; // 10% of slack before warning
      else state = 'ok';
    }
    return { name, spent: s, limit, monthlyLimit: name in limits ? limits[name] : null, state,
             pace: limit !== null ? round(limit * share) : null };
  });
  categories.sort((a, b) => (b.limit !== null) - (a.limit !== null) || b.spent - a.spent);

  const budgeted = categories.filter(c => c.limit !== null);
  return {
    start: startIso, end: endIso, days, elapsed, isCurrent, factor,
    totals: {
      limit: round(budgeted.reduce((a, c) => a + c.limit, 0)),
      spentBudgeted: round(budgeted.reduce((a, c) => a + c.spent, 0)),
      spentAll: round(categories.reduce((a, c) => a + c.spent, 0)),
    },
    categories,
    needsReview: inRange.filter(t => t.category === UNCATEGORIZED).length,
    transactions: inRange,
  };
}

/** Everything the page needs for a date range ("yyyy-MM-dd"). Blank = this month. */
function getDashboard(startIso, endIso) {
  const ss = sheet_();
  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (!startIso || !endIso) {
    const [y, m] = today.split('-').map(Number);
    startIso = today.slice(0, 8) + '01';
    endIso = today.slice(0, 8) + String(new Date(y, m, 0).getDate()).padStart(2, '0');
  }
  if (startIso > endIso) throw new Error('The start date is after the end date');
  const limits = readLimits_(ss);
  const summary = summarize_(readTransactions_(ss), limits, startIso, endIso, today);
  summary.transactions.sort((a, b) => (a.date < b.date ? 1 : -1));
  summary.categoryOptions = knownCategories_(ss, limits);
  summary.today = today;
  return summary;
}

/** Sets a category's monthly limit. A blank value removes the limit. */
function setLimit(category, value) {
  category = String(category || '').trim();
  if (!category) throw new Error('Missing category');
  const clear = value === '' || value === null || value === undefined;
  const amount = clear ? '' : Number(value);
  if (!clear && (isNaN(amount) || amount < 0)) throw new Error('The limit must be a positive number');

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sh = limitsTab_(sheet_());
    const n = sh.getLastRow() - 1;
    const names = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim()) : [];
    const i = names.indexOf(category);
    if (i >= 0) {
      if (clear) sh.deleteRow(i + 2);
      else sh.getRange(i + 2, 2).setValue(Math.round(amount * 100) / 100);
    } else if (!clear) {
      sh.appendRow([category, Math.round(amount * 100) / 100]);
    }
  } finally {
    lock.releaseLock();
  }
  return true;
}

/**
 * Changes one transaction's category and marks it Reviewed so the script never overwrites it.
 * With alwaysForMerchant, also adds a rule at the top of the Rules tab.
 */
function setCategory(row, id, category, alwaysForMerchant) {
  return setCategories([{ row: row, id: id }], category, alwaysForMerchant);
}

/** Same as setCategory for several transactions at once (items: [{row, id}]). */
function setCategories(items, category, alwaysForMerchant) {
  category = String(category || '').trim();
  if (!category) throw new Error('Choose a category');
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const ss = sheet_();
    const sh = ss.getSheetByName('Transactions');
    const merchants = new Set();
    items.forEach(it => {
      if (String(sh.getRange(it.row, 9).getValue()) !== String(it.id)) {
        throw new Error('The sheet changed. Refresh and try again.');
      }
    });
    items.forEach(it => {
      sh.getRange(it.row, 6).setValue(category);
      sh.getRange(it.row, 8).setValue('Reviewed');
      const merchant = String(sh.getRange(it.row, 5).getValue()).trim();
      if (merchant) merchants.add(merchant);
    });
    if (alwaysForMerchant && merchants.size) {
      const rules = ss.getSheetByName('Rules');
      const rows = Array.from(merchants).map(m =>
        ['^' + m.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), category, '', 'Added from the dashboard']);
      rules.insertRowsBefore(2, rows.length);
      rules.getRange(2, 1, rows.length, 4).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
  return true;
}


/* ------------------------------------------------------------------ */
/* API                                                                  */
/* ------------------------------------------------------------------ */

/** Run once: creates the secret key the phone app sends with every request. */
function createKey() {
  const bytes = Utilities.getUuid() + Utilities.getUuid();
  const key = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)).replace(/=+$/, '');
  PropertiesService.getScriptProperties().setProperty('API_KEY', key);
  Logger.log('Your app key (keep it private): ' + key);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function checkKey_(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!expected) throw new Error('No key yet: run createKey() in Apps Script.');
  if (!key || String(key) !== expected) throw new Error('Wrong key');
}

/** Reads: ?key=…&action=dashboard&start=yyyy-MM-dd&end=yyyy-MM-dd */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    checkKey_(p.key);
    if (p.action === 'dashboard') return json_({ ok: true, data: getDashboard(p.start, p.end) });
    if (p.action === 'ping') return json_({ ok: true });
    throw new Error('Unknown action');
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

/** Changes: a JSON body {key, action: "setLimit"|"setCategories", ...}. */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    checkKey_(body.key);
    if (body.action === 'setLimit') setLimit(body.category, body.value);
    else if (body.action === 'setCategories') setCategories(body.items, body.category, body.always);
    else throw new Error('Unknown action');
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}
