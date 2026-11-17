/**
 * Budget App — Phase 2: read BMO alerts into a Google Sheet.
 *
 * Runs inside your own Google account, where the alerts arrive.
 * It only ever searches for emails from BMO's alert address and Interac.
 * 1. Paste this file into an Apps Script project (script.google.com).
 * 2. Run setup() once and approve the permissions.
 * 3. Every 15 minutes, processAlerts() adds new transactions to the "Transactions" tab.
 *
 * Three kinds of email are read:
 *  - BMO credit card alert   -> a purchase with amount, card and merchant
 *  - BMO withdrawal alert    -> a chequing withdrawal with amount and account (no description)
 *  - Interac e-Transfer note -> names the recipient; matched to its withdrawal by amount and time
 */

const CONFIG = {
  sheetName: 'Budget App',
  checkEveryMinutes: 15,       // Apps Script allows 1, 5, 10, 15 or 30
  searchWindow: '3d',          // each run only looks back this far
  backfillWindow: '60d',       // used once by setup()
  cards: [                     // card or account ending -> person
    ['4235', 'Yassine'],
    ['1525', 'Aleksandra'],
    ['909', 'Joint'],          // BMO joint chequing account
  ],
};

const SOURCES = {
  bmo: { from: 'alertesbmo@bmo.com', dkim: 'bmo.com' },
  interac: { from: 'catch@payments.interac.ca', dkim: 'payments.interac.ca' },
};
const SUBJECT_CARD = 'Alerte carte de crédit BMO';
const SUBJECT_WITHDRAWAL = 'Alerte retrait dans votre compte de BMO';
const TRANSFER_MATCH_DAYS = 7; // a transfer can be accepted days after it was sent

const TX_HEADERS = ['Date', 'Amount', 'Card or account', 'Person', 'Merchant', 'Category', 'Source', 'Status', 'Message-ID'];
const TRANSFER_HEADERS = ['Date', 'Amount', 'Recipient', 'Reference', 'Matched withdrawal', 'Message-ID'];
const REJECT_HEADERS = ['Received', 'From', 'Subject', 'Reason', 'Message-ID'];

/** Run once: creates the sheet, the 15-minute trigger, and backfills recent alerts. */
function setup() {
  const ss = getOrCreateSpreadsheet_();
  ensureTab_(ss, 'Transactions', TX_HEADERS);
  ensureTab_(ss, 'Transfers', TRANSFER_HEADERS);
  ensureTab_(ss, 'Rejected', REJECT_HEADERS);
  ensureRulesTab_(ss);
  const cards = ensureTab_(ss, 'Cards', ['Card or account ending', 'Person']);
  cards.getRange('A:A').setNumberFormat('@');
  const known = cardMap_(ss);
  const missing = CONFIG.cards.filter(c => !(c[0] in known));
  if (missing.length) cards.getRange(cards.getLastRow() + 1, 1, missing.length, 2).setValues(missing);
  const first = ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processAlerts')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processAlerts').timeBased().everyMinutes(CONFIG.checkEveryMinutes).create();

  const added = processAlerts(CONFIG.backfillWindow);
  Logger.log('Sheet: ' + ss.getUrl());
  Logger.log('Backfill added ' + added + ' transaction(s).');
}

/** Every 15 minutes: reads new alerts, appends them, then matches e-transfers. Returns transactions added. */
function processAlerts(window) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return 0;
  try {
    const ss = getOrCreateSpreadsheet_();
    const tx = ss.getSheetByName('Transactions');
    const transfers = ss.getSheetByName('Transfers') || ensureTab_(ss, 'Transfers', TRANSFER_HEADERS);
    const rejected = ss.getSheetByName('Rejected');
    const seen = new Set(existingIds_(tx, 9).concat(existingIds_(transfers, 6), existingIds_(rejected, 5)));
    const cards = cardMap_(ss);

    const query = '(from:' + SOURCES.bmo.from + ' OR from:' + SOURCES.interac.from + ') newer_than:' +
                  (typeof window === 'string' ? window : CONFIG.searchWindow);
    const txRows = [], transferRows = [], rejects = [];

    GmailApp.search(query, 0, 200).forEach(thread => thread.getMessages().forEach(msg => {
      const id = msg.getHeader('Message-ID') || msg.getId();
      if (seen.has(id)) return;
      seen.add(id);
      const reject = r => rejects.push([msg.getDate(), msg.getFrom(), msg.getSubject(), r, id]);

      const kind = classify_(msg);
      if (!kind) return; // e.g. an incoming e-transfer or a BMO marketing email: not spending
      const reason = signatureProblem_(msg, kind === 'transfer' ? SOURCES.interac : SOURCES.bmo);
      if (reason) return reject(reason);

      if (kind === 'card') {
        const p = parseCardAlert_(msg.getBody());
        if (!p) return reject('Could not read amount, card or merchant');
        txRows.push([msg.getDate(), p.amount, p.card, cards[p.card] || 'Unknown', p.merchant,
                     '', 'BMO card alert', 'Needs review', id]);
      } else if (kind === 'withdrawal') {
        const p = parseWithdrawalAlert_(msg.getBody());
        if (!p) return reject('Could not read amount or account');
        txRows.push([msg.getDate(), p.amount, p.account, cards[p.account] || 'Unknown', '',
                     '', 'BMO withdrawal alert', 'Needs review', id]);
      } else if (kind === 'transfer') {
        const p = parseTransferNotice_(msg.getPlainBody(), msg.getSubject());
        if (!p) return reject('Could not read e-Transfer amount or recipient');
        transferRows.push([msg.getDate(), p.amount, p.recipient, p.reference, '', id]);
      }
    }));

    appendRows_(tx, txRows, TX_HEADERS.length, { 1: 'yyyy-mm-dd hh:mm', 2: '#,##0.00 $', 3: '@' });
    appendRows_(transfers, transferRows, TRANSFER_HEADERS.length, { 1: 'yyyy-mm-dd hh:mm', 2: '#,##0.00 $' });
    appendRows_(rejected, rejects, REJECT_HEADERS.length, {});
    matchTransfers_(tx, transfers);
    applyRules_(tx, ensureRulesTab_(ss));
    return txRows.length;
  } finally {
    lock.releaseLock();
  }
}

/** Which kind of email this is, or null if it isn't one the app uses. */
function classify_(msg) {
  const from = (msg.getFrom() || '').toLowerCase();
  const subject = msg.getSubject().trim();
  if (from.indexOf(SOURCES.bmo.from) !== -1) {
    if (subject === SUBJECT_CARD) return 'card';
    if (subject === SUBJECT_WITHDRAWAL) return 'withdrawal';
    return null;
  }
  if (from.indexOf(SOURCES.interac.from) !== -1) {
    // Only transfers you sent; money you received is not spending.
    if (/accepted your transfer|your transfer .* (was )?deposited/i.test(subject)) return 'transfer';
    return null;
  }
  return null;
}

/** Null if Gmail confirmed the sender's DKIM signature, otherwise the reason it failed. */
function signatureProblem_(msg, source) {
  // Gmail adds its own Authentication-Results header on top when an email arrives.
  const raw = msg.getRawContent();
  const headerBlock = raw.split(/\r?\n\r?\n/)[0].replace(/\r?\n[ \t]+/g, ' ');
  const auth = headerBlock.split(/\r?\n/).find(l => /^Authentication-Results:/i.test(l)) || '';
  const domain = source.dkim.replace(/\./g, '\\.');
  const ok = new RegExp('dkim=pass[^;]*header\\.(i|d)=@?' + domain, 'i').test(auth);
  return ok ? null : 'Signature from ' + source.dkim + ' did not verify';
}

function htmlToText_(html) {
  return decodeEntities_(
    html.replace(/<style[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, ' ')
  ).replace(/[\s  ‌]+/g, ' ');
}

/** French amount "1 234,56" -> 1234.56 */
function frenchAmount_(s) {
  const v = Number(s.replace(/ /g, '').replace(',', '.'));
  return isNaN(v) ? null : v;
}

/** BMO credit card alert: amount, card ending, merchant. */
function parseCardAlert_(html) {
  const text = htmlToText_(html);
  const amount = text.match(/Montant\s*:\s*(-?[\d .,]+?)\s*\$/);
  const card = text.match(/Se terminant par\s*(\d{3,4})/);
  const merchant = text.match(/Endroit\s*:\s*(.+?)\s*Ouvrez une session/);
  if (!amount || !card || !merchant) return null;
  const value = frenchAmount_(amount[1]);
  if (value === null) return null;
  return { amount: value, card: card[1], merchant: merchant[1].trim() };
}

/** BMO chequing withdrawal alert: amount and account ending only. */
function parseWithdrawalAlert_(html) {
  const text = htmlToText_(html);
  const amount = text.match(/Montant\s*:\s*(-?[\d .,]+?)\s*\$/);
  const account = text.match(/Compte\s*:\s*Se terminant par\s*(\d{3,4})/);
  if (!amount || !account) return null;
  const value = frenchAmount_(amount[1]);
  if (value === null) return null;
  return { amount: value, account: account[1] };
}

/** Interac notice for a transfer you sent: amount, recipient, reference. */
function parseTransferNotice_(plain, subject) {
  const text = plain.replace(/\s+/g, ' ');
  let m = text.match(/sent to (.+?) for the amount of \$([\d,]+\.\d{2})/i);
  let recipient, amount;
  if (m) { recipient = m[1]; amount = m[2]; }
  else {
    m = subject.match(/^Interac e-Transfer: (.+?) has accepted your transfer of \$([\d,]+\.\d{2})/i);
    if (!m) return null;
    recipient = m[1]; amount = m[2];
  }
  const ref = text.match(/Reference Number:\s*(\w+)/i);
  return { amount: Number(amount.replace(/,/g, '')), recipient: recipient.trim(), reference: ref ? ref[1] : '' };
}

/* ------------------------------------------------------------------ */
/* Categorization                                                       */
/* ------------------------------------------------------------------ */

const RULES_HEADERS = ['Pattern (matched against Merchant)', 'Category', 'Amount (optional)', 'Note'];

/**
 * Starting rules, checked top to bottom; the first match wins.
 * They are copied into the "Rules" tab once. After that, edit the tab, not this list:
 * add a row at the top to fix a merchant, and it applies on the next run.
 */
const DEFAULT_RULES = [
  // E-transfers (merchant is "e-Transfer to <name>" once the Interac email is matched)
  ['^e-Transfer to .*BASHKAR', 'Parking', '', 'Monthly parking'],
  ['^e-Transfer to ', 'Family support', 500, '$500 family transfers'],
  ['^e-Transfer to ', 'Other transfers', '', 'Any other e-transfer'],
  // Groceries (convenience stores included)
  ['COSTCO|LOBLAW|METRO|RCSS|SUPERSTORE|FARM BOY|FORTINOS|NO ?FRILLS|JESSE.S NF|MAXI|MARCHE ATLAS|FRESH BUY|MS\\.POTATOS|AL-MUMTAZ|FOODY WORLD|AL RAZIQ|A PAGNOTTA|COBS BREAD|SUPERMART|GREEN NATURAL|7 ELEVEN|COUCHE ?TARD|CIRCLE K STORE|VARIETY', 'Groceries', '', ''],
  // Eating out and takeout
  ['UBER ?EATS|PIZZA|SUSHI|POKE|KEBAB|DONER|DINER|CAFE|COFFEE|ESPRESSO|TIM HORTONS|OSMOWS|BURGER|THAI|DARNA|SULTAN AHMET|GONG CHA|CONG CAPHE|NIJI|PARIS BAGUETTE|TAHINIS|BARRELHOUSE|LEMONTEA|BALZAC|GELAT|DESSERT|CREPES|CREMERIE|MONTEIRO|SISTERS CORNER|SB M-WING|ADA VENDING|^SQ \\*|^TST-', 'Eating out and takeout', '', 'SQ * and TST- are Square and Toast, used mostly by cafés and restaurants'],
  // Shopping
  ['AMZN|AMAZON(?!.*PRIME)|WINNERS|MARSHALLS|HOMESENSE|CANADIAN TIRE|WAL-MART|WALMART|DOLLARAMA|H&M|HM CA|HENNES|NEWBALANCE|ARDENE|LA SENZA|VALUE VILLAGE|BEST BUY|BOOK STORE|FANFARE|LUSH|IKEA', 'Shopping', '', ''],
  // Health
  ['SHOPPERS DRUG|PHARMAPRIX|DENTISTRY|DENTAL|EYE CLINIC|LIFELABS|TELEMEDMD|ALPHA HEALTH|PHYSIO', 'Health', '', ''],
  // Transport
  ['PRESTO|PETRO-CANADA|ESSO|ULTRAMAR|SHELL|HONK|PARKING|TORPRKAUT|MTO |CARLOAD|BIKE SHARE|SWIFTCHARGE|FLO SERVICES|UBER ?TRIP|AIR-SERV|NAYAX|ACDC AUTO|REM STATION|^STM ', 'Transport', '', ''],
  // Bills and subscriptions
  ['TELUS|BELL|NETFLIX|APPLE\\.COM|GOOGLE \\*|TODOIST|ANTHROPIC|AMAZON.*PRIME|SPOTIFY|CRAVE|MOVE WITH', 'Phone, internet and subscriptions', '', ''],
  ['CARMA|KUBRA|CLAVIS|PACE PROPERTIES', 'Housing', '', ''],
  ['ECONOMICAL', 'Insurance', '', ''],
  ['GOODLIFE|HDN TRACK|RUNNERS ACADEMY|RECREATION AT U OF T|ATHLETICS', 'Sports and fitness', '', ''],
  ['BARBIER|BARBER|SALON', 'Personal care', '', ''],
  ['FAMOUS PLAYER|CINEPLEX|TORONTO INTERNATIONAL', 'Entertainment', '', ''],
  ['SAVE THE CHILDREN|FOUNDATION', 'Giving', '', ''],
  ['CANADIAN COLLEGE|CANADIAN ORGANIZATION|AAPM|ELSEVIER', 'Work: reimbursable dues', '', 'Reimbursed by Sunnybrook; left out of the budget'],
  ['TRANSAVIA|AIRBNB|AIR CANADA|PORTER|WESTJET|AIRALO|KISSANDFLY|AVIS|ENTERPRISE RENT|EXPEDIA|BOOKING\\.COM', 'Travel', '', ''],
  ['FLEXITI', 'Financing payments', '', ''],
];

/** Creates the Rules tab and fills it with DEFAULT_RULES the first time. */
function ensureRulesTab_(ss) {
  let sh = ss.getSheetByName('Rules');
  if (sh) return sh;
  sh = ensureTab_(ss, 'Rules', RULES_HEADERS);
  sh.getRange(2, 1, DEFAULT_RULES.length, 4).setValues(DEFAULT_RULES);
  sh.setColumnWidth(1, 420);
  return sh;
}

function loadRules_(rulesSheet) {
  const n = rulesSheet.getLastRow() - 1;
  if (n < 1) return [];
  return rulesSheet.getRange(2, 1, n, 3).getValues()
    .filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => {
      try {
        return { re: new RegExp(String(r[0]), 'i'), category: String(r[1]).trim(),
                 amount: r[2] === '' ? null : Number(r[2]) };
      } catch (e) {
        return null; // a typo in a pattern skips that rule instead of stopping the run
      }
    })
    .filter(Boolean);
}

/** Category for a merchant and amount, or null if no rule matches. */
function categorize_(merchant, amount, rules) {
  if (!merchant) return null;
  const rule = rules.find(r => r.re.test(merchant) &&
                               (r.amount === null || Math.abs(Number(amount) - r.amount) < 0.005));
  return rule ? rule.category : null;
}

/**
 * Fills the category of every row still marked "Needs review".
 * Rows you have categorized yourself (any other status) are never changed.
 */
function applyRules_(tx, rulesSheet) {
  const n = tx.getLastRow() - 1;
  if (n < 1) return;
  const rules = loadRules_(rulesSheet);
  const range = tx.getRange(2, 1, n, TX_HEADERS.length);
  const rows = range.getValues();
  // Write only the rows that changed, so an edit made in the sheet meanwhile is never overwritten
  rows.forEach((r, i) => {
    if (r[7] !== 'Needs review') return;
    const cat = categorize_(String(r[4]), r[1], rules);
    if (cat) tx.getRange(i + 2, 6, 1, 3).setValues([[cat, r[6], 'Auto']]);
  });
}

/** Run by hand after editing the Rules tab, to re-check everything still in "Needs review". */
function recategorize() {
  const ss = getOrCreateSpreadsheet_();
  applyRules_(ss.getSheetByName('Transactions'), ensureRulesTab_(ss));
}

/** Gives each unmatched withdrawal the recipient of a same-amount e-transfer sent around then. */
function matchTransfers_(tx, transfers) {
  const nT = transfers.getLastRow() - 1, nX = tx.getLastRow() - 1;
  if (nT < 1 || nX < 1) return;
  const tRows = transfers.getRange(2, 1, nT, TRANSFER_HEADERS.length).getValues();
  const xRows = tx.getRange(2, 1, nX, TX_HEADERS.length).getValues();
  const day = 864e5;

  tRows.forEach((t, ti) => {
    if (t[4]) return; // already matched
    const tTime = new Date(t[0]).getTime();
    let best = -1, bestGap = Infinity;
    xRows.forEach((x, xi) => {
      if (x[6] !== 'BMO withdrawal alert' || x[4]) return;
      if (Math.abs(Number(x[1]) - Number(t[1])) > 0.005) return;
      const gap = tTime - new Date(x[0]).getTime(); // notice comes after the withdrawal
      if (gap < -day || gap > TRANSFER_MATCH_DAYS * day) return;
      if (Math.abs(gap) < bestGap) { best = xi; bestGap = Math.abs(gap); }
    });
    if (best === -1) return;
    xRows[best][4] = 'e-Transfer to ' + t[2];
    tx.getRange(best + 2, 5).setValue(xRows[best][4]);
    transfers.getRange(ti + 2, 5).setValue(xRows[best][8]);
  });
}

function appendRows_(sheet, rows, width, formats) {
  if (!rows.length) return;
  rows.sort((a, b) => a[0] - b[0]);
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, width).setValues(rows);
  Object.keys(formats).forEach(col =>
    sheet.getRange(start, Number(col), rows.length, 1).setNumberFormat(formats[col]));
}

function decodeEntities_(s) {
  const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', zwnj: '',
                  eacute: 'é', egrave: 'è', agrave: 'à', ccedil: 'ç', ecirc: 'ê' };
  return s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => {
    if (e[0] === '#') {
      const code = e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return String.fromCodePoint(code);
    }
    return named[e.toLowerCase()] !== undefined ? named[e.toLowerCase()] : m;
  });
}

function getOrCreateSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create(CONFIG.sheetName);
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function ensureTab_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function existingIds_(sheet, col) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return [];
  return sheet.getRange(2, col, n, 1).getValues().map(r => String(r[0]));
}

function cardMap_(ss) {
  const sh = ss.getSheetByName('Cards');
  const map = {};
  if (!sh) return map;
  const n = sh.getLastRow() - 1;
  if (n < 1) return map;
  sh.getRange(2, 1, n, 2).getValues().forEach(r => { map[String(r[0]).trim()] = r[1]; });
  return map;
}

/** Quick check without touching the sheet: logs what the parser reads from the latest alerts. */
function testLatestAlerts() {
  const query = '(from:' + SOURCES.bmo.from + ' OR from:' + SOURCES.interac.from + ') newer_than:7d';
  GmailApp.search(query, 0, 5).forEach(thread => {
    const msgs = thread.getMessages();
    const msg = msgs[msgs.length - 1];
    const kind = classify_(msg);
    let parsed = null;
    if (kind === 'card') parsed = parseCardAlert_(msg.getBody());
    if (kind === 'withdrawal') parsed = parseWithdrawalAlert_(msg.getBody());
    if (kind === 'transfer') parsed = parseTransferNotice_(msg.getPlainBody(), msg.getSubject());
    const sig = kind ? signatureProblem_(msg, kind === 'transfer' ? SOURCES.interac : SOURCES.bmo) : '';
    Logger.log(msg.getSubject() + ' | ' + (kind || 'ignored') + ' | ' + (sig || 'OK') + ' | ' + JSON.stringify(parsed));
  });
}
