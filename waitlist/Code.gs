/**
 * Hero Made Smart waitlist: receives the landing page's form POST and adds one row to the
 * "Signups" tab of the Google Sheet this script is bound to.
 * Deploy: Extensions > Apps Script > Deploy > New deployment > Web app,
 * Execute as: Me, Who has access: Anyone. Put the /exec URL into ENDPOINT in index.html.
 */

const SHEET_NAME = 'Signups';
const HEADERS = ['time_utc', 'email', 'interest', 'test_name', 'price', 'phone', 'adult', 'source'];

const ALLOWED = {
  interest: ['memory', 'puzzles', 'both', 'test-prep'],
  price: ['free', '5', '10', '15+'],
  phone: ['iphone', 'android'],
};

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Spam trap: real people never fill the hidden "website" field.
  if (p.website) return reply_(true);

  const email = clean_(p.email, 200).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return reply_(false, 'bad email');
  if (p.adult !== 'yes') return reply_(false, 'age not confirmed');

  const row = [
    new Date().toISOString(),
    email,
    pick_(p.interest, ALLOWED.interest),
    clean_(p.test_name, 80),
    pick_(p.price, ALLOWED.price),
    pick_(p.phone, ALLOWED.phone),
    'yes',
    clean_(p.source, 60),
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return reply_(true);
}

function doGet() {
  return reply_(true, 'Hero Made Smart waitlist is running.');
}

function getSheet_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = book.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = book.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Trims, caps length, and stops spreadsheet formula tricks (=, +, -, @ at the start).
function clean_(value, max) {
  let s = String(value || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// Keeps only known answers; anything else becomes blank.
function pick_(value, allowed) {
  return allowed.indexOf(String(value || '')) >= 0 ? String(value) : '';
}

function reply_(ok, note) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, note: note || '' }))
    .setMimeType(ContentService.MimeType.JSON);
}
