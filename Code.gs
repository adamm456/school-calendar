/**
 * שרת הנתונים של לו"ז בית הספר.
 * - בלי מפתח: מחזיר רק את גיליון החודש הנוכחי.
 * - עם מפתח תקין מהגיליון Access: מחזיר את כל החודשים.
 * הקובץ הזה צריך לשבת בתוך קובץ ה-Google Sheets (Extensions > Apps Script).
 */

// ===== הגדרות =====
const PAGE_URL = '';            // כתובת האתר ב-GitHub Pages, למשל https://name.github.io/school-calendar/
const ACCESS_SHEET = 'Access';  // גיליון המפתחות (שם, מפתח, פעיל)
const TZ = 'Asia/Jerusalem';
const MONTH_NAMES = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];

function doGet(e) {
  const key = ((e && e.parameter && e.parameter.key) || '').trim();
  const full = key !== '' && isValidKey_(key);
  if (key !== '' && !full) return json_({ error: 'bad key' });

  let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(function (s) { return s.getName() !== ACCESS_SHEET; });

  if (!full) {
    const month = MONTH_NAMES[Number(Utilities.formatDate(new Date(), TZ, 'M')) - 1];
    sheets = sheets.filter(function (s) { return s.getName() === month; });
  }
  return json_({
    mode: full ? 'full' : 'public',
    tabs: sheets.map(function (s) { return { name: s.getName(), csv: toCsv_(s) }; })
  });
}

function isValidKey_(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCESS_SHEET);
  if (!sh || sh.getLastRow() < 2) return false;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === key && String(rows[i][2]).trim() !== 'לא') return true;
  }
  return false;
}

function toCsv_(sheet) {
  const rows = sheet.getDataRange().getDisplayValues();
  while (rows.length && rows[rows.length - 1].every(function (c) { return c === ''; })) rows.pop();
  return rows.map(function (r) {
    return r.map(function (c) {
      return /[",\n\r]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c;
    }).join(',');
  }).join('\n');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===== תפריט ניהול מפתחות בתוך הגיליון =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('גישה למורשים')
    .addItem('יצירת מפתח חדש', 'createKey')
    .addToUi();
}

function createKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('שם האדם שמקבל גישה מלאה');
  if (r.getSelectedButton() !== ui.Button.OK || !r.getResponseText().trim()) return;
  const name = r.getResponseText().trim();
  const key = Utilities.getUuid().replace(/-/g, '');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ACCESS_SHEET);
  if (!sh) { sh = ss.insertSheet(ACCESS_SHEET); sh.appendRow(['שם', 'מפתח', 'פעיל']); }
  sh.appendRow([name, key, 'כן']);
  ui.alert('המפתח נוצר',
    'שלחו ל' + name + ' את הקישור הבא:\n\n' +
    (PAGE_URL ? PAGE_URL + '#key=' + key : '(PAGE_URL לא הוגדר בקוד)\nהמפתח: ' + key),
    ui.ButtonSet.OK);
}
