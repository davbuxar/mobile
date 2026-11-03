/**
 * COUPON VALIDATOR BACKEND (Google Apps Script)
 * Sheet tab "Coupons" columns: A Coupon Code | B Status | C Scanned At | D Scanned By | E Note/Name (optional)
 * Row 1 = headers. Format column A as Plain Text.
 */
const SHEET_NAME = 'Coupons';
const SECRET_KEY = 'change-me-123';   // <-- change this; same key goes in the app Settings
const TZ = 'Asia/Kolkata';

function doGet(e) {
  const p = e.parameter || {};
  let out;
  try {
    if (p.key !== SECRET_KEY) out = { status: 'DENIED' };
    else if (p.action === 'verify') out = verify_(String(p.code || '').trim(), p.user || 'unknown');
    else if (p.action === 'stats') out = stats_();
    else out = { status: 'ERROR', msg: 'bad action' };
  } catch (err) {
    out = { status: 'ERROR', msg: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function verify_(code, user) {
  if (!code) return { status: 'INVALID' };
  const lock = LockService.getScriptLock();   // makes simultaneous scans safe: no coupon can pass twice
  lock.waitLock(25000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    const last = sh.getLastRow();
    if (last < 2) return { status: 'INVALID' };
    const hit = sh.getRange(2, 1, last - 1, 1).createTextFinder(code)
      .matchEntireCell(true).matchCase(true).findNext();
    if (!hit) return { status: 'INVALID' };
    const row = hit.getRow();
    const v = sh.getRange(row, 1, 1, 5).getValues()[0];
    if (String(v[1]).toUpperCase() === 'USED') {
      return { status: 'USED', at: fmt_(v[2]), by: v[3], info: v[4] };
    }
    sh.getRange(row, 2, 1, 3).setValues([['USED', new Date(), user]]);
    SpreadsheetApp.flush();
    return { status: 'OK', info: v[4] };
  } finally {
    lock.releaseLock();
  }
}

function stats_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const last = sh.getLastRow();
  if (last < 2) return { status: 'OK', total: 0, used: 0 };
  const s = sh.getRange(2, 2, last - 1, 1).getValues().flat();
  const used = s.filter(x => String(x).toUpperCase() === 'USED').length;
  return { status: 'OK', total: s.length, used: used, left: s.length - used };
}

function fmt_(d) {
  return d instanceof Date ? Utilities.formatDate(d, TZ, 'dd MMM, hh:mm a') : String(d || '');
}
