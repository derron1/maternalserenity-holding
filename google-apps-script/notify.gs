// "Notify me" sign-ups for the maternalserenity.co.uk holding page.
// Paste into the Sheet's Extensions → Apps Script editor, then
// Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).

const SHEET_NAME = 'Signups';

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Honeypot field that people never see; bots fill it. Pretend success.
  if (p.website) return json({ ok: true });

  const email = String(p.email || '').trim().toLowerCase();
  const valid =
    email.length <= 254 &&
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) &&
    !/^[=+\-@]/.test(email); // a leading = + - @ would make Sheets treat the cell as a formula
  if (!valid) return json({ ok: false, error: 'invalid_email' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) sheet.appendRow(['Signed up at', 'Email']);

    const last = sheet.getLastRow();
    const existing = last > 1 ? sheet.getRange(2, 2, last - 1, 1).getValues().flat() : [];
    if (!existing.includes(email)) sheet.appendRow([new Date(), email]);
  } finally {
    lock.releaseLock();
  }

  return json({ ok: true });
}

function json(body) {
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}
