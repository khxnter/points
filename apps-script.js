// Paste this into Extensions > Apps Script in a Google Sheet.
// Then Deploy > New deployment > Web app:
//   Execute as: Me
//   Who has access: Anyone
// Copy the web app URL into ENDPOINT in index.html.
// Columns used: id | date | kieran | liam | note

function sheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function doGet() {
  const tz = Session.getScriptTimeZone();
  const rows = sheet().getDataRange().getValues()
    .filter(r => r[0] !== '')
    .map(r => ({
      id: String(r[0]),
      date: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
      k: Number(r[2]) || 0,
      l: Number(r[3]) || 0,
      note: String(r[4] || '')
    }));
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const p = e.parameter;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet();
    if (p.action === 'delete') {
      const ids = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === String(p.id)) { s.deleteRow(i + 1); break; }
      }
    } else {
      const k = parseFloat(p.k) || 0;
      const l = parseFloat(p.l) || 0;
      if (!k && !l) return ContentService.createTextOutput('bad input');
      s.appendRow([String(p.id), String(p.date), k, l, String(p.note || '').slice(0, 80)]);
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}
