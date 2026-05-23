/**
 * AppLog.js — Structured logging ke sheet "AppLog"
 *
 * Sheet AppLog: A=Timestamp, B=UserEmail, C=Action, D=Detail, E=Status
 * Sheet dibuat otomatis jika belum ada.
 */


function logActivity_(userEmail, action, detail, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let   sh = ss.getSheetByName(SHEET_LOG);

    if (!sh) {
      sh = ss.insertSheet(SHEET_LOG);
      sh.appendRow(['Timestamp', 'UserEmail', 'Action', 'Detail', 'Status']);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    }

    sh.appendRow([
      new Date(),
      (userEmail || '').toString().toLowerCase(),
      (action    || '').toString(),
      (detail    || '').toString(),
      (status    || 'OK').toString()
    ]);
  } catch (e) {
    // Logging tidak boleh break main flow
    Logger.log('logActivity_ ERROR: ' + e);
  }
}


function getRecentLogs_(limit) {
  limit = Math.max(1, Math.min(parseInt(limit) || 50, 500));
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_LOG);
    if (!sh || sh.getLastRow() < 2) return [];

    const last  = sh.getLastRow();
    const start = Math.max(2, last - limit + 1);
    const count = last - start + 1;

    return sh.getRange(start, 1, count, 5).getValues()
      .reverse()
      .map(r => ({
        timestamp: r[0] instanceof Date
          ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd HH:mm:ss')
          : String(r[0] || ''),
        userEmail: String(r[1] || ''),
        action:    String(r[2] || ''),
        detail:    String(r[3] || ''),
        status:    String(r[4] || '')
      }));
  } catch (e) {
    Logger.log('getRecentLogs_ ERROR: ' + e);
    return [];
  }
}
