const USAGE_LOG_SHEET = 'Usage Log';
const USAGE_LOG_HEADER = ['Timestamp', 'LDAP', 'Event', 'Detail'];
const USAGE_LOG_MAX_BATCH = 50;

function ensureUsageLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(USAGE_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USAGE_LOG_SHEET);
    sheet.appendRow(USAGE_LOG_HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Appends a batch of client-buffered usage events in one write.
 * events: [{ event: 'view_change', detail: 'compare', ts: 169... }, ...]
 * LDAP is resolved server-side from the session — never trusted from the client.
 */
function logEvents(events) {
  if (!events || !events.length) return { logged: 0 };

  const email = Session.getActiveUser().getEmail();
  const ldap = email ? email.split('@')[0] : 'unknown';

  if (!checkRateLimit_('usage_log_flush', 10, 60)) {
    Logger.log('Usage log flush throttled for %s', ldap);
    return { logged: 0, throttled: true };
  }

  const trimmed = events.slice(0, USAGE_LOG_MAX_BATCH);
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000);
  if (!gotLock) return { logged: 0, throttled: true };
  try {
    const sheet = ensureUsageLogSheet_();
    const rows = trimmed.map(function (e) {
      const ts = e.ts ? new Date(e.ts) : new Date();
      return [ts, ldap, String(e.event || '').slice(0, 60), String(e.detail || '').slice(0, 200)];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, USAGE_LOG_HEADER.length).setValues(rows);
    return { logged: rows.length };
  } finally {
    lock.releaseLock();
  }
}
