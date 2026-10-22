/**
 * Roster provider - the system that knows whether a person should still have
 * accounts at all. An HR system, a student information system, a CRM: it does
 * not matter which, as long as it can export to a Google Sheet.
 *
 * Reading an exported Sheet rather than calling the roster API directly is a
 * deliberate trade. Roster vendors tend to have rate-limited, awkwardly
 * paginated, or paid APIs, and most of them will already export to a Sheet on
 * a schedule. Pointing at that export costs one Script Property, consumes no
 * API quota, and means this tool needs no credentials for the roster system
 * at all - it only needs Viewer access to one spreadsheet.
 *
 * Set ROSTER_SHEET_ID to that spreadsheet's id. Every tab in it is scanned,
 * and any tab whose header row contains both an email column and a status
 * column is treated as roster data. That way a newly added tab (a second
 * population, a new campus, a new department) is picked up with no code
 * change.
 */

var ROSTER_SOURCE = 'Roster';

var ROSTER_HEADERS = ['email', 'status', 'active', 'externalId', 'sourceTab'];

function isRosterConfigured() {
  return !!getSecretOptional_(SECRET_KEYS.ROSTER_SHEET_ID);
}

function syncRoster() {
  if (!isRosterConfigured()) {
    Logger.log('Roster sync skipped - ROSTER_SHEET_ID not set.');
    return 0;
  }

  var settings = getSettings();
  var inactivePattern = new RegExp(settings.rosterInactivePattern, 'i');
  var external = SpreadsheetApp.openById(getSecret_(SECRET_KEYS.ROSTER_SHEET_ID));
  var seen = {};
  var records = [];

  external.getSheets().forEach(function (sheet) {
    var rows = readSheetAsObjects_(sheet);
    if (!rows.length) {
      return;
    }

    var emailKey = findHeaderKey_(rows[0], 'email');
    var statusKey = findHeaderKey_(rows[0], 'status');
    if (!emailKey || !statusKey) {
      return; // not a roster tab
    }
    var idKey = findHeaderKey_(rows[0], 'id') || findHeaderKey_(rows[0], 'externalid');

    rows.forEach(function (row) {
      var email = normalizeEmail_(row[emailKey]);
      // Roster exports commonly repeat a person once per related contact
      // (each guardian, each emergency contact). Keep the first occurrence
      // and drop the rest, or one person becomes five rows downstream.
      if (!email || seen[email]) {
        return;
      }
      seen[email] = true;

      var status = String(row[statusKey] || '');
      records.push({
        email: email,
        status: status,
        // Active is determined by exclusion. Roster systems label the
        // departed precisely ("Withdrawn", "Terminated") but describe the
        // present tense inconsistently ("Enrolled", "Current", "FT", blank),
        // so anything that does not match the inactive pattern is treated as
        // active rather than guessed at.
        active: !inactivePattern.test(status),
        externalId: idKey ? String(row[idKey] || '') : '',
        sourceTab: sheet.getName()
      });
    });
  });

  writeTable_(ROSTER_SOURCE, ROSTER_HEADERS, recordsToRows_(records, ROSTER_HEADERS));
  return records.length;
}

function readRosterRecords() {
  return readTable_(ROSTER_SOURCE).map(function (row) {
    return {
      email: normalizeEmail_(row.email),
      status: String(row.status || ''),
      active: parseBoolean_(row.active),
      externalId: String(row.externalId || ''),
      sourceTab: String(row.sourceTab || ''),
      source: ROSTER_SOURCE
    };
  }).filter(function (record) {
    return !!record.email;
  });
}
