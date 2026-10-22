/**
 * Manual-import provider, for the long tail of SaaS products that have no
 * usable read API.
 *
 * This is more common than vendor marketing suggests. Plenty of platforms
 * expose a SCIM endpoint that is write-only by design - built for an identity
 * provider to push provisioning INTO the product, not for a third party to
 * read users back out. Others offer a CSV export in the admin console and
 * nothing else. Either way the account list exists; it just cannot be
 * fetched on a schedule.
 *
 * Rather than leave those systems out of the reconciliation entirely - which
 * is how orphaned licenses go unnoticed for years - this provider accepts a
 * pasted export.
 *
 * Usage: export the account list from the product's admin console and paste
 * it unmodified into a tab named "Import: <ServiceName>". Keep whatever
 * columns the vendor gave you; only an email column is required, and it is
 * located by header name rather than position. The text after "Import: "
 * becomes the service name in reports, so "Import: Zoom" and
 * "Import: Adobe" both work with no code change.
 */

var IMPORT_TAB_PREFIX = 'Import: ';

/**
 * Returns a map of email -> array of service names that person has an account
 * in, assembled from every Import: tab in the active spreadsheet.
 */
function getServiceAssignments() {
  var assignments = {};

  SpreadsheetApp.getActive().getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    if (name.indexOf(IMPORT_TAB_PREFIX) !== 0) {
      return;
    }
    var service = name.substring(IMPORT_TAB_PREFIX.length).trim();
    if (!service) {
      return;
    }

    var rows = readSheetAsObjects_(sheet);
    if (!rows.length) {
      return;
    }
    var emailKey = findHeaderKey_(rows[0], 'email');
    if (!emailKey) {
      Logger.log('Tab "' + name + '" has no email column - skipped.');
      return;
    }

    rows.forEach(function (row) {
      var email = normalizeEmail_(row[emailKey]);
      if (!email) {
        return;
      }
      if (!assignments[email]) {
        assignments[email] = [];
      }
      if (assignments[email].indexOf(service) === -1) {
        assignments[email].push(service);
      }
    });
  });

  return assignments;
}

/** Names of every service currently represented by an Import: tab. */
function getImportedServiceNames() {
  return SpreadsheetApp.getActive().getSheets()
    .map(function (sheet) { return sheet.getName(); })
    .filter(function (name) { return name.indexOf(IMPORT_TAB_PREFIX) === 0; })
    .map(function (name) { return name.substring(IMPORT_TAB_PREFIX.length).trim(); })
    .filter(Boolean);
}
