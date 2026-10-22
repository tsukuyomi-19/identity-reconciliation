/**
 * Thin layer over SpreadsheetApp so the rest of the project speaks in rows of
 * objects rather than in A1 notation.
 *
 * Two things here are load-bearing:
 *
 *   - Reads and writes are batched. Apps Script charges a round trip per
 *     getValue()/setValue() call, so a per-cell loop over a few thousand
 *     accounts will blow the six-minute execution limit. Everything below
 *     uses getValues()/setValues() exactly once.
 *
 *   - Email is the join key everywhere, so it is normalized in exactly one
 *     place. Directory systems disagree about case and whitespace, and a
 *     single stray capital letter silently turns one person into two rows.
 */

var SHEETS = {
  SETTINGS: 'Settings',
  MASTER: 'Master',
  FINDINGS: 'Findings',
  DASHBOARD: 'Dashboard',
  RULES: 'LicenseRules'
};

function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function normalizeEmail_(value) {
  return String(value === undefined || value === null ? '' : value).trim().toLowerCase();
}

/**
 * Finds a header key case-insensitively and ignoring surrounding whitespace.
 * Source systems export "Email", "email", and "E-mail Address" with equal
 * enthusiasm, so callers name the concept and let this resolve the column.
 */
function findHeaderKey_(sampleRow, wanted) {
  var target = String(wanted).toLowerCase().replace(/[^a-z0-9]/g, '');
  var match = null;
  Object.keys(sampleRow).forEach(function (key) {
    if (match) {
      return;
    }
    if (String(key).toLowerCase().replace(/[^a-z0-9]/g, '') === target) {
      match = key;
    }
  });
  return match;
}

/** Reads any sheet as an array of objects keyed by its header row. */
function readSheetAsObjects_(sheet) {
  if (!sheet) {
    return [];
  }
  var range = sheet.getDataRange();
  if (range.getNumRows() < 2) {
    return [];
  }
  var values = range.getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  return values.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (header, i) {
      if (header) {
        obj[header] = row[i];
      }
    });
    return obj;
  });
}

/** Reads a named tab as objects. Returns [] when the tab does not exist. */
function readTable_(name) {
  return readSheetAsObjects_(SpreadsheetApp.getActive().getSheetByName(name));
}

/**
 * Replaces a tab's contents with the given headers and rows in one write.
 * Clearing by range rather than clear() preserves column widths and any
 * conditional formatting an administrator has set up by hand.
 */
function writeTable_(name, headers, rows) {
  var sheet = getOrCreateSheet_(name);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow > 0 && lastCol > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clearContent();
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.setFrozenRows(1);
  return rows.length;
}

/** Converts a provider's record objects into rows matching `headers`. */
function recordsToRows_(records, headers) {
  return records.map(function (record) {
    return headers.map(function (header) {
      var value = record[header];
      if (value === undefined || value === null) {
        return '';
      }
      if (value instanceof Date) {
        return value;
      }
      if (typeof value === 'object') {
        return JSON.stringify(value);
      }
      return value;
    });
  });
}

function parseBoolean_(value) {
  if (typeof value === 'boolean') {
    return value;
  }
  var text = String(value).trim().toLowerCase();
  if (text === 'true' || text === 'yes' || text === '1') {
    return true;
  }
  if (text === 'false' || text === 'no' || text === '0') {
    return false;
  }
  return null; // genuinely unknown, which is different from false
}

function daysSince_(value) {
  if (!value) {
    return null;
  }
  var then = value instanceof Date ? value : new Date(value);
  if (isNaN(then.getTime()) || then.getTime() === 0) {
    return null;
  }
  return Math.floor((Date.now() - then.getTime()) / 86400000);
}
