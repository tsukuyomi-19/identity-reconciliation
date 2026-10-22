/**
 * Configuration lives in two places, deliberately:
 *
 *   Script Properties  - secrets only (API keys, client secrets, tenant ids).
 *                        Never in source, never in the spreadsheet, never in
 *                        git. Set them under Project Settings > Script
 *                        Properties.
 *
 *   The "Settings" tab - non-secret knobs an administrator should be able to
 *                        change without opening the script editor (which
 *                        source is authoritative, what an OU path looks like
 *                        when an account is disabled, and so on).
 *
 * The split matters: the people who tune the rules are usually not the people
 * who should be handling credentials.
 */

var SECRET_KEYS = {
  MS_TENANT_ID: 'MS_TENANT_ID',
  MS_CLIENT_ID: 'MS_CLIENT_ID',
  MS_CLIENT_SECRET: 'MS_CLIENT_SECRET',
  ROSTER_SHEET_ID: 'ROSTER_SHEET_ID'
};

var SETTING_DEFAULTS = {
  // Which provider tab is treated as the authoritative list of people.
  // Every Master row is anchored on a record from this source; records that
  // exist only in other sources surface as orphans in Findings instead.
  anchorSource: 'GoogleWorkspace',

  // Some directories express "disabled" by moving the account into a
  // particular org unit rather than by setting a suspended flag. Any OU path
  // matching this regex counts as disabled in addition to the explicit flag.
  disabledOuPattern: '_disabled|/suspended|/archive',

  // Status values from the roster source that mean "no longer with us".
  // Matching is by exclusion: anything NOT matching this is treated as
  // active, because roster systems rarely spell out "active" explicitly.
  rosterInactivePattern: 'withdraw|graduat|inactive|terminat|separat|drop|expell',

  // How stale a last-login timestamp has to be before it is worth reporting,
  // in days. Set to 0 to disable the dormant-account rule.
  dormantAfterDays: '120'
};

function getSecret_(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error(
      'Missing Script Property "' + key + '". Set it under Project Settings > Script Properties.'
    );
  }
  return value;
}

function getSecretOptional_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * Reads the Settings tab into a plain object, falling back to the defaults
 * above for anything missing. A provider or rule that depends on a setting
 * should call this rather than reading the sheet itself, so an absent or
 * half-filled Settings tab still produces a working run.
 */
function getSettings() {
  var settings = {};
  Object.keys(SETTING_DEFAULTS).forEach(function (key) {
    settings[key] = SETTING_DEFAULTS[key];
  });

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  if (!sheet) {
    return settings;
  }

  readTable_(SHEETS.SETTINGS).forEach(function (row) {
    var key = String(row.key || '').trim();
    var value = String(row.value === undefined ? '' : row.value).trim();
    if (key && value) {
      settings[key] = value;
    }
  });

  return settings;
}

/** Writes the Settings tab with current values, creating it if absent. */
function initializeSettingsTab() {
  var settings = getSettings();
  var rows = Object.keys(SETTING_DEFAULTS).map(function (key) {
    return [key, settings[key], SETTING_DESCRIPTIONS[key] || ''];
  });
  writeTable_(SHEETS.SETTINGS, ['key', 'value', 'description'], rows);
  return rows.length;
}

var SETTING_DESCRIPTIONS = {
  anchorSource: 'Provider tab treated as the authoritative roster of people.',
  disabledOuPattern: 'Regex (case-insensitive). An OU path matching this counts as disabled.',
  rosterInactivePattern: 'Regex (case-insensitive). Roster statuses matching this mean inactive.',
  dormantAfterDays: 'Days without a login before an enabled account is reported. 0 disables.'
};
