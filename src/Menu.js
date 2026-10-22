/**
 * Entry points: the spreadsheet menu, the full-refresh orchestration, and the
 * scheduled trigger.
 *
 * Every step runs independently and catches its own failures. A permission
 * hiccup on one source should not stop the others from refreshing, and should
 * not stop Master and Findings from rebuilding off whatever data IS
 * available. A partial report that says which part is stale beats no report.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Reconciliation')
    .addItem('Run full refresh', 'runFullRefresh')
    .addSeparator()
    .addItem('Sync Google Workspace', 'syncGoogleWorkspace')
    .addItem('Sync Microsoft', 'syncMicrosoft')
    .addItem('Sync roster', 'syncRoster')
    .addSeparator()
    .addItem('Rebuild Master', 'buildMasterTab')
    .addItem('Rebuild Findings', 'buildFindingsTab')
    .addItem('Rebuild Dashboard', 'buildDashboardTab')
    .addSeparator()
    .addItem('First-time setup', 'initializeWorkbook')
    .addItem('Install 6-hour refresh', 'installRefreshTrigger')
    .addToUi();
}

function runFullRefresh() {
  var summary = [];
  var errors = [];

  function step(label, fn) {
    try {
      summary.push(label + ': ' + fn());
    } catch (err) {
      summary.push(label + ': FAILED');
      errors.push(label + ' - ' + err.message);
      Logger.log(label + ' failed: ' + err.message + '\n' + (err.stack || ''));
    }
  }

  step('Google', syncGoogleWorkspace);
  step('Microsoft', syncMicrosoft);
  step('Roster', syncRoster);
  step('Master', buildMasterTab);
  step('Findings', buildFindingsTab);
  step('Dashboard', buildDashboardTab);

  var message = summary.join('  |  ');
  if (errors.length) {
    message += '   ERRORS: ' + errors.join('; ');
  }

  try {
    SpreadsheetApp.getActive().toast(message, 'Reconciliation', 15);
  } catch (ignored) {
    // No UI when running from a trigger; the log is the record there.
  }
  Logger.log(message);
  return message;
}

/**
 * Creates the Settings and LicenseRules tabs so a new deployment has
 * something to edit rather than a blank spreadsheet and a README.
 */
function initializeWorkbook() {
  initializeSettingsTab();

  if (!SpreadsheetApp.getActive().getSheetByName(SHEETS.RULES)) {
    writeTable_(SHEETS.RULES, ['orgPathPrefix', 'expectedLicense'], []);
  }

  SpreadsheetApp.getActive().toast(
    'Created Settings and LicenseRules. Add your secrets under Project Settings > Script Properties, then run a full refresh.',
    'Reconciliation',
    15
  );
}

function installRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runFullRefresh') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('runFullRefresh').timeBased().everyHours(6).create();
  return 'Refresh trigger installed (every 6 hours).';
}
