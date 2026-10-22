/**
 * A one-screen summary for whoever is not going to scroll the Findings tab.
 *
 * Counts only - the detail lives in Master and Findings. The point is to make
 * a trend visible: if "disabled but still enabled elsewhere" is nonzero every
 * week, the offboarding process is broken, and that is a different
 * conversation from any individual row.
 */

function buildDashboardTab() {
  var model = loadReconciliationModel();
  var findings = readTable_(SHEETS.FINDINGS);

  var bySeverity = { high: 0, medium: 0, low: 0 };
  var byRule = {};
  findings.forEach(function (row) {
    var severity = String(row.severity || 'low');
    if (bySeverity[severity] !== undefined) {
      bySeverity[severity] += 1;
    }
    var rule = String(row.rule || 'Unclassified');
    byRule[rule] = (byRule[rule] || 0) + 1;
  });

  var enabled = model.people.filter(function (p) { return p.anchorEnabled; }).length;

  var rows = [
    ['People (anchor source)', model.people.length],
    ['  enabled', enabled],
    ['  disabled', model.people.length - enabled],
    ['Microsoft records', Object.keys(model.microsoft).length],
    ['Roster records', Object.keys(model.roster).length],
    ['People with service assignments', Object.keys(model.services).length],
    ['Imported services', getImportedServiceNames().join(', ') || '(none)'],
    ['', ''],
    ['Findings - high', bySeverity.high],
    ['Findings - medium', bySeverity.medium],
    ['Findings - low', bySeverity.low],
    ['', '']
  ];

  Object.keys(byRule).sort(function (a, b) {
    return byRule[b] - byRule[a];
  }).forEach(function (rule) {
    rows.push([rule, byRule[rule]]);
  });

  rows.push(['', '']);
  rows.push(['Last refreshed', new Date()]);

  writeTable_(SHEETS.DASHBOARD, ['metric', 'value'], rows);
  styleHeader_(getOrCreateSheet_(SHEETS.DASHBOARD), 2);
  return rows.length;
}
