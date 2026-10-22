/**
 * Presentation only. Nothing here changes what the data says - it just makes
 * the difference between "fine" and "needs attention" visible at a glance,
 * because a report nobody can skim is a report nobody reads.
 */

var SEVERITY_COLORS = {
  high: '#f4c7c3',
  medium: '#fce8b2',
  low: '#e8eaed'
};

function formatMasterTab_() {
  var sheet = getOrCreateSheet_(SHEETS.MASTER);
  var rows = Math.max(sheet.getLastRow() - 1, 0);
  if (!rows) {
    return;
  }

  styleHeader_(sheet, MASTER_HEADERS.length);

  // Highlight the disagreements rather than coloring everything: a person
  // disabled in the primary directory whose other accounts are still live.
  var anchorCol = MASTER_HEADERS.indexOf('anchorEnabled') + 1;
  var msCol = MASTER_HEADERS.indexOf('microsoftEnabled') + 1;
  var range = sheet.getRange(2, 1, rows, MASTER_HEADERS.length);

  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND($' + columnLetter_(anchorCol) + '2=FALSE, $' + columnLetter_(msCol) + '2=TRUE)'
    )
    .setBackground(SEVERITY_COLORS.high)
    .setRanges([range])
    .build();

  sheet.setConditionalFormatRules([rule]);
  sheet.autoResizeColumns(1, MASTER_HEADERS.length);
}

function formatFindingsTab_() {
  var sheet = getOrCreateSheet_(SHEETS.FINDINGS);
  var rows = Math.max(sheet.getLastRow() - 1, 0);
  if (!rows) {
    return;
  }

  styleHeader_(sheet, FINDINGS_HEADERS.length);

  var range = sheet.getRange(2, 1, rows, FINDINGS_HEADERS.length);
  var rules = Object.keys(SEVERITY_COLORS).map(function (severity) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="' + severity + '"')
      .setBackground(SEVERITY_COLORS[severity])
      .setRanges([range])
      .build();
  });

  sheet.setConditionalFormatRules(rules);
  sheet.autoResizeColumns(1, FINDINGS_HEADERS.length);
}

function styleHeader_(sheet, columnCount) {
  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#f1f3f4');
  sheet.setFrozenRows(1);
}

function columnLetter_(index) {
  var letter = '';
  var n = index;
  while (n > 0) {
    var remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - remainder) / 26);
  }
  return letter;
}
