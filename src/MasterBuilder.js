/**
 * Builds the Master tab: one row per person, every system's view of them side
 * by side.
 *
 * Master is anchored on one authoritative source (see the anchorSource
 * setting). A record that exists only in a secondary system is deliberately
 * NOT a Master row - Master answers "what is this person's status across our
 * systems", and an account with no corresponding person is a different kind
 * of problem. Those surface as orphans in Findings, where they belong.
 */

var MASTER_HEADERS = [
  'email',
  'fullName',
  'orgPath',
  'anchorEnabled',
  'lastLoginDays',
  'microsoftEnabled',
  'microsoftLicenses',
  'rosterActive',
  'rosterStatus',
  'services'
];

function buildMasterTab() {
  var model = loadReconciliationModel();
  var rows = model.people.map(function (person) {
    return [
      person.email,
      person.fullName,
      person.orgPath,
      person.anchorEnabled,
      person.lastLoginDays === null ? '' : person.lastLoginDays,
      person.microsoft ? person.microsoft.enabled : '',
      person.microsoft ? person.microsoft.licenses : '',
      person.roster ? person.roster.active : '',
      person.roster ? person.roster.status : '',
      person.services.join(', ')
    ];
  });

  writeTable_(SHEETS.MASTER, MASTER_HEADERS, rows);
  formatMasterTab_();
  return rows.length;
}

/**
 * Loads every provider's records once and joins them into a single in-memory
 * model. Both buildMasterTab and the rule engine run off this, so a full
 * refresh reads each tab once instead of once per consumer.
 */
function loadReconciliationModel() {
  var settings = getSettings();
  var disabledOu = new RegExp(settings.disabledOuPattern, 'i');

  var anchorRecords = settings.anchorSource === MICROSOFT_SOURCE
    ? readMicrosoftRecords()
    : readGoogleWorkspaceRecords();

  var microsoftByEmail = indexByEmail_(readMicrosoftRecords());
  var rosterByEmail = indexByEmail_(readRosterRecords());
  var services = getServiceAssignments();

  var people = anchorRecords.map(function (record) {
    return {
      email: record.email,
      fullName: record.fullName || '',
      orgPath: record.orgPath || '',
      // An account is effectively disabled if the directory says so OR if it
      // has been parked in a disabled-by-convention org unit. Organizations
      // use both mechanisms, often at the same time.
      anchorEnabled: record.enabled === true && !disabledOu.test(record.orgPath || ''),
      lastLoginDays: daysSince_(record.lastLogin),
      microsoft: microsoftByEmail[record.email] || null,
      roster: rosterByEmail[record.email] || null,
      services: services[record.email] || []
    };
  });

  return {
    settings: settings,
    people: people,
    peopleByEmail: indexByEmail_(people),
    microsoft: microsoftByEmail,
    roster: rosterByEmail,
    services: services
  };
}

function indexByEmail_(records) {
  var index = {};
  records.forEach(function (record) {
    if (record.email) {
      index[record.email] = record;
    }
  });
  return index;
}
