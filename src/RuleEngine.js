/**
 * Findings: concrete, named rules rather than a generic diff.
 *
 * A diff-everything pass over four systems produces thousands of
 * discrepancies, nearly all of them meaningless, and gets ignored within a
 * week. Each rule below describes a specific situation somebody should
 * actually act on, and says so in plain language.
 *
 * Severity is advisory: "high" means a live access-control gap (someone can
 * still log in somewhere they should not), "medium" means wasted spend or
 * drift, "low" means worth a look.
 *
 * Rules degrade quietly. If Microsoft is unconfigured, every Microsoft rule
 * sees an empty index and contributes nothing, rather than reporting the
 * entire organization as missing.
 */

var FINDINGS_HEADERS = ['severity', 'email', 'rule', 'detail'];

function buildFindingsTab() {
  var model = loadReconciliationModel();
  var findings = [];

  findStaleOffboarding_(model, findings);
  findOrphanedAccounts_(model, findings);
  findRosterContradictions_(model, findings);
  findLicenseGaps_(model, findings);
  findDormantAccounts_(model, findings);

  var order = { high: 0, medium: 1, low: 2 };
  findings.sort(function (a, b) {
    var bySeverity = order[a[0]] - order[b[0]];
    return bySeverity !== 0 ? bySeverity : String(a[1]).localeCompare(String(b[1]));
  });

  writeTable_(SHEETS.FINDINGS, FINDINGS_HEADERS, findings);
  formatFindingsTab_();
  return findings.length;
}

/**
 * The offboarding that stopped halfway: the primary account is disabled but
 * the person can still sign in somewhere else. This is the rule that earns
 * the tool its keep.
 */
function findStaleOffboarding_(model, findings) {
  model.people.forEach(function (person) {
    if (person.anchorEnabled) {
      return;
    }
    if (person.microsoft && person.microsoft.enabled === true) {
      findings.push([
        'high',
        person.email,
        'Disabled in primary directory but still enabled in Microsoft',
        'orgPath=' + (person.orgPath || '(none)')
      ]);
    }
    if (person.services.length) {
      findings.push([
        'high',
        person.email,
        'Disabled in primary directory but still assigned to services',
        person.services.join(', ')
      ]);
    }
  });
}

/**
 * An enabled account in a secondary system with no matching active person.
 * Already-disabled orphans are skipped - they are history, not a live gap,
 * and reporting them trains people to ignore the report.
 */
function findOrphanedAccounts_(model, findings) {
  Object.keys(model.microsoft).forEach(function (email) {
    var record = model.microsoft[email];
    var person = model.peopleByEmail[email];
    if (record.enabled === true && (!person || !person.anchorEnabled)) {
      findings.push([
        'high',
        email,
        'Microsoft account enabled with no active person record',
        person ? 'Person exists but is disabled in the primary directory' : 'No matching person record'
      ]);
    }
  });

  Object.keys(model.services).forEach(function (email) {
    var person = model.peopleByEmail[email];
    if (!person || !person.anchorEnabled) {
      findings.push([
        'medium',
        email,
        'Service assignment with no active person record',
        model.services[email].join(', ')
      ]);
    }
  });
}

/** The roster says this person is gone, but their accounts disagree. */
function findRosterContradictions_(model, findings) {
  Object.keys(model.roster).forEach(function (email) {
    var roster = model.roster[email];
    if (roster.active !== false) {
      return;
    }
    var person = model.peopleByEmail[email];
    var microsoft = model.microsoft[email];
    var services = model.services[email] || [];

    var stillActive = (person && person.anchorEnabled) ||
      (microsoft && microsoft.enabled === true) ||
      services.length > 0;

    if (stillActive) {
      findings.push([
        'high',
        email,
        'Roster shows inactive but accounts remain active',
        'rosterStatus=' + roster.status
      ]);
    }
  });
}

/**
 * Entitlement drift, driven entirely by a LicenseRules tab the administrator
 * fills in - columns: orgPathPrefix, expectedLicense.
 *
 * Deliberately empty by default. Which org unit implies which license is a
 * policy decision specific to each organization, and a guessed default would
 * generate a page of confident, wrong findings on first run.
 */
function findLicenseGaps_(model, findings) {
  var rules = readTable_(SHEETS.RULES).map(function (row) {
    return {
      orgPathPrefix: String(row.orgPathPrefix || ''),
      expectedLicense: String(row.expectedLicense || '')
    };
  }).filter(function (rule) {
    return rule.orgPathPrefix && rule.expectedLicense;
  });

  if (!rules.length) {
    return;
  }

  model.people.forEach(function (person) {
    if (!person.anchorEnabled) {
      return;
    }
    rules.forEach(function (rule) {
      if ((person.orgPath || '').indexOf(rule.orgPathPrefix) !== 0) {
        return;
      }
      var licenses = person.microsoft ? String(person.microsoft.licenses || '') : '';
      if (licenses.indexOf(rule.expectedLicense) === -1) {
        findings.push([
          'medium',
          person.email,
          'Org unit implies license "' + rule.expectedLicense + '" but it is not assigned',
          person.microsoft ? 'Microsoft account found, license missing' : 'No Microsoft account found'
        ]);
      }
    });
  });
}

/** Enabled accounts nobody has signed into in a long time. */
function findDormantAccounts_(model, findings) {
  var threshold = parseInt(model.settings.dormantAfterDays, 10);
  if (!threshold || threshold <= 0) {
    return;
  }

  model.people.forEach(function (person) {
    if (!person.anchorEnabled || person.lastLoginDays === null) {
      return;
    }
    if (person.lastLoginDays >= threshold) {
      findings.push([
        'low',
        person.email,
        'Enabled account dormant for ' + person.lastLoginDays + ' days',
        'Threshold is ' + threshold + ' days'
      ]);
    }
  });
}
