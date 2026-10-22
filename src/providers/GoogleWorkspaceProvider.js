/**
 * Google Workspace provider, via the Admin SDK Directory API.
 *
 * Requires the "Admin Directory API" advanced service to be enabled
 * (Services > Add a service > Admin SDK API), and the executing account must
 * hold a role with User Read privileges. No API key or client secret is
 * involved - Apps Script authorizes as the user running the script.
 *
 * Paging is explicit. Directory.Users.list caps out at 500 users per page and
 * returns a nextPageToken; a single unpaged call silently truncates, which is
 * the kind of bug that produces a clean-looking report that is simply missing
 * half the organization.
 */

var GOOGLE_SOURCE = 'GoogleWorkspace';

var GOOGLE_HEADERS = [
  'email', 'enabled', 'orgPath', 'lastLogin', 'licenses', 'externalId', 'fullName'
];

function syncGoogleWorkspace() {
  var users = [];
  var pageToken = null;

  do {
    var page = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      orderBy: 'email',
      projection: 'full',
      pageToken: pageToken
    });
    (page.users || []).forEach(function (user) {
      users.push(toGoogleRecord_(user));
    });
    pageToken = page.nextPageToken;
  } while (pageToken);

  writeTable_(GOOGLE_SOURCE, GOOGLE_HEADERS, recordsToRows_(users, GOOGLE_HEADERS));
  return users.length;
}

function toGoogleRecord_(user) {
  return {
    email: normalizeEmail_(user.primaryEmail),
    // `suspended` is the explicit signal; archived accounts are disabled too
    // but report suspended=false, so both are folded in here.
    enabled: !(user.suspended === true || user.archived === true),
    orgPath: user.orgUnitPath || '',
    lastLogin: user.lastLoginTime || '',
    // A workspace license is not exposed on the user object; callers who need
    // it should enable the Licensing API. Left blank rather than guessed.
    licenses: '',
    externalId: extractExternalId_(user),
    fullName: (user.name && user.name.fullName) || ''
  };
}

/**
 * Pulls an employee/student id out of whichever field the organization
 * actually populates. Directories vary wildly here, so this checks the
 * standard externalIds array first and falls back to nothing rather than
 * inventing a key - a wrong join key is worse than no join key.
 */
function extractExternalId_(user) {
  var ids = user.externalIds || [];
  for (var i = 0; i < ids.length; i++) {
    if (ids[i] && ids[i].value) {
      return String(ids[i].value);
    }
  }
  return '';
}

function readGoogleWorkspaceRecords() {
  return readTable_(GOOGLE_SOURCE).map(function (row) {
    return {
      email: normalizeEmail_(row.email),
      enabled: parseBoolean_(row.enabled),
      orgPath: String(row.orgPath || ''),
      lastLogin: row.lastLogin || '',
      licenses: String(row.licenses || ''),
      externalId: String(row.externalId || ''),
      fullName: String(row.fullName || ''),
      source: GOOGLE_SOURCE
    };
  }).filter(function (record) {
    return !!record.email;
  });
}
