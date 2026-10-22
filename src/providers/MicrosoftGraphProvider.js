/**
 * Microsoft Entra ID (Azure AD) provider, via Microsoft Graph.
 *
 * Authenticates with the OAuth2 client-credentials flow - an app registration
 * with the application permission User.Read.All, admin-consented. That makes
 * this provider independent of whoever happens to be running the script,
 * which matters for a time-driven trigger.
 *
 * Register the app in Entra ID, then set three Script Properties:
 *   MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET
 *
 * If those are unset the provider reports itself unconfigured and the rest of
 * the run proceeds without it. A partially configured deployment should still
 * produce a useful report over the sources it does have.
 */

var MICROSOFT_SOURCE = 'Microsoft';

var MICROSOFT_HEADERS = ['email', 'enabled', 'licenses', 'immutableId', 'fullName'];

var GRAPH_USER_FIELDS = [
  'userPrincipalName',
  'mail',
  'accountEnabled',
  'assignedLicenses',
  'onPremisesImmutableId',
  'displayName'
].join(',');

function isMicrosoftConfigured() {
  return !!(
    getSecretOptional_(SECRET_KEYS.MS_TENANT_ID) &&
    getSecretOptional_(SECRET_KEYS.MS_CLIENT_ID) &&
    getSecretOptional_(SECRET_KEYS.MS_CLIENT_SECRET)
  );
}

function getGraphToken_() {
  var tenant = getSecret_(SECRET_KEYS.MS_TENANT_ID);
  var response = UrlFetchApp.fetch(
    'https://login.microsoftonline.com/' + encodeURIComponent(tenant) + '/oauth2/v2.0/token',
    {
      method: 'post',
      muteHttpExceptions: true,
      payload: {
        client_id: getSecret_(SECRET_KEYS.MS_CLIENT_ID),
        client_secret: getSecret_(SECRET_KEYS.MS_CLIENT_SECRET),
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      }
    }
  );

  var body = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200 || !body.access_token) {
    // Surface Entra's own error description - "invalid_client" versus
    // "unauthorized_client" versus a consent problem are three very different
    // fixes, and collapsing them into "auth failed" wastes an afternoon.
    throw new Error(
      'Graph token request failed (' + response.getResponseCode() + '): ' +
      (body.error_description || response.getContentText())
    );
  }
  return body.access_token;
}

function syncMicrosoft() {
  if (!isMicrosoftConfigured()) {
    Logger.log('Microsoft sync skipped - MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET not set.');
    return 0;
  }

  var token = getGraphToken_();
  var url = 'https://graph.microsoft.com/v1.0/users?$top=999&$select=' + GRAPH_USER_FIELDS;
  var records = [];

  // Graph pages with an opaque @odata.nextLink. Follow it verbatim rather
  // than rebuilding the query - the link carries a skiptoken that cannot be
  // reconstructed from the original parameters.
  while (url) {
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Graph /users failed (' + response.getResponseCode() + '): ' + response.getContentText());
    }
    var page = JSON.parse(response.getContentText());
    (page.value || []).forEach(function (user) {
      var record = toMicrosoftRecord_(user);
      if (record.email) {
        records.push(record);
      }
    });
    url = page['@odata.nextLink'] || null;
  }

  writeTable_(MICROSOFT_SOURCE, MICROSOFT_HEADERS, recordsToRows_(records, MICROSOFT_HEADERS));
  return records.length;
}

function toMicrosoftRecord_(user) {
  return {
    // Prefer `mail` over userPrincipalName: the UPN is a login identifier and
    // is not always a working address, which makes it a poor join key when
    // the two diverge.
    email: normalizeEmail_(user.mail || user.userPrincipalName),
    enabled: user.accountEnabled === true,
    licenses: (user.assignedLicenses || []).map(function (l) { return l.skuId; }).join(', '),
    immutableId: user.onPremisesImmutableId || '',
    fullName: user.displayName || ''
  };
}

function readMicrosoftRecords() {
  return readTable_(MICROSOFT_SOURCE).map(function (row) {
    return {
      email: normalizeEmail_(row.email),
      enabled: parseBoolean_(row.enabled),
      licenses: String(row.licenses || ''),
      immutableId: String(row.immutableId || ''),
      fullName: String(row.fullName || ''),
      source: MICROSOFT_SOURCE
    };
  }).filter(function (record) {
    return !!record.email;
  });
}
