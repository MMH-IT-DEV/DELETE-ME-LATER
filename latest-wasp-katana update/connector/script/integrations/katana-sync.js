function refreshKatanaAccounts() {
  return refreshKatanaAccountsInternal_(true);
}

function refreshKatanaAccountsInternal_(runActivityRefresh) {
  ensureThreeTabWorkbookReady_();

  var users = fetchKatanaUsers_();
  var accountRows = getAccountRows_();
  var katanaAccountsByEmail = buildKatanaAccountMapByEmail_(accountRows);
  var seenEmails = {};
  var created = 0;
  var updated = 0;
  var activated = 0;
  var stale = 0;

  for (var i = 0; i < users.length; i++) {
    var user = users[i];
    var email = normalizeText_(user.email).toLowerCase();
    if (!email) {
      continue;
    }

    seenEmails[email] = true;
    var result = reconcileKatanaUser_(user, katanaAccountsByEmail[email]);
    if (!result) {
      continue;
    }

    if (result.action === 'created') {
      created++;
      katanaAccountsByEmail[email] = {
        data: result.account
      };
    } else if (result.action === 'updated') {
      updated++;
    }

    if (result.activated) {
      activated++;
    }
  }

  stale = markMissingKatanaAccounts_(accountRows, seenEmails);
  sortAccountSection_();

  if (runActivityRefresh !== false) {
    refreshActivitySignals();
  }

  sendHeartbeat('Katana Account Refresh', 'Created ' + created + ', updated ' + updated + ', activated ' + activated + ', stale ' + stale);
  return {
    created: created,
    updated: updated,
    activated: activated,
    stale: stale,
    totalUsers: users.length
  };
}

function fetchKatanaUsers_() {
  var apiKey = PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.KATANA_API_KEY_PROPERTY);
  if (!apiKey) {
    throw new Error('Missing KATANA_API_KEY script property.');
  }

  var response = UrlFetchApp.fetch('https://api.katanamrp.com/v1/users', {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Katana API returned HTTP ' + response.getResponseCode() + ': ' + response.getContentText());
  }

  var payload = JSON.parse(response.getContentText());
  if (!payload || !payload.data || !payload.data.length) {
    return [];
  }

  return payload.data;
}

function buildKatanaAccountMapByEmail_(accountRows) {
  var map = {};
  for (var i = 0; i < accountRows.length; i++) {
    var account = accountRows[i].data;
    if (accountSystemGroupLabel_(account.GMP_SYSTEM) !== 'Katana') {
      continue;
    }

    var email = normalizeText_(account.COMPANY_EMAIL || account.PLATFORM_ACCOUNT).toLowerCase();
    if (!email || map[email]) {
      continue;
    }
    map[email] = accountRows[i];
  }
  return map;
}

function reconcileKatanaUser_(user, existingAccountRow) {
  var now = new Date();
  var email = normalizeText_(user.email);
  if (!email) {
    return null;
  }

  var displayName = katanaUserDisplayName_(user);
  var verificationNote = formatDateOnly_(now) + ': verified in Katana API (' + user.id + ')';

  if (existingAccountRow && existingAccountRow.data) {
    var account = existingAccountRow.data;
    var wasProvisioning = normalizePresenceState_(account.CURRENT_STATE) === 'Setting Up';

    account.GMP_SYSTEM = account.GMP_SYSTEM || 'Katana';
    account.PERSON = displayName || account.PERSON;
    account.COMPANY_EMAIL = email;
    account.PLATFORM_ACCOUNT = email;
    account.EXPECTED_USE = account.EXPECTED_USE || defaultExpectedUse_('Katana');
    account.MIN_LOGINS_30D = account.MIN_LOGINS_30D || defaultMinLogins30d_(account.EXPECTED_USE);
    account.STALE_AFTER_DAYS = account.STALE_AFTER_DAYS || defaultStaleAfterDays_('Katana', account.EXPECTED_USE);
    account.ACTIVITY_SOURCE = 'Katana Users API';
    account.ACTIVITY_CONFIDENCE = 'Medium';
    account.LAST_VERIFIED_AT = now;
    if (normalizePresenceState_(account.CURRENT_STATE) !== 'Revoked') {
      account.CURRENT_STATE = normalizeText_(account.SOURCE_REQUEST_ID) ? 'Access Granted' : 'Unmanaged';
    }
    account.NOTES = appendUniqueNoteLine_(account.NOTES, verificationNote);

    if (wasProvisioning) {
      account.GRANTED_ON = account.GRANTED_ON || now;
      account.NEXT_REVIEW_DUE = account.NEXT_REVIEW_DUE || defaultNextReviewDue_(account.GMP_SYSTEM, account.ACCESS_LEVEL);
    }

    updateAccountRow_(existingAccountRow.rowNumber, account);

    if (normalizeText_(account.SOURCE_REQUEST_ID)) {
      syncKatanaProvisioningRequest_(account.SOURCE_REQUEST_ID, user);
    }

    return {
      action: 'updated',
      activated: wasProvisioning,
      account: account
    };
  }

  var discoveredAccount = buildEmptyRecord_(GMP_SCHEMA.accounts.keys);
  discoveredAccount.ACCESS_ID = nextStableId_('ACC');
  discoveredAccount.SOURCE_REQUEST_ID = '';
  discoveredAccount.GMP_SYSTEM = 'Katana';
  discoveredAccount.PERSON = displayName;
  discoveredAccount.COMPANY_EMAIL = email;
  discoveredAccount.DEPT = '';
  discoveredAccount.PLATFORM_ACCOUNT = email;
  discoveredAccount.ACCESS_LEVEL = '';
  discoveredAccount.CURRENT_STATE = 'Unmanaged';
  discoveredAccount.LAST_VERIFIED_AT = now;
  discoveredAccount.ACTIVITY_BAND = 'Unknown';
  discoveredAccount.GRANTED_ON = '';
  discoveredAccount.LAST_LOGIN = '';
  discoveredAccount.DAYS_SINCE_LOGIN = '';
  discoveredAccount.LOGINS_30D = '';
  discoveredAccount.LOGINS_90D = '';
  discoveredAccount.NEXT_REVIEW_DUE = now;
  discoveredAccount.LAST_REVIEW_DATE = '';
  discoveredAccount.REVIEW_STATUS = '';
  discoveredAccount.OWNER = '';
  discoveredAccount.MFA = '';
  discoveredAccount.PRIVILEGED = 'No';
  discoveredAccount.ACTIVITY_SCORE = '';
  discoveredAccount.REVIEW_TRIGGER = '';
  discoveredAccount.NOTES = formatDateOnly_(now) + ': Katana no request (' + user.id + ')';
  discoveredAccount.EXPECTED_USE = defaultExpectedUse_('Katana');
  discoveredAccount.MIN_LOGINS_30D = defaultMinLogins30d_(discoveredAccount.EXPECTED_USE);
  discoveredAccount.STALE_AFTER_DAYS = defaultStaleAfterDays_('Katana', discoveredAccount.EXPECTED_USE);
  discoveredAccount.ACTIVITY_SOURCE = 'Katana Users API';
  discoveredAccount.ACTIVITY_CONFIDENCE = 'Medium';
  discoveredAccount.SLACK_THREAD = '';

  appendAccountRecord_(discoveredAccount);

  return {
    action: 'created',
    activated: false,
    account: discoveredAccount
  };
}

function syncKatanaProvisioningRequest_(requestId, user) {
  var requestRow = findRequestRowById_(requestId);
  if (!requestRow) {
    return;
  }

  var request = requestRow.data;
  var provisioning = normalizeSetupStatus_(request.PROVISIONING);
  if (provisioning === 'Completed' || provisioning === 'Closed') {
    return;
  }

  request.PROVISIONING = 'Completed';
  request.DECISION_DATE = request.DECISION_DATE || new Date();
  request.NEXT_SLACK_TRIGGER = 'Post setup update';
  request.NOTES = appendUniqueNoteLine_(
    request.NOTES,
    formatDateOnly_(new Date()) + ': Katana verified ' + normalizeText_(user.email) + ' (' + user.id + ')'
  );
  updateRequestRow_(requestRow.rowNumber, request, requestRow.sheetName);
  sortRequestSection_();
}

function katanaUserDisplayName_(user) {
  var first = normalizeText_(user.firstName);
  var last = normalizeText_(user.lastName);
  var fullName = normalizeText_((first + ' ' + last).trim());
  if (fullName) {
    return fullName;
  }

  var email = normalizeText_(user.email);
  if (!email) {
    return '';
  }
  return email.split('@')[0];
}

function appendUniqueNoteLine_(existingValue, noteLine) {
  var current = normalizeText_(existingValue);
  if (current.indexOf(noteLine) !== -1) {
    return current;
  }
  return appendNoteLine_(current, noteLine);
}

function markMissingKatanaAccounts_(accountRows, seenEmails) {
  var now = new Date();
  var staleCount = 0;

  for (var i = 0; i < accountRows.length; i++) {
    var account = accountRows[i].data;
    if (accountSystemGroupLabel_(account.GMP_SYSTEM) !== 'Katana') {
      continue;
    }
    if (normalizeText_(account.CURRENT_STATE) === 'Revoked') {
      continue;
    }

    var email = normalizeText_(account.COMPANY_EMAIL || account.PLATFORM_ACCOUNT).toLowerCase();
    if (!email || seenEmails[email]) {
      continue;
    }

    account.CURRENT_STATE = 'Missing';
    account.NEXT_REVIEW_DUE = now;
    account.ACTIVITY_SOURCE = 'Katana Users API';
    account.ACTIVITY_CONFIDENCE = 'Medium';
    account.NOTES = appendUniqueNoteLine_(account.NOTES, formatDateOnly_(now) + ': Katana missing');
    updateAccountRow_(accountRows[i].rowNumber, account);
    staleCount++;
  }

  return staleCount;
}
