function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var advancedMenu = ui.createMenu('Advanced')
    .addItem('Setup Sheet', 'setupSheet')
    .addItem('Setup Triggers', 'setupTriggers')
    .addSeparator()
    .addItem('Run Maintenance', 'runDailySecurityMaintenance')
    .addItem('Test Slack', 'testSlackConnection');

  ui
    .createMenu('System Security')
    .addItem('Sync Requests', 'syncAccessRequestsToAccounts')
    .addItem('Generate Reviews', 'generateReviewChecks')
    .addItem('Sync Reviews', 'syncReviewDecisionsToAccounts')
    .addItem('Refresh Sheet', 'refreshSheet')
    .addSeparator()
    .addSubMenu(advancedMenu)
    .addToUi();
}

function syncAccessRequestsToAccounts() {
  ensureThreeTabWorkbookReady_();
  var requestRows = getRequestRows_();
  var processedRequestIds = {};
  var created = 0;
  var updated = 0;
  var revoked = 0;
  var denied = 0;
  var archived = 0;
  var skipped = 0;

  // Pre-build lookup — avoids re-reading the sheet for every request
  var requestRowMap = {};
  for (var r = 0; r < requestRows.length; r++) {
    var rid = normalizeText_(requestRows[r].data.REQUEST_ID);
    if (rid && !requestRowMap[rid]) {
      requestRowMap[rid] = requestRows[r];
    }
  }

  // Defer archives until after all syncs so row numbers stay stable
  var toArchive = [];
  var toDelete = [];

  for (var i = requestRows.length - 1; i >= 0; i--) {
    var requestId = normalizeText_(requestRows[i].data.REQUEST_ID);
    if (!requestId || processedRequestIds[requestId]) {
      continue;
    }
    processedRequestIds[requestId] = true;

    var liveRequestRow = requestRowMap[requestId];
    if (!liveRequestRow) {
      continue;
    }

    var result = syncAccessRequestRowToAccount_(liveRequestRow, true);
    if (!result) {
      continue;
    }
    if (result.action === 'created') {
      created++;
    } else if (result.action === 'updated') {
      updated++;
    } else if (result.action === 'revoked') {
      revoked++;
    } else if (result.action === 'denied') {
      denied++;
    }

    if (!requestSyncOutcomeReady_(liveRequestRow.data, result)) {
      skipped++;
      continue;
    }

    if (result.archiveRequest) {
      toArchive.push(requestId);
    } else {
      toDelete.push(requestId);
    }
  }

  // Archive/delete in one pass after all syncs are done
  for (var a = 0; a < toArchive.length; a++) {
    if (archiveActiveRequestById_(toArchive[a])) {
      archived++;
    }
  }
  for (var d = 0; d < toDelete.length; d++) {
    deleteActiveRequestById_(toDelete[d]);
  }

  SpreadsheetApp.flush();
  reconcileRequestSections_();
  dedupeAccountRowsByUniqueSignature_();
  sortRequestSection_();
  sortArchivedRequestSection_();
  sortAccountSection_();
  sortReviewSection_();

  sendHeartbeat('Request Sync', 'Created ' + created + ', updated ' + updated + ', revoked ' + revoked + ', denied ' + denied + ', archived ' + archived + ', skipped ' + skipped);

  var message = 'Request sync complete.\n\n' +
    'Created accounts: ' + created + '\n' +
    'Updated accounts: ' + updated + '\n' +
    'Revoked accounts: ' + revoked + '\n' +
    'Denied requests: ' + denied + '\n' +
    'Archived requests: ' + archived + '\n' +
    'Safety skips: ' + skipped;

  try {
    SpreadsheetApp.getUi().alert('Sync Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    Logger.log(message);
  }
}

function syncAccessRequestRowToAccount_(requestRow, skipInitAndSort) {
  var request = requestRow.data;
  var approval = normalizeText_(request.APPROVAL);
  var provisioning = normalizeSetupStatus_(request.PROVISIONING);
  var requestType = normalizeText_(request.REQUEST_TYPE);

  if (approval === 'Denied') {
    return finalizeDeniedRequest_(requestRow);
  }

  if (approval !== 'Approved') {
    return null;
  }

  if (requestType === 'New Access' || requestType === 'Change Access') {
    return upsertActiveAccountFromRequest_(requestRow, skipInitAndSort);
  }

  if (requestType === 'Remove Access') {
    request.PROVISIONING = 'Completed';
    return revokeActiveAccountFromRequest_(requestRow);
  }

  return null;
}

function finalizeDeniedRequest_(requestRow) {
  var request = requestRow.data;
  request.PROVISIONING = 'Closed';
  request.DECISION_DATE = request.DECISION_DATE || new Date();
  request.NEXT_SLACK_TRIGGER = 'Post denial update';
  updateRequestRow_(requestRow.rowNumber, request, requestRow.sheetName);

  return {
    action: 'denied',
    accessId: normalizeText_(request.ACCESS_ID),
    archiveRequest: true
  };
}

function upsertActiveAccountFromRequest_(requestRow, skipInitAndSort) {
  var request = requestRow.data;
  var canonicalSystem = canonicalSystemName_(request.GMP_SYSTEM);
  var accessId = normalizeIdentifier_(request.ACCESS_ID);
  var existingAccountRow = findBestAccountRowForRequest_(request);
  var provisioning = normalizeSetupStatus_(request.PROVISIONING);

  if (!provisioning || provisioning === 'Open') {
    provisioning = 'In Progress';
    request.PROVISIONING = provisioning;
  }

  if (!accessId) {
    accessId = existingAccountRow ? normalizeIdentifier_(existingAccountRow.data.ACCESS_ID) : nextStableId_('ACC');
  }

  var account = existingAccountRow ? existingAccountRow.data : buildEmptyRecord_(GMP_SCHEMA.accounts.keys);
  var privileged = isPrivilegedAccessLevel_(request.ACCESS_LEVEL) ? 'Yes' : 'No';
  var isProvisioned = provisioning === 'Completed' || provisioning === 'Closed';
  var requestType = normalizeText_(request.REQUEST_TYPE);
  var currentState = normalizePresenceState_(account.CURRENT_STATE || '');
  var targetState = isProvisioned ? 'Access Granted' : ((requestType === 'Change Access' && existingAccountRow) ? (currentState || 'Access Granted') : 'Setting Up');
  var grantedOn = isProvisioned ? (account.GRANTED_ON || request.DECISION_DATE || new Date()) : (account.GRANTED_ON || '');

  account.ACCESS_ID = accessId;
  account.SOURCE_REQUEST_ID = request.REQUEST_ID;
  account.GMP_SYSTEM = canonicalSystem;
  account.PERSON = request.TARGET_USER;
  account.COMPANY_EMAIL = request.COMPANY_EMAIL;
  account.DEPT = request.DEPT;
  account.PLATFORM_ACCOUNT = request.COMPANY_EMAIL;
  account.ACCESS_LEVEL = request.ACCESS_LEVEL;
  account.PRIVILEGED = privileged;
  account.CURRENT_STATE = targetState;
  account.LAST_VERIFIED_AT = isProvisioned ? new Date() : (account.LAST_VERIFIED_AT || '');
  account.GRANTED_ON = grantedOn;
  account.ACTIVITY_BAND = account.ACTIVITY_BAND || 'Unknown';
  account.EXPECTED_USE = account.EXPECTED_USE || defaultExpectedUse_(canonicalSystem);
  account.MIN_LOGINS_30D = account.MIN_LOGINS_30D || defaultMinLogins30d_(account.EXPECTED_USE);
  account.STALE_AFTER_DAYS = account.STALE_AFTER_DAYS || defaultStaleAfterDays_(canonicalSystem, account.EXPECTED_USE);
  account.ACTIVITY_SOURCE = account.ACTIVITY_SOURCE || 'Manual review';
  account.ACTIVITY_CONFIDENCE = account.ACTIVITY_CONFIDENCE || 'Low';
  account.REVIEW_TRIGGER = isProvisioned && privileged === 'Yes' ? 'Privileged Access' : account.REVIEW_TRIGGER;
  account.REVIEW_STATUS = isProvisioned ? (account.REVIEW_STATUS || '') : account.REVIEW_STATUS;
  account.NEXT_REVIEW_DUE = isProvisioned ? (account.NEXT_REVIEW_DUE || defaultNextReviewDue_(canonicalSystem, request.ACCESS_LEVEL)) : account.NEXT_REVIEW_DUE;
  account.MFA = account.MFA || '';
  account.OWNER = request.IT_OWNER;
  account.SLACK_THREAD = request.SLACK_THREAD || account.SLACK_THREAD;
  account.NOTES = appendNoteLine_(
    account.NOTES,
    formatDateOnly_(new Date()) + ': synced ' + request.REQUEST_ID + (isProvisioned ? '' : ' setup')
  );

  if (existingAccountRow) {
    updateAccountRow_(existingAccountRow.rowNumber, account);
  } else {
    appendAccountRecord_(account, skipInitAndSort);
    createInitialReviewForAccount_(account, request, skipInitAndSort);
  }

  request.ACCESS_ID = accessId;
  request.GMP_SYSTEM = canonicalSystem;
  request.DECISION_DATE = request.DECISION_DATE || new Date();
  request.NEXT_SLACK_TRIGGER = isProvisioned ? 'Post setup update' : 'Continue setup';
  updateRequestRow_(requestRow.rowNumber, request, requestRow.sheetName);

  return {
    action: existingAccountRow ? 'updated' : 'created',
    accessId: accessId,
    archiveRequest: false
  };
}

function createInitialReviewForAccount_(account, request, skipInitAndSort) {
  var review = buildEmptyRecord_(GMP_SCHEMA.reviews.keys);
  review.REVIEW_ID = nextStableId_('REV');
  review.REVIEW_CYCLE = reviewCycleLabel_(new Date());
  review.REVIEW_PRIORITY = isPrivilegedAccessLevel_(account.ACCESS_LEVEL) ? 'High' : 'Medium';
  review.PERSON = account.PERSON;
  review.SYSTEM = account.GMP_SYSTEM;
  review.ACCESS_LEVEL = account.ACCESS_LEVEL;
  review.PRESENCE_STATUS = account.CURRENT_STATE;
  review.WHY_FLAGGED = 'New account — initial review';
  review.DECISION_STATUS = 'Open';
  review.TRIGGER_TYPE = 'New Access';
  review.ACCESS_ID = account.ACCESS_ID;
  review.LINKED_REQUEST = normalizeText_(request.REQUEST_ID);
  review.NEXT_REVIEW_DUE = '';
  review.NEXT_SLACK_TRIGGER = 'Post review queue alert';
  review.COMPANY_EMAIL = account.COMPANY_EMAIL;
  appendReviewRecord_(review, skipInitAndSort);
}

function revokeActiveAccountFromRequest_(requestRow) {
  var request = requestRow.data;
  var canonicalSystem = canonicalSystemName_(request.GMP_SYSTEM);
  var accessRow = findBestAccountRowForRequest_(request);

  if (!accessRow) {
    return null;
  }

  var account = accessRow.data;
  account.CURRENT_STATE = 'Revoked';
  account.REVIEW_STATUS = 'Done';
  account.REVIEW_TRIGGER = 'Access Removed';
  account.NEXT_REVIEW_DUE = '';
  account.LAST_VERIFIED_AT = new Date();
  account.SLACK_THREAD = request.SLACK_THREAD || account.SLACK_THREAD;
  account.NOTES = appendNoteLine_(
    account.NOTES,
    formatDateOnly_(new Date()) + ': revoked ' + request.REQUEST_ID
  );
  updateAccountRow_(accessRow.rowNumber, account);

  // Close any open reviews for this account
  closeOpenReviewsForAccount_(account);

  request.ACCESS_ID = account.ACCESS_ID;
  request.GMP_SYSTEM = canonicalSystem;
  request.DECISION_DATE = request.DECISION_DATE || new Date();
  request.PROVISIONING = normalizeSetupStatus_(request.PROVISIONING);
  request.NEXT_SLACK_TRIGGER = 'Confirm removal in Slack';
  updateRequestRow_(requestRow.rowNumber, request, requestRow.sheetName);

  return {
    action: 'revoked',
    accessId: account.ACCESS_ID,
    archiveRequest: true
  };
}

function syncReviewDecisionsToAccounts() {
  ensureThreeTabWorkbookReady_();
  var reviewRows = getReviewRows_();
  var synced = 0;
  var reduced = 0;
  var removed = 0;

  for (var i = 0; i < reviewRows.length; i++) {
    var review = reviewRows[i].data;
    if (normalizeText_(review.DECISION_STATUS) !== 'Done') {
      continue;
    }

    var action = normalizeText_(review.REVIEWER_ACTION);
    if (!action) {
      continue;
    }

    var accountRow = findAccountRowForReview_(review);
    if (!accountRow) {
      continue;
    }

    var account = accountRow.data;
    var changed = false;

    if (action === 'Keep') {
      if (normalizeText_(account.REVIEW_STATUS) !== 'Done') {
        account.REVIEW_STATUS = 'Done';
        account.REVIEW_TRIGGER = '';
        account.LAST_REVIEW_DATE = review.DECISION_DATE || new Date();
        account.NEXT_REVIEW_DUE = account.NEXT_REVIEW_DUE || nextReviewDueFromDecision_(review.DECISION_DATE || new Date(), account.GMP_SYSTEM, account.ACCESS_LEVEL, action);
        changed = true;
      }
    } else if (action === 'Reduce') {
      var newLevel = normalizeText_(review.ACCESS_LEVEL);
      if (newLevel && newLevel !== normalizeText_(account.ACCESS_LEVEL)) {
        account.ACCESS_LEVEL = newLevel;
        account.NOTES = appendNoteLine_(account.NOTES, formatDateOnly_(new Date()) + ': access reduced to ' + newLevel + ' per review ' + review.REVIEW_ID);
        reduced++;
      }
      account.REVIEW_STATUS = 'Done';
      account.LAST_REVIEW_DATE = review.DECISION_DATE || new Date();
      account.NEXT_REVIEW_DUE = account.NEXT_REVIEW_DUE || nextReviewDueFromDecision_(review.DECISION_DATE || new Date(), account.GMP_SYSTEM, newLevel || account.ACCESS_LEVEL, action);
      changed = true;
    } else if (action === 'Remove') {
      if (normalizeText_(account.CURRENT_STATE) !== 'Revoked') {
        account.CURRENT_STATE = 'Revoked';
        account.REVIEW_STATUS = 'Done';
        account.REVIEW_TRIGGER = 'Removal Requested';
        account.LAST_REVIEW_DATE = review.DECISION_DATE || new Date();
        account.NOTES = appendNoteLine_(account.NOTES, formatDateOnly_(new Date()) + ': revoked per review ' + review.REVIEW_ID);
        removed++;
        changed = true;
      }
    }

    if (changed) {
      updateAccountRow_(accountRow.rowNumber, account);
      synced++;
    }

    // Ensure next check date is on the review row
    if (account.NEXT_REVIEW_DUE && !review.NEXT_REVIEW_DUE) {
      review.NEXT_REVIEW_DUE = account.NEXT_REVIEW_DUE;
      updateReviewRow_(reviewRows[i].rowNumber, review);
    }
  }

  sortAccountSection_();
  sortReviewSection_();

  var message = 'Review sync complete.\n\n' +
    'Accounts updated: ' + synced + '\n' +
    'Access reduced: ' + reduced + '\n' +
    'Accounts revoked: ' + removed;

  try {
    SpreadsheetApp.getUi().alert('Sync Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    Logger.log(message);
  }
}

function closeOpenReviewsForAccount_(account) {
  try {
    var reviewRows = getReviewRows_();
    var accountIdentity = accountReusableReviewIdentityKey_(account);
    for (var i = 0; i < reviewRows.length; i++) {
      var review = reviewRows[i].data;
      if (normalizeText_(review.DECISION_STATUS) === 'Done') {
        continue;
      }
      if (reviewReusableIdentityKey_(review) === accountIdentity) {
        review.DECISION_STATUS = 'Done';
        review.REVIEWER_ACTION = review.REVIEWER_ACTION || 'Remove';
        review.DECISION_DATE = new Date();
        review.NOTES = appendNoteLine_(review.NOTES, formatDateOnly_(new Date()) + ': auto-closed — account revoked');
        updateReviewRow_(reviewRows[i].rowNumber, review);
      }
    }
  } catch (err) {
    Logger.log('closeOpenReviewsForAccount_ error: ' + err.message);
  }
}

function requestSyncOutcomeReady_(request, result) {
  if (!request || !result) {
    return false;
  }

  if (result.archiveRequest) {
    if (result.action === 'denied') {
      return true;
    }

    if (result.action === 'revoked') {
      return !!findAccountRowByRequestOutcome_(request, result);
    }

    return true;
  }

  if (result.action === 'created' || result.action === 'updated') {
    return !!findAccountRowByRequestOutcome_(request, result);
  }

  return true;
}

function findAccountRowByRequestOutcome_(request, result) {
  if (result && normalizeIdentifier_(result.accessId)) {
    var directMatch = findAccountRowByAccessId_(result.accessId);
    if (directMatch) {
      return directMatch;
    }
  }

  var sourceRequestId = normalizeText_(request.REQUEST_ID);
  if (sourceRequestId) {
    var accountRows = getAccountRows_();
    for (var i = 0; i < accountRows.length; i++) {
      if (normalizeText_(accountRows[i].data.SOURCE_REQUEST_ID) === sourceRequestId) {
        return accountRows[i];
      }
    }
  }

  return findBestAccountRowForRequest_(request);
}

function buildEmptyRecord_(keys) {
  var record = {};
  for (var i = 0; i < keys.length; i++) {
    record[keys[i]] = '';
  }
  return record;
}

function isPrivilegedAccessLevel_(accessLevel) {
  var value = normalizeText_(accessLevel).toLowerCase();
  return value === 'admin' || value === 'full access' || value === 'read-write';
}

function defaultExpectedUse_(systemName) {
  var value = canonicalSystemName_(systemName).toLowerCase();
  if (value.indexOf('katana') !== -1) {
    return 'Daily';
  }
  if (value.indexOf('google drive') !== -1) {
    return 'Ad Hoc';
  }
  if (value.indexOf('shopify') !== -1) {
    return 'Weekly';
  }
  return 'Weekly';
}

function defaultMinLogins30d_(expectedUse) {
  var value = normalizeText_(expectedUse).toLowerCase();
  if (value === 'daily') {
    return 12;
  }
  if (value === 'weekly') {
    return 4;
  }
  if (value === 'ad hoc') {
    return 1;
  }
  return 2;
}

function defaultStaleAfterDays_(systemName, expectedUse) {
  var value = canonicalSystemName_(systemName).toLowerCase();
  if (value.indexOf('katana') !== -1) {
    return 14;
  }
  if (value.indexOf('shopify') !== -1) {
    return 30;
  }

  var useValue = normalizeText_(expectedUse).toLowerCase();
  if (useValue === 'daily') {
    return 14;
  }
  if (useValue === 'weekly') {
    return 30;
  }
  return 45;
}

function reviewIntervalDays_(systemName, accessLevel) {
  var system = canonicalSystemName_(systemName).toLowerCase();
  var level = normalizeText_(accessLevel).toLowerCase();
  var isAdmin = level === 'admin';
  var isFullAccess = level === 'full access';

  // HIGH risk: Katana, 1Password
  if (system.indexOf('katana') !== -1 || system.indexOf('1password') !== -1) {
    return isAdmin ? 30 : (isFullAccess ? 60 : 90);
  }

  // MEDIUM risk: Shopify, Wasp, Amazon
  if (system.indexOf('shopify') !== -1 || system.indexOf('wasp') !== -1 || system.indexOf('amazon') !== -1) {
    return isAdmin ? 60 : (isFullAccess ? 90 : 180);
  }

  // LOW risk: ShipStation, Other
  return isAdmin ? 90 : (isFullAccess ? 180 : 365);
}

function scheduledReviewTriggerLabel_(systemName, accessLevel) {
  return reviewIntervalDays_(systemName, accessLevel) + '-day review due';
}

function nextReviewDueFromDecision_(decisionDate, systemName, accessLevel, reviewerAction) {
  if (!(decisionDate instanceof Date) || isNaN(decisionDate)) {
    return '';
  }
  if (normalizeText_(reviewerAction) === 'Remove') {
    return addDays_(decisionDate, 14);
  }
  return addDays_(decisionDate, reviewIntervalDays_(systemName, accessLevel));
}

function defaultNextReviewDue_(systemName, accessLevel) {
  return addDays_(new Date(), reviewIntervalDays_(systemName, accessLevel));
}
