var AUTOMATION_MARKERS = [
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().requestHeaderRow; },
    column: function() { return GMP_SCHEMA.requests.columns.APPROVAL; },
    emoji: '⚡',
    note: '⚡ Approval logic\nThis field contributes to the request state.\nIt only sends Slack directly when the request is denied.\nApproved requests wait for the Access Setup Status column to send one consolidated Slack update.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().requestHeaderRow; },
    column: function() { return GMP_SCHEMA.requests.columns.PROVISIONING; },
    emoji: '🔔',
    note: '🔔 Request Slack update\nThis is the single request-status notification column.\nWhen it changes, Slack gets one consolidated thread update with approval, setup status, IT owner, and linked access information.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().requestHeaderRow; },
    column: function() { return GMP_SCHEMA.requests.columns.NEXT_SLACK_TRIGGER; },
    emoji: '⚡',
    note: '⚡ Slack follow-up queue\nThis field is automation-owned guidance for the next Slack action or reminder.\nIt does not send a Slack message by itself.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().requestHeaderRow; },
    column: function() { return GMP_SCHEMA.requests.columns.SLACK_THREAD; },
    emoji: '🧵',
    note: '🧵 Slack thread link\nThe slash-command intake stores the Slack thread permalink here so future status updates can post back to the same thread.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().accountHeaderRow; },
    column: function() { return GMP_SCHEMA.accounts.columns.LAST_LOGIN; },
    emoji: '🔄',
    note: '🔄 Activity refresh\nThis field is refreshed by account activity sync, such as Shopify login evidence.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().accountHeaderRow; },
    column: function() { return GMP_SCHEMA.accounts.columns.DAYS_SINCE_LOGIN; },
    emoji: '📊',
    note: '📊 Calculated metric\nThis value is recalculated from Last Login during activity refresh.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().accountHeaderRow; },
    column: function() { return GMP_SCHEMA.accounts.columns.ACTIVITY_SCORE; },
    emoji: '📊',
    note: '📊 Activity score\nDaily maintenance recalculates this score from login freshness, usage, MFA, and review posture.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().accountHeaderRow; },
    column: function() { return GMP_SCHEMA.accounts.columns.ACTIVITY_BAND; },
    emoji: '📊',
    note: '📊 Activity band\nDerived from the Activity Score and used to drive the Review Checks queue.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().accountHeaderRow; },
    column: function() { return GMP_SCHEMA.accounts.columns.REVIEW_TRIGGER; },
    emoji: '⏰',
    note: '⏰ Review trigger\nDaily maintenance sets this when an account becomes due, stale, low usage, or privileged.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().accountHeaderRow; },
    column: function() { return GMP_SCHEMA.accounts.columns.REVIEW_STATUS; },
    emoji: '⚡',
    note: '⚡ Review sync\nThis value is updated from decisions made in the Review Checks tab.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().accountHeaderRow; },
    column: function() { return GMP_SCHEMA.accounts.columns.NEXT_REVIEW_DUE; },
    emoji: '⏰',
    note: '⏰ Review due date\nDaily maintenance reads this date to generate or refresh Review Checks.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.ACCESS_CONTROL,
    row: function() { return getAccessControlContext_().accountHeaderRow; },
    column: function() { return GMP_SCHEMA.accounts.columns.SLACK_THREAD; },
    emoji: '🧵',
    note: '🧵 Slack thread link\nUsed to send access and review updates back into the original Slack conversation.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.REVIEW_CHECKS,
    row: function() { return GMP_SCHEMA.layout.REVIEW_HEADER_ROW; },
    column: function() { return GMP_SCHEMA.reviews.columns.REVIEWER_ACTION; },
    emoji: '⚡',
    note: '⚡ Review action logic\nThis value is read when Decision Status changes.\nIt does not send a Slack message by itself.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.REVIEW_CHECKS,
    row: function() { return GMP_SCHEMA.layout.REVIEW_HEADER_ROW; },
    column: function() { return GMP_SCHEMA.reviews.columns.DECISION_STATUS; },
    emoji: '🔔',
    note: '🔔 Review Slack update\nThis is the single review notification column.\nChanging it can update the linked account and send one consolidated thread update in Slack.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.REVIEW_CHECKS,
    row: function() { return GMP_SCHEMA.layout.REVIEW_HEADER_ROW; },
    column: function() { return GMP_SCHEMA.reviews.columns.NEXT_SLACK_TRIGGER; },
    emoji: '⚡',
    note: '⚡ Slack follow-up queue\nAutomation uses this field to show the next Slack follow-up that should happen for this review.\nIt does not send a Slack message by itself.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.REVIEW_CHECKS,
    row: function() { return GMP_SCHEMA.layout.REVIEW_HEADER_ROW; },
    column: function() { return GMP_SCHEMA.reviews.columns.SLACK_THREAD; },
    emoji: '🧵',
    note: '🧵 Slack thread link\nUsed for posting review decisions back into the same conversation thread.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.INCIDENTS,
    row: function() { return GMP_SCHEMA.layout.INCIDENT_HEADER_ROW; },
    column: function() { return GMP_SCHEMA.incidents.columns.STATUS; },
    emoji: '🔔',
    note: '🔔 Incident Slack update\nThis is the single incident notification column.\nChanging it posts one consolidated update back to the linked Slack thread.\nWhen Status = Resolved, Resolved At is stamped automatically.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.INCIDENTS,
    row: function() { return GMP_SCHEMA.layout.INCIDENT_HEADER_ROW; },
    column: function() { return GMP_SCHEMA.incidents.columns.NEXT_SLACK_TRIGGER; },
    emoji: '⚡',
    note: '⚡ Slack follow-up queue\nAutomation uses this field to show the next Slack action for the incident.\nIt does not send a Slack message by itself.'
  },
  {
    sheetName: GMP_SCHEMA.tabs.INCIDENTS,
    row: function() { return GMP_SCHEMA.layout.INCIDENT_HEADER_ROW; },
    column: function() { return GMP_SCHEMA.incidents.columns.SLACK_THREAD; },
    emoji: '🧵',
    note: '🧵 Slack thread link\nThe slash-command intake stores the incident thread permalink here so follow-up messages stay in one thread.'
  }
];

function setupTriggers() {
  removeAllTriggers();

  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  ScriptApp.newTrigger('runDailySecurityMaintenance')
    .timeBased()
    .atHour(GMP_CONFIG.DAILY_TRIGGER_HOUR)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('keepSecurityCommandWarm_')
    .timeBased()
    .everyMinutes(5)
    .create();

  ScriptApp.newTrigger('processPendingSecurityIntake_')
    .timeBased()
    .everyMinutes(1)
    .create();

  var markerResult = tryMarkAutomationColumns_();

  try {
    SpreadsheetApp.getUi().alert(
      'Triggers Active',
      'Installed:\n- onEdit workflow updates\n- daily maintenance\n- 5 minute keepalive\n- 1 minute intake queue processor\n' + markerResult,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    Logger.log('Triggers installed');
  }
}

function removeAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  clearAutomationColumnMarkers();
}

function keepSecurityCommandWarm_() {
  try {
    var url = ScriptApp.getService().getUrl();
    if (url) {
      UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    }
  } catch (err) {
    Logger.log('Warm ping skipped: ' + err.message);
  }
}

function sendHeartbeat(trigger, details) {
  try {
    var url = PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.HEALTH_MONITOR_URL_PROPERTY);
    if (!url) {
      return;
    }

    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        system: '2026 Security & GMP Connection Tracker',
        status: 'success',
        trigger: trigger || 'Security Tracker',
        details: details || ''
      }),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log('Heartbeat error: ' + err);
  }
}

function getSlackBotToken_() {
  var token = PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.BOT_TOKEN_PROPERTY);
  if (!token) {
    throw new Error('Missing SLACK_BOT_TOKEN script property.');
  }
  return token;
}

function hasSlackBotToken_() {
  return !!PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.BOT_TOKEN_PROPERTY);
}

function getConfiguredAccessReviewChannel_() {
  return PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.ACCESS_REVIEW_CHANNEL_PROPERTY) ||
    PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.SUPPORT_CHANNEL_PROPERTY) ||
    GMP_CONFIG.DEFAULT_SUPPORT_CHANNEL ||
    GMP_CONFIG.DEFAULT_ACCESS_REVIEW_CHANNEL;
}

function getConfiguredSupportChannel_() {
  return PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.SUPPORT_CHANNEL_PROPERTY) ||
    PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.ACCESS_REVIEW_CHANNEL_PROPERTY) ||
    GMP_CONFIG.DEFAULT_SUPPORT_CHANNEL ||
    GMP_CONFIG.DEFAULT_ACCESS_REVIEW_CHANNEL;
}

function getConfiguredAccessReviewChannels_() {
  return uniqueNonBlankValues_([
    PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.ACCESS_REVIEW_CHANNEL_PROPERTY),
    PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.SUPPORT_CHANNEL_PROPERTY),
    GMP_CONFIG.DEFAULT_SUPPORT_CHANNEL,
    GMP_CONFIG.DEFAULT_ACCESS_REVIEW_CHANNEL
  ]);
}

function getConfiguredSupportChannels_() {
  return uniqueNonBlankValues_([
    PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.SUPPORT_CHANNEL_PROPERTY),
    PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.ACCESS_REVIEW_CHANNEL_PROPERTY),
    GMP_CONFIG.DEFAULT_SUPPORT_CHANNEL,
    GMP_CONFIG.DEFAULT_ACCESS_REVIEW_CHANNEL
  ]);
}

function postSlackMessage_(channel, text, blocks, threadTs) {
  var resolvedChannel = resolveSlackChannelRef_(channel);
  var payload = {
    channel: resolvedChannel || channel,
    text: text,
    unfurl_links: false,
    unfurl_media: false
  };

  if (blocks && blocks.length) {
    payload.blocks = blocks;
  }
  if (threadTs) {
    payload.thread_ts = threadTs;
  }

  var data = callSlackApiJson_('chat.postMessage', payload);
  if (!data.ok && data.error === 'not_in_channel') {
    var joinedChannel = ensureSlackChannelMembership_(resolvedChannel || channel);
    if (joinedChannel) {
      payload.channel = joinedChannel;
      data = callSlackApiJson_('chat.postMessage', payload);
    }
  }

  if (!data.ok) {
    throw new Error('Slack chat.postMessage failed: ' + (data.error || JSON.stringify(data)));
  }
  return data;
}

function postSlackMessageWithFallback_(channels, text, blocks, threadTs) {
  var candidates = uniqueNonBlankValues_(channels);
  var failures = [];

  for (var i = 0; i < candidates.length; i++) {
    try {
      return postSlackMessage_(candidates[i], text, blocks, threadTs);
    } catch (err) {
      failures.push(candidates[i] + ': ' + err.message);
      if (!isSlackChannelRoutingError_(err)) {
        throw err;
      }
    }
  }

  throw new Error(failures.join(' | ') || 'No Slack channels available.');
}

function callSlackApiJson_(methodName, payload, requestMethod) {
  var options = {
    method: requestMethod || 'post',
    headers: { Authorization: 'Bearer ' + getSlackBotToken_() },
    muteHttpExceptions: true
  };

  if ((requestMethod || 'post').toLowerCase() === 'get') {
    options.contentType = 'application/x-www-form-urlencoded';
  } else {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload || {});
  }

  var response = UrlFetchApp.fetch('https://slack.com/api/' + methodName, options);
  return JSON.parse(response.getContentText());
}

function uniqueNonBlankValues_(values) {
  var unique = [];
  var seen = {};
  for (var i = 0; i < values.length; i++) {
    var value = normalizeText_(values[i]);
    if (!value || seen[value]) {
      continue;
    }
    seen[value] = true;
    unique.push(value);
  }
  return unique;
}

function isSlackChannelRoutingError_(err) {
  var message = normalizeText_(err && err.message).toLowerCase();
  return message.indexOf('channel_not_found') !== -1 || message.indexOf('not_in_channel') !== -1;
}

function resolveSlackChannelRef_(channelRef) {
  var value = normalizeText_(channelRef);
  if (!value) {
    return '';
  }

  if (/^[CDG][A-Z0-9]+$/.test(value)) {
    return value;
  }

  var channelName = value.charAt(0) === '#' ? value.slice(1) : value;
  if (!channelName) {
    return value;
  }

  try {
    return lookupSlackChannelIdByName_(channelName) || value;
  } catch (err) {
    Logger.log('Slack channel lookup failed for ' + value + ': ' + err.message);
    return value;
  }
}

function ensureSlackChannelMembership_(channelRef) {
  var channelId = resolveSlackChannelRef_(channelRef);
  if (!/^[CG][A-Z0-9]+$/.test(channelId)) {
    return '';
  }

  var response = callSlackApiJson_('conversations.join', { channel: channelId });
  if (response.ok || response.error === 'already_in_channel') {
    return channelId;
  }

  Logger.log('Slack join failed for ' + channelId + ': ' + (response.error || 'unknown_error'));
  return '';
}

function lookupSlackChannelIdByName_(channelName) {
  var cursor = '';
  var normalizedNeedle = normalizeText_(channelName).replace(/^#/, '').toLowerCase();

  while (true) {
    var url = 'https://slack.com/api/conversations.list?exclude_archived=true&limit=1000&types=public_channel,private_channel';
    if (cursor) {
      url += '&cursor=' + encodeURIComponent(cursor);
    }

    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + getSlackBotToken_() },
      muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    if (!data.ok) {
      throw new Error('Slack conversations.list failed: ' + (data.error || response.getContentText()));
    }

    var channels = data.channels || [];
    for (var i = 0; i < channels.length; i++) {
      var candidate = channels[i];
      var candidateName = normalizeText_(candidate.name).toLowerCase();
      var candidateNormalized = normalizeText_(candidate.name_normalized).toLowerCase();
      if (candidateName === normalizedNeedle || candidateNormalized === normalizedNeedle) {
        return candidate.id;
      }
    }

    cursor = normalizeText_(data.response_metadata && data.response_metadata.next_cursor);
    if (!cursor) {
      break;
    }
  }

  return '';
}

function normalizeStoredSlackPeopleLabels_() {
  if (!hasSlackBotToken_()) {
    return;
  }

  var requestRows = getRequestRows_();
  for (var i = 0; i < requestRows.length; i++) {
    var request = requestRows[i].data;
    var resolvedRequester = resolveSlackUserDisplayName_(request.SLACK_USER);
    if (resolvedRequester && resolvedRequester !== normalizeText_(request.SLACK_USER)) {
      request.SLACK_USER = resolvedRequester;
      updateRequestRow_(requestRows[i].rowNumber, request, requestRows[i].sheetName);
    }
  }

  var incidentRows = getIncidentRows_();
  for (var j = 0; j < incidentRows.length; j++) {
    var incident = incidentRows[j].data;
    var resolvedReporter = resolveSlackUserDisplayName_(incident.SLACK_USER);
    if (resolvedReporter && resolvedReporter !== normalizeText_(incident.SLACK_USER)) {
      incident.SLACK_USER = resolvedReporter;
      updateIncidentRow_(incidentRows[j].rowNumber, incident);
    }
  }
}

function resolveSlackUserDisplayName_(rawValue) {
  var normalized = normalizeText_(rawValue);
  if (!normalized) {
    return '';
  }

  if (!hasSlackBotToken_()) {
    return normalized;
  }

  var userId = extractSlackUserId_(normalized);
  if (!userId) {
    return normalized;
  }

  try {
    return lookupSlackUserDisplayName_(userId) || normalized;
  } catch (err) {
    Logger.log('Slack user lookup failed for ' + userId + ': ' + err.message);
    return normalized;
  }
}

function extractSlackUserId_(value) {
  var normalized = normalizeText_(value);
  if (!normalized) {
    return '';
  }

  var mentionMatch = normalized.match(/^<@([A-Z0-9]+)>$/);
  if (mentionMatch) {
    return mentionMatch[1];
  }

  if (/^[A-Z0-9]+$/.test(normalized) && normalized.charAt(0) === 'U') {
    return normalized;
  }

  return '';
}

function lookupSlackUserDisplayName_(userId) {
  var response = UrlFetchApp.fetch(
    'https://slack.com/api/users.info?user=' + encodeURIComponent(userId),
    {
      method: 'get',
      headers: { Authorization: 'Bearer ' + getSlackBotToken_() },
      muteHttpExceptions: true
    }
  );

  var data = JSON.parse(response.getContentText());
  if (!data.ok) {
    throw new Error('Slack users.info failed: ' + (data.error || response.getContentText()));
  }

  var user = data.user || {};
  var profile = user.profile || {};
  return normalizeText_(profile.display_name || profile.real_name || user.real_name || user.name || userId);
}

function getSlackPermalink_(channelId, messageTs) {
  var response = UrlFetchApp.fetch(
    'https://slack.com/api/chat.getPermalink?channel=' + encodeURIComponent(channelId) +
      '&message_ts=' + encodeURIComponent(messageTs),
    {
      method: 'get',
      headers: { Authorization: 'Bearer ' + getSlackBotToken_() },
      muteHttpExceptions: true
    }
  );

  var data = JSON.parse(response.getContentText());
  if (!data.ok) {
    throw new Error('Slack chat.getPermalink failed: ' + (data.error || response.getContentText()));
  }
  return data.permalink;
}

function buildSlackThreadRef_(channelId, messageTs) {
  if (!channelId || !messageTs) {
    return '';
  }
  return channelId + '|' + messageTs;
}

function parseSlackThreadRef_(threadLink) {
  var value = normalizeText_(threadLink);
  if (value.indexOf('|') !== -1) {
    var parts = value.split('|');
    if (parts.length === 2 && parts[0] && parts[1]) {
      return {
        channel: parts[0],
        threadTs: parts[1]
      };
    }
  }
  var match = value.match(/archives\/([A-Z0-9]+)\/p([0-9]+)/);
  if (!match) {
    return null;
  }

  var rawTs = match[2];
  if (rawTs.length <= 6) {
    return null;
  }

  return {
    channel: match[1],
    threadTs: rawTs.slice(0, rawTs.length - 6) + '.' + rawTs.slice(rawTs.length - 6)
  };
}

function postThreadReplyByUrl_(threadLink, text, blocks) {
  var ref = parseSlackThreadRef_(threadLink);
  if (!ref) {
    return null;
  }
  return postSlackMessage_(ref.channel, text, blocks, ref.threadTs);
}

function buildSectionBlock_(text) {
  return {
    type: 'section',
    text: { type: 'mrkdwn', text: text }
  };
}

function buildContextBlock_(text) {
  return {
    type: 'context',
    elements: [{ type: 'mrkdwn', text: text }]
  };
}

function testSlackConnection() {
  var response = postSlackMessageWithFallback_(
    getConfiguredAccessReviewChannels_(),
    'GMP security tracker connection test',
    [
      buildSectionBlock_('*GMP security tracker connection test*'),
      buildSectionBlock_('Workbook: `' + SpreadsheetApp.getActiveSpreadsheet().getName() + '`'),
      buildContextBlock_(formatTimestamp_(new Date()))
    ]
  );

  var permalink = getSlackPermalink_(response.channel, response.ts);
  try {
    SpreadsheetApp.getUi().alert(
      'Slack Connected',
      'Posted a test message successfully.\n' + permalink,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    Logger.log('Slack test message posted: ' + permalink);
  }
}

function sendOpenReviewSummary() {
  ensureThreeTabWorkbookReady_();
  if (!hasSlackBotToken_()) {
    Logger.log('Skipping review summary: SLACK_BOT_TOKEN is not configured.');
    return;
  }
  var reviewRows = getReviewRows_();
  var openRows = [];

  for (var i = 0; i < reviewRows.length; i++) {
    if (normalizeText_(reviewRows[i].data.DECISION_STATUS) !== 'Open') {
      continue;
    }
    openRows.push(reviewRows[i]);
  }

  if (openRows.length === 0) {
    return;
  }

  openRows.sort(function(a, b) {
    return priorityRank_(a.data.REVIEW_PRIORITY) - priorityRank_(b.data.REVIEW_PRIORITY);
  });

  var blocks = [
    buildSectionBlock_('*Open GMP review queue summary*'),
    buildContextBlock_('Open items: ' + openRows.length)
  ];

  var limit = Math.min(openRows.length, GMP_CONFIG.REVIEW_SUMMARY_LIMIT);
  for (var j = 0; j < limit; j++) {
    var review = openRows[j].data;
    blocks.push(
      buildSectionBlock_(
        '*`' + review.REVIEW_ID + '`* ' +
          normalizeText_(review.REVIEW_PRIORITY) + ' | ' +
          normalizeText_(review.TRIGGER_TYPE) + '\n' +
          normalizeText_(review.PERSON) + ' | ' +
          normalizeText_(review.SYSTEM) + ' | ' + normalizeText_(review.PRESENCE_STATUS || review.ACTIVITY_STATUS)
      )
    );
  }

  postSlackMessageWithFallback_(
    getConfiguredAccessReviewChannels_(),
    'Open GMP review queue summary',
    blocks
  );
}

function markAutomationColumns() {
  ensureThreeTabWorkbookReady_();

  for (var i = 0; i < AUTOMATION_MARKERS.length; i++) {
    var marker = AUTOMATION_MARKERS[i];
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(marker.sheetName);
    if (!sheet) {
      continue;
    }

    var rowNumber = typeof marker.row === 'function' ? marker.row() : marker.row;
    var columnNumber = typeof marker.column === 'function' ? marker.column() : marker.column;
    var cell = sheet.getRange(rowNumber, columnNumber);
    var baseHeader = normalizeHeaderText_(cell.getDisplayValue());
    cell.setValue(baseHeader + ' ' + marker.emoji);
    cell.setNote(marker.note);
  }
}

function tryMarkAutomationColumns_() {
  var issues = getThreeTabWorkbookIssues_();
  if (issues.length > 0) {
    Logger.log('Skipping automation markers: ' + issues.join(' '));
    return '- automation markers skipped until setupSheet() repairs the workbook schema';
  }

  markAutomationColumns();
  return '- automation markers on header columns';
}

function clearAutomationColumnMarkers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < AUTOMATION_MARKERS.length; i++) {
    var marker = AUTOMATION_MARKERS[i];
    var sheet = ss.getSheetByName(marker.sheetName);
    if (!sheet) {
      continue;
    }

    try {
      var rowNumber = typeof marker.row === 'function' ? marker.row() : marker.row;
      var columnNumber = typeof marker.column === 'function' ? marker.column() : marker.column;
      var cell = sheet.getRange(rowNumber, columnNumber);
      cell.setValue(normalizeHeaderText_(cell.getDisplayValue()));
      cell.clearNote();
    } catch (err) {
      Logger.log('Skipping marker cleanup for ' + marker.sheetName + ': ' + err.message);
    }
  }
}

function showSheetToast_(title, message, seconds) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, seconds || 4);
  } catch (err) {
    Logger.log((title || 'Notice') + ': ' + message);
  }
}

function onEditInstallable(e) {
  if (!e || !e.range) {
    return;
  }

  try {
    ensureThreeTabWorkbookReady_();
  } catch (err) {
    Logger.log('Schema not ready for onEdit: ' + err.message);
    return;
  }

  var sheet = e.range.getSheet();
  var rowNumber = e.range.getRow();
  var columnNumber = e.range.getColumn();
  if (rowNumber <= 1) {
    return;
  }

  if (sheet.getName() === GMP_SCHEMA.tabs.ACCESS_CONTROL) {
    handleAccessControlEdit_(rowNumber, columnNumber, e);
    return;
  }

  if (sheet.getName() === GMP_SCHEMA.tabs.REVIEW_CHECKS) {
    handleReviewCheckEdit_(rowNumber, columnNumber, e);
    return;
  }

  if (sheet.getName() === GMP_SCHEMA.tabs.INCIDENTS) {
    handleIncidentEdit_(rowNumber, columnNumber, e);
  }
}

function handleAccessControlEdit_(rowNumber, columnNumber, e) {
  var ctx = getAccessControlContext_();
  if (rowNumber <= ctx.requestHeaderRow) {
    return;
  }

  if (rowNumber < ctx.accountSectionRow) {
    handleRequestEdit_(getRequestRows_(), rowNumber, columnNumber);
    return;
  }

  if (ctx.archiveSectionRow !== -1 && rowNumber > ctx.archiveHeaderRow) {
    handleRequestEdit_(getArchivedRequestRows_(), rowNumber, columnNumber);
  }
}

function handleRequestEdit_(requestRows, rowNumber, columnNumber) {
  if (columnNumber !== GMP_SCHEMA.requests.columns.APPROVAL &&
      columnNumber !== GMP_SCHEMA.requests.columns.PROVISIONING) {
    return;
  }

  var requestRow = findTableRowByRowNumber_(requestRows, rowNumber);
  if (!requestRow) {
    return;
  }

  var result = syncAccessRequestRowToAccount_(requestRow);
  if (result) {
    archiveRequestRow_(requestRow);
    sortAccountSection_();
  }
  var request = findRequestRowById_(requestRow.data.REQUEST_ID);
  if (!request) {
    request = requestRow;
  }

  sendRequestThreadUpdate_(request.data, result);
}

function handleReviewCheckEdit_(rowNumber, columnNumber, e) {
  if (rowNumber === GMP_SCHEMA.layout.REVIEW_CONTROL_ROW) {
    if (columnNumber === GMP_SCHEMA.reviews.columns.NEXT_REVIEW_DUE) {
      handleBulkReviewCheckDateEdit_();
    }
    return;
  }

  if (rowNumber <= GMP_SCHEMA.layout.REVIEW_HEADER_ROW) {
    return;
  }

  if (columnNumber === GMP_SCHEMA.reviews.columns.NEXT_REVIEW_DUE) {
    handleReviewCheckDateEdit_(rowNumber);
    return;
  }

  if (columnNumber !== GMP_SCHEMA.reviews.columns.DECISION_STATUS) {
    return;
  }

  var reviewRow = findTableRowByRowNumber_(getReviewRows_(), rowNumber);
  if (!reviewRow) {
    return;
  }

  if (normalizeText_(reviewRow.data.DECISION_STATUS) === 'Done' &&
      normalizeText_(reviewRow.data.REVIEWER_ACTION) === '') {
    reviewRow.data.DECISION_STATUS = 'Open';
    updateReviewRow_(reviewRow.rowNumber, reviewRow.data);
    showSheetToast_('Review Action Required', 'Choose Reviewer Action before setting Review Status to Done.', 5);
    sortReviewSection_();
    formatReviewChecksSheet_(getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS));
    return;
  }

  if (normalizeText_(reviewRow.data.DECISION_STATUS) === 'Done' &&
      normalizeText_(reviewRow.data.REVIEWER_ACTION) === 'Need Info') {
    reviewRow.data.DECISION_STATUS = 'Waiting';
    updateReviewRow_(reviewRow.rowNumber, reviewRow.data);
    showSheetToast_('Use Waiting For Need Info', 'Need Info is a follow-up state. Use Waiting until the review is actually completed.', 5);
    sortReviewSection_();
    formatReviewChecksSheet_(getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS));
    return;
  }

  applyReviewDecisionToAccount_(reviewRow);
  sortReviewSection_();
  sortAccountSection_();
  formatReviewChecksSheet_(getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS));

  var review = findReviewRowById_(reviewRow.data.REVIEW_ID);
  if (!review) {
    review = reviewRow;
  }

  var threadLink = normalizeText_(review.data.SLACK_THREAD);
  if (!threadLink) {
    var accountRow = findAccountRowForReview_(review.data);
    threadLink = accountRow ? normalizeText_(accountRow.data.SLACK_THREAD) : '';
  }

  if (threadLink) {
    sendReviewThreadUpdate_(review.data, threadLink, review.rowNumber);
  }
}

function handleBulkReviewCheckDateEdit_() {
  var sheet = getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS);
  var controlCell = sheet.getRange(
    GMP_SCHEMA.layout.REVIEW_CONTROL_ROW,
    GMP_SCHEMA.reviews.columns.NEXT_REVIEW_DUE
  );
  var controlValue = controlCell.getValue();
  if (!controlValue) {
    return;
  }

  var today = new Date();
  var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var checkDate = controlValue instanceof Date && !isNaN(controlValue)
    ? new Date(controlValue.getFullYear(), controlValue.getMonth(), controlValue.getDate())
    : null;

  if (!checkDate || checkDate.getTime() <= todayStart.getTime()) {
    controlCell.clearContent();
    showSheetToast_('Future Date Required', 'Set a future Next Check Date in the orange control cell to schedule the open review queue.', 5);
    return;
  }

  var reviewRows = getReviewRows_();
  var updatedCount = 0;
  for (var i = 0; i < reviewRows.length; i++) {
    var review = reviewRows[i].data;
    if (normalizeText_(review.DECISION_STATUS) === 'Done') {
      continue;
    }

    review.NEXT_REVIEW_DUE = checkDate;
    review.DECISION_STATUS = 'Open';
    review.DECISION_DATE = '';
    review.NEXT_SLACK_TRIGGER = 'Resume review on check date';
    review.NOTES = appendUniqueNoteLine_(
      review.NOTES,
      formatDateOnly_(today) + ': bulk follow-up set ' + formatDateOnly_(checkDate)
    );
    updateReviewRow_(reviewRows[i].rowNumber, review);

    var accountRow = findAccountRowForReview_(review);
    if (accountRow) {
      var account = accountRow.data;
      account.REVIEW_STATUS = 'Open';
      account.NEXT_REVIEW_DUE = checkDate;
      account.NOTES = appendUniqueNoteLine_(
        account.NOTES,
        formatDateOnly_(today) + ': bulk review follow-up ' + formatDateOnly_(checkDate)
      );
      updateAccountRow_(accountRow.rowNumber, account);
    }

    updatedCount++;
  }

  controlCell.clearContent();

  if (updatedCount === 0) {
    showSheetToast_('No Active Reviews', 'There are no active review rows to schedule right now.', 4);
    return;
  }

  sortReviewSection_();
  sortAccountSection_();
  formatReviewChecksSheet_(sheet);
}

function handleReviewCheckDateEdit_(rowNumber) {
  var reviewRow = findTableRowByRowNumber_(getReviewRows_(), rowNumber);
  if (!reviewRow) {
    return;
  }

  var review = reviewRow.data;
  if (normalizeText_(review.DECISION_STATUS) === 'Done') {
    return;
  }

  var accountRow = findAccountRowForReview_(review);
  var account = accountRow ? accountRow.data : null;
  var today = new Date();
  var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var checkDate = review.NEXT_REVIEW_DUE instanceof Date && !isNaN(review.NEXT_REVIEW_DUE) ? review.NEXT_REVIEW_DUE : '';

  if (!checkDate) {
    review.DECISION_STATUS = 'Open';
    review.NEXT_SLACK_TRIGGER = 'Post review queue alert';
    review.NOTES = appendUniqueNoteLine_(review.NOTES, formatDateOnly_(today) + ': follow-up cleared');
    updateReviewRow_(reviewRow.rowNumber, review);

    if (account) {
      account.REVIEW_STATUS = 'Open';
      account.NEXT_REVIEW_DUE = today;
      account.NOTES = appendNoteLine_(account.NOTES, formatDateOnly_(today) + ': follow-up cleared');
      updateAccountRow_(accountRow.rowNumber, account);
    }

    sortReviewSection_();
    sortAccountSection_();
    formatReviewChecksSheet_(getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS));
    return;
  }

  var checkDay = new Date(checkDate.getFullYear(), checkDate.getMonth(), checkDate.getDate());
  var isFutureFollowUp = checkDay.getTime() > todayStart.getTime();

  review.DECISION_STATUS = 'Open';
  review.DECISION_DATE = '';
  review.NEXT_SLACK_TRIGGER = isFutureFollowUp ? 'Resume review on check date' : 'Post review queue alert';
  review.NOTES = appendUniqueNoteLine_(
    review.NOTES,
    formatDateOnly_(today) + ': ' + (isFutureFollowUp ? 'follow-up set ' : 'check date updated ') + formatDateOnly_(checkDate)
  );
  updateReviewRow_(reviewRow.rowNumber, review);

  if (account) {
    account.REVIEW_STATUS = review.DECISION_STATUS;
    account.NEXT_REVIEW_DUE = checkDate;
    account.NOTES = appendNoteLine_(
      account.NOTES,
      formatDateOnly_(today) + ': ' + (isFutureFollowUp ? 'review follow-up ' : 'review due ') + formatDateOnly_(checkDate)
    );
    updateAccountRow_(accountRow.rowNumber, account);
  }

  sortReviewSection_();
  sortAccountSection_();
  formatReviewChecksSheet_(getRequiredSheet_(GMP_SCHEMA.tabs.REVIEW_CHECKS));
}

function handleIncidentEdit_(rowNumber, columnNumber, e) {
  if (rowNumber <= GMP_SCHEMA.layout.INCIDENT_HEADER_ROW) {
    return;
  }
  if (columnNumber !== GMP_SCHEMA.incidents.columns.STATUS) {
    return;
  }

  var incidentRow = findTableRowByRowNumber_(getIncidentRows_(), rowNumber);
  if (!incidentRow) {
    return;
  }

  var incident = incidentRow.data;
  if (normalizeText_(incident.STATUS) === 'Resolved' && !incident.RESOLVED_AT) {
    incident.RESOLVED_AT = new Date();
    updateIncidentRow_(rowNumber, incident);
  }

  sortIncidentSection_();
  incidentRow = findIncidentRowById_(incidentRow.data.INCIDENT_ID) || incidentRow;
  incident = incidentRow.data;

  if (normalizeText_(incident.SLACK_THREAD)) {
    sendIncidentThreadUpdate_(incident, incidentRow.rowNumber);
  }
}

function applyReviewDecisionToAccount_(reviewRow) {
  var review = reviewRow.data;
  var accountRow = findAccountRowForReview_(review);
  if (!accountRow) {
    return;
  }

  var account = accountRow.data;
  account.REVIEW_STATUS = review.DECISION_STATUS;

  if (normalizeText_(review.DECISION_STATUS) !== 'Done') {
    updateAccountRow_(accountRow.rowNumber, account);
    return;
  }

  review.DECISION_DATE = review.DECISION_DATE || new Date();
  review.NEXT_REVIEW_DUE = '';
  updateReviewRow_(reviewRow.rowNumber, review);
  removeSiblingReviewRowsForCycle_(review.ACCESS_ID, review.REVIEW_CYCLE, review.REVIEW_ID, review.PERSON, review.SYSTEM, review.ACCESS_LEVEL);

  account.LAST_REVIEW_DATE = review.DECISION_DATE;
  if (normalizeText_(review.REVIEWER_ACTION) === 'Keep') {
    account.CURRENT_STATE = normalizePresenceState_(account.CURRENT_STATE);
    account.REVIEW_TRIGGER = '';
    account.REVIEW_STATUS = 'Done';
    account.NEXT_REVIEW_DUE = nextReviewDueFromDecision_(review.DECISION_DATE, account.GMP_SYSTEM, account.ACCESS_LEVEL, review.REVIEWER_ACTION);
    account.NOTES = appendUniqueNoteLine_(account.NOTES, formatDateOnly_(review.DECISION_DATE) + ': review completed Keep');
  } else if (normalizeText_(review.REVIEWER_ACTION) === 'Reduce') {
    account.CURRENT_STATE = normalizePresenceState_(account.CURRENT_STATE);
    account.REVIEW_STATUS = 'Done';
    account.NEXT_REVIEW_DUE = nextReviewDueFromDecision_(review.DECISION_DATE, account.GMP_SYSTEM, account.ACCESS_LEVEL, review.REVIEWER_ACTION);
    account.NOTES = appendUniqueNoteLine_(account.NOTES, formatDateOnly_(review.DECISION_DATE) + ': review completed Reduce');
  } else if (normalizeText_(review.REVIEWER_ACTION) === 'Remove') {
    account.REVIEW_STATUS = 'Done';
    account.REVIEW_TRIGGER = 'Removal Requested';
    account.NEXT_REVIEW_DUE = nextReviewDueFromDecision_(review.DECISION_DATE, account.GMP_SYSTEM, account.ACCESS_LEVEL, review.REVIEWER_ACTION);
    account.NOTES = appendUniqueNoteLine_(account.NOTES, formatDateOnly_(review.DECISION_DATE) + ': review completed Remove');
    ensureRemovalRequestForReview_(account, review);
  } else if (normalizeText_(review.REVIEWER_ACTION) === 'Need Info') {
    account.REVIEW_STATUS = 'Waiting';
  }

  updateAccountRow_(accountRow.rowNumber, account);
}

function ensureRemovalRequestForReview_(account, review) {
  var requestRows = getRequestRows_();
  for (var i = 0; i < requestRows.length; i++) {
    var request = requestRows[i].data;
    if (normalizeText_(request.ACCESS_ID) !== normalizeText_(account.ACCESS_ID)) {
      continue;
    }
    if (normalizeText_(request.REQUEST_TYPE) !== 'Remove Access') {
      continue;
    }
    if (normalizeSetupStatus_(request.PROVISIONING) === 'Closed' || normalizeSetupStatus_(request.PROVISIONING) === 'Completed') {
      return;
    }
    return;
  }

  var requestRecord = buildEmptyRecord_(GMP_SCHEMA.requests.keys);
  requestRecord.REQUEST_ID = nextStableId_('REQ');
  requestRecord.SUBMITTED_AT = new Date();
  requestRecord.QUEUE_PRIORITY = 'High';
  requestRecord.REQUEST_TYPE = 'Remove Access';
  requestRecord.SLACK_USER = account.OWNER || '';
  requestRecord.TARGET_USER = account.PERSON;
  requestRecord.COMPANY_EMAIL = account.COMPANY_EMAIL;
  requestRecord.DEPT = account.DEPT;
  requestRecord.GMP_SYSTEM = account.GMP_SYSTEM;
  requestRecord.ACCESS_LEVEL = account.ACCESS_LEVEL;
  requestRecord.GMP_IMPACT = 'Yes';
  requestRecord.REASON = 'Created automatically from review decision ' + review.REVIEW_ID;
  requestRecord.MANAGER = '';
  requestRecord.IT_OWNER = account.OWNER;
  requestRecord.APPROVAL = 'Approved';
  requestRecord.PROVISIONING = 'Removing Access';
  requestRecord.NEXT_SLACK_TRIGGER = 'Confirm removal in Slack';
  requestRecord.SLA_DUE = addHours_(new Date(), 4);
  requestRecord.DECISION_DATE = new Date();
  requestRecord.ACCESS_ID = account.ACCESS_ID;
  requestRecord.LINKED_REVIEW_ID = review.REVIEW_ID;
  requestRecord.SLACK_THREAD = account.SLACK_THREAD;
  requestRecord.NOTES = 'Auto-created from review action Remove.';

  appendAccessRequestRecord_(requestRecord);
}

function findTableRowByRowNumber_(rows, rowNumber) {
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].rowNumber === rowNumber) {
      return rows[i];
    }
  }
  return null;
}

function priorityRank_(priority) {
  var value = normalizeText_(priority);
  if (value === 'High') {
    return 0;
  }
  if (value === 'Medium') {
    return 1;
  }
  return 2;
}

function sendRequestThreadUpdate_(request, result) {
  var threadLink = normalizeText_(request.SLACK_THREAD);
  if (!threadLink) {
    return;
  }

  var notification = buildRequestNotification_(request, result);
  if (!notification) {
    return;
  }

  if (!shouldSendNotificationState_('REQ', request.REQUEST_ID, notification.stateKey)) {
    return;
  }

  postThreadReplyByUrl_(threadLink, notification.text, notification.blocks);
}

function buildRequestNotification_(request, result) {
  var approval = normalizeText_(request.APPROVAL);
  var provisioning = normalizeSetupStatus_(request.PROVISIONING);
  var accessId = normalizeText_(request.ACCESS_ID || (result && result.accessId));
  var requestRow = findRequestRowById_(request.REQUEST_ID);
  var rowLink = requestRow ? getRowLink_(requestRow.sheetName || GMP_SCHEMA.tabs.ACCESS_CONTROL, requestRow.rowNumber) : '';

  if (approval === 'Denied') {
    return {
      stateKey: ['Denied', provisioning || 'Open', accessId].join('|'),
      text: 'Access request denied: ' + request.REQUEST_ID,
      blocks: [
        buildSectionBlock_('*Access request denied*'),
        buildSectionBlock_(
          '*Request:* `' + request.REQUEST_ID + '`\n' +
            '*Target:* ' + normalizeText_(request.TARGET_USER) + '\n' +
            '*System:* `' + normalizeText_(request.GMP_SYSTEM) + '`\n' +
            '*Access Level:* `' + normalizeText_(request.ACCESS_LEVEL) + '`'
        ),
        buildSectionBlock_(
          '*Approval:* `' + approval + '`\n' +
            '*Access Setup Status:* `' + (provisioning || 'Open') + '`\n' +
            '*IT Owner:* `' + normalizeText_(request.IT_OWNER) + '`'
        ),
        buildContextBlock_(rowLink)
      ]
    };
  }

  if (approval !== 'Approved') {
    return null;
  }

  if (provisioning !== 'In Progress' &&
      provisioning !== 'Completed' &&
      provisioning !== 'Removing Access' &&
      provisioning !== 'Closed') {
    return null;
  }

  var title = 'Access request updated';
  if (provisioning === 'In Progress') {
    title = 'Access setup in progress';
  } else if (provisioning === 'Completed') {
    title = 'Access granted';
  } else if (provisioning === 'Removing Access') {
    title = 'Access removal in progress';
  } else if (provisioning === 'Closed') {
    title = 'Access workflow closed';
  }

  return {
    stateKey: ['Approved', provisioning, accessId].join('|'),
    text: title + ': ' + request.REQUEST_ID,
    blocks: [
      buildSectionBlock_('*' + title + '*'),
      buildSectionBlock_(
        '*Request:* `' + request.REQUEST_ID + '`\n' +
          '*Target:* ' + normalizeText_(request.TARGET_USER) + '\n' +
          '*System:* `' + normalizeText_(request.GMP_SYSTEM) + '`\n' +
          '*Access Level:* `' + normalizeText_(request.ACCESS_LEVEL) + '`'
      ),
      buildSectionBlock_(
        '*Approval:* `' + approval + '`\n' +
          '*Access Setup Status:* `' + provisioning + '`\n' +
          '*IT Owner:* `' + normalizeText_(request.IT_OWNER) + '`\n' +
          '*Access ID:* `' + accessId + '`'
      ),
      buildSectionBlock_('*Sync Result:* `' + normalizeText_(result && result.action) + '`'),
      buildContextBlock_(rowLink)
    ]
  };
}

function sendReviewThreadUpdate_(review, threadLink, rowNumber) {
  var notification = buildReviewNotification_(review, rowNumber);
  if (!notification) {
    return;
  }

  if (!shouldSendNotificationState_('REV', review.REVIEW_ID, notification.stateKey)) {
    return;
  }

  postThreadReplyByUrl_(threadLink, notification.text, notification.blocks);
}

function buildReviewNotification_(review, rowNumber) {
  var decision = normalizeText_(review.DECISION_STATUS);
  if (!decision) {
    return null;
  }

  return {
    stateKey: [decision, normalizeText_(review.REVIEWER_ACTION), formatTimestamp_(review.DECISION_DATE)].join('|'),
    text: 'Review updated: ' + review.REVIEW_ID,
    blocks: [
      buildSectionBlock_('*Review update*'),
      buildSectionBlock_(
        '*Review:* `' + review.REVIEW_ID + '`\n' +
          '*Cycle:* `' + normalizeText_(review.REVIEW_CYCLE) + '`\n' +
          '*Decision:* `' + decision + '`\n' +
          '*Action:* `' + normalizeText_(review.REVIEWER_ACTION) + '`'
      ),
      buildSectionBlock_(
        '*Access ID:* `' + normalizeText_(review.ACCESS_ID) + '`\n' +
          '*System:* `' + normalizeText_(review.SYSTEM) + '`\n' +
          '*Person:* ' + normalizeText_(review.PERSON) + '\n' +
          '*Access Level:* `' + normalizeText_(review.ACCESS_LEVEL) + '`\n' +
          '*Presence:* `' + normalizeText_(review.PRESENCE_STATUS) + '`\n' +
          '*Activity:* `' + normalizeText_(review.ACTIVITY_STATUS) + '`'
      ),
      buildContextBlock_(getRowLink_(GMP_SCHEMA.tabs.REVIEW_CHECKS, rowNumber))
    ]
  };
}

function sendIncidentThreadUpdate_(incident, rowNumber) {
  var notification = buildIncidentNotification_(incident, rowNumber);
  if (!notification) {
    return;
  }

  if (!shouldSendNotificationState_('INC', incident.INCIDENT_ID, notification.stateKey)) {
    return;
  }

  postThreadReplyByUrl_(incident.SLACK_THREAD, notification.text, notification.blocks);
}

function buildIncidentNotification_(incident, rowNumber) {
  var status = normalizeText_(incident.STATUS);
  if (!status) {
    return null;
  }

  return {
    stateKey: [status, normalizeText_(incident.ASSIGNED_TO), normalizeText_(incident.RESOLUTION)].join('|'),
    text: 'Incident updated: ' + incident.INCIDENT_ID,
    blocks: [
      buildSectionBlock_('*Incident update*'),
      buildSectionBlock_(
        '*Incident:* `' + incident.INCIDENT_ID + '`\n' +
          '*Status:* `' + status + '`\n' +
          '*Severity:* `' + normalizeText_(incident.SEVERITY) + '`'
      ),
      buildSectionBlock_(
        '*Assigned:* `' + normalizeText_(incident.ASSIGNED_TO) + '`\n' +
          '*System:* `' + normalizeText_(incident.SYSTEM) + '`\n' +
          '*Summary:* ' + normalizeText_(incident.SUMMARY)
      ),
      buildSectionBlock_('*Resolution:* ' + normalizeText_(incident.RESOLUTION)),
      buildContextBlock_(getRowLink_(GMP_SCHEMA.tabs.INCIDENTS, rowNumber))
    ]
  };
}

function shouldSendNotificationState_(entityPrefix, entityId, stateKey) {
  if (!entityId || !stateKey) {
    return false;
  }

  var propertyName = 'LAST_NOTIFY_' + entityPrefix + '_' + entityId;
  var props = PropertiesService.getDocumentProperties();
  var previous = props.getProperty(propertyName);
  if (previous === stateKey) {
    return false;
  }

  props.setProperty(propertyName, stateKey);
  return true;
}
