var GMP_SLASH_CONFIG = {
  ACCESS_COMMAND: '/gmp-access',
  ACCESS_ALIAS_COMMAND: '/gmp-request-access',
  ISSUE_COMMAND: '/gmp-issue',
  ISSUE_ALIAS_COMMAND: '/it-alert',
  ACCESS_CALLBACK: 'gmp_access_modal',
  ISSUE_CALLBACK: 'gmp_issue_modal'
};

function doPost(e) {
  return routeSecurityTrackerPost_(e);
}

function doGet() {
  return ContentService.createTextOutput('GMP security tracker endpoint active.');
}

function routeSecurityTrackerPost_(e) {
  try {
    if (!verifySlackWebRequest_(e)) {
      Logger.log('Rejected Slack request: verification token mismatch.');
      return slackAckText_(':warning: Unauthorized Slack request.');
    }

    if (e.parameter &&
        (e.parameter.command === GMP_SLASH_CONFIG.ACCESS_COMMAND ||
         e.parameter.command === GMP_SLASH_CONFIG.ACCESS_ALIAS_COMMAND)) {
      return handleAccessSlashCommand_(e.parameter);
    }

    if (e.parameter &&
        (e.parameter.command === GMP_SLASH_CONFIG.ISSUE_COMMAND ||
         e.parameter.command === GMP_SLASH_CONFIG.ISSUE_ALIAS_COMMAND)) {
      return handleIssueSlashCommand_(e.parameter);
    }

    if (e.parameter && e.parameter.payload) {
      return handleSlackInteraction_(JSON.parse(e.parameter.payload));
    }

    return slackAck_();
  } catch (err) {
    Logger.log('routeSecurityTrackerPost_ error: ' + err + '\n' + err.stack);
    return slackAckText_(':warning: GMP security tracker failed: ' + err.message);
  }
}

function verifySlackWebRequest_(e) {
  var expectedToken = normalizeText_(
    PropertiesService.getScriptProperties().getProperty(GMP_CONFIG.VERIFICATION_TOKEN_PROPERTY)
  );
  if (!expectedToken) {
    return true;
  }
  return extractSlackVerificationToken_(e) === expectedToken;
}

function extractSlackVerificationToken_(e) {
  var directToken = normalizeText_(e && e.parameter && e.parameter.token);
  if (directToken) {
    return directToken;
  }

  var rawPayload = normalizeText_(e && e.parameter && e.parameter.payload);
  if (!rawPayload) {
    return '';
  }

  try {
    return normalizeText_(JSON.parse(rawPayload).token);
  } catch (err) {
    return '';
  }
}

function handleAccessSlashCommand_(params) {
  if (!params.trigger_id) {
    return slackAckText_(':warning: Slack did not provide a trigger id.');
  }

  var result = openSlackModal_(params.trigger_id, buildAccessRequestModal_(params));
  if (!result.ok) {
    return slackAckText_(':warning: Could not open /gmp-access modal: ' + (result.detail || result.error || 'unknown error'));
  }
  return slackAck_();
}

function handleIssueSlashCommand_(params) {
  if (!params.trigger_id) {
    return slackAckText_(':warning: Slack did not provide a trigger id.');
  }

  var result = openSlackModal_(params.trigger_id, buildIssueModal_(params));
  if (!result.ok) {
    return slackAckText_(':warning: Could not open /gmp-issue modal: ' + (result.detail || result.error || 'unknown error'));
  }
  return slackAck_();
}

function openSlackModal_(triggerId, view) {
  var response = UrlFetchApp.fetch('https://slack.com/api/views.open', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + getSlackBotToken_() },
    payload: JSON.stringify({
      trigger_id: triggerId,
      view: view
    }),
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());
  if (!result.ok && result.response_metadata && result.response_metadata.messages) {
    result.detail = result.response_metadata.messages.join(' | ');
  }
  return result;
}

function handleSlackInteraction_(interaction) {
  if (interaction.type !== 'view_submission' || !interaction.view) {
    return slackAck_();
  }

  if (interaction.view.callback_id === GMP_SLASH_CONFIG.ACCESS_CALLBACK) {
    return handleAccessSubmission_(interaction);
  }

  if (interaction.view.callback_id === GMP_SLASH_CONFIG.ISSUE_CALLBACK) {
    return handleIssueSubmission_(interaction);
  }

  return slackAck_();
}

function handleAccessSubmission_(interaction) {
  var requestRecord = buildAccessRequestRecord_(interaction);
  if (!reserveSecuritySubmissionKey_(requestRecord._SUBMISSION_KEY, requestRecord.REQUEST_ID)) {
    Logger.log('Skipping duplicate access modal submission: ' + requestRecord._SUBMISSION_KEY);
    return slackClearModal_();
  }
  queueSecurityIntakeJob_('access', sanitizeRecordForQueue_(requestRecord));
  return slackClearModal_();
}

function handleIssueSubmission_(interaction) {
  var incidentRecord = buildIncidentRecord_(interaction);
  if (!reserveSecuritySubmissionKey_(incidentRecord._SUBMISSION_KEY, incidentRecord.INCIDENT_ID)) {
    Logger.log('Skipping duplicate incident modal submission: ' + incidentRecord._SUBMISSION_KEY);
    return slackClearModal_();
  }
  queueSecurityIntakeJob_('incident', sanitizeRecordForQueue_(incidentRecord));
  return slackClearModal_();
}

function buildAccessRequestRecord_(interaction) {
  var metadata = parseModalMetadata_(interaction.view.private_metadata, GMP_SLASH_CONFIG.ACCESS_COMMAND);
  var values = interaction.view.state.values;
  var requestType = getSelectedValue_(values, 'blk_request_type', 'request_type');
  var targetUser = getInputValue_(values, 'blk_target_user', 'target_user');
  var companyEmail = getInputValue_(values, 'blk_company_email', 'company_email');
  var department = getInputValue_(values, 'blk_department', 'department');
  var systemName = getInputValue_(values, 'blk_system', 'system_name');
  var accessLevel = getSelectedValue_(values, 'blk_access_level', 'access_level');
  var reason = getInputValue_(values, 'blk_reason', 'reason');
  var manager = getInputValue_(values, 'blk_manager', 'manager');
  var gmpImpact = getSelectedValue_(values, 'blk_gmp_impact', 'gmp_impact');
  var accessId = getInputValue_(values, 'blk_access_id', 'access_id');

  var record = buildEmptyRecord_(GMP_SCHEMA.requests.keys);
  record.REQUEST_ID = nextStableId_('REQ');
  record.SUBMITTED_AT = new Date();
  record.QUEUE_PRIORITY = deriveRequestPriority_(requestType, gmpImpact);
  record.REQUEST_TYPE = requestType;
  record.SLACK_USER = slackUserReference_(interaction.user);
  record.TARGET_USER = targetUser;
  record.COMPANY_EMAIL = companyEmail;
  record.DEPT = department;
  record.GMP_SYSTEM = canonicalSystemName_(systemName);
  record.ACCESS_LEVEL = accessLevel;
  record.GMP_IMPACT = gmpImpact;
  record.REASON = reason;
  record.MANAGER = manager;
  record.IT_OWNER = '';
  record.APPROVAL = 'Submitted';
  record.PROVISIONING = 'Open';
  record.NEXT_SLACK_TRIGGER = 'Notify IT review channel';
  record.SLA_DUE = defaultRequestSlaDue_(record.QUEUE_PRIORITY);
  record.DECISION_DATE = '';
  record.ACCESS_ID = accessId;
  record.LINKED_REVIEW_ID = '';
  record.LINKED_INCIDENT_ID = '';
  record.SLACK_THREAD = '';
  record.NOTES = 'Created from ' + metadata.sourceCommand;
  record._SUBMISSION_KEY = buildSlackInteractionKey_(interaction, 'access');
  record._SOURCE_CHANNEL_ID = metadata.sourceChannelId;
  record._SOURCE_CHANNEL_NAME = metadata.sourceChannelName;
  return record;
}

function buildIncidentRecord_(interaction) {
  var metadata = parseModalMetadata_(interaction.view.private_metadata, GMP_SLASH_CONFIG.ISSUE_COMMAND);
  var values = interaction.view.state.values;
  var systemName = getInputValue_(values, 'blk_system', 'system_name');
  var severity = getSelectedValue_(values, 'blk_severity', 'severity');
  var issueType = getSelectedValue_(values, 'blk_issue_type', 'issue_type');
  var summary = getInputValue_(values, 'blk_summary', 'summary');
  var details = getInputValue_(values, 'blk_details', 'details');
  var firstSeen = getInputValue_(values, 'blk_first_seen', 'first_seen');
  var ongoing = hasCheckboxValue_(values, 'blk_ongoing', 'ongoing');
  var linkedAccessId = getInputValue_(values, 'blk_linked_access', 'linked_access');
  var linkedRequestId = getInputValue_(values, 'blk_linked_request', 'linked_request');
  var responseTarget = getInputValue_(values, 'blk_response_target', 'response_target');
  var sourceCommand = metadata.sourceCommand;

  var record = buildEmptyRecord_(GMP_SCHEMA.incidents.keys);
  record.INCIDENT_ID = nextStableId_('INC');
  record.REPORTED_AT = new Date();
  record.SEVERITY = severity;
  record.SYSTEM = canonicalSystemName_(systemName);
  record.SLACK_USER = slackUserReference_(interaction.user);
  record.ISSUE_TYPE = issueType;
  record.SUMMARY = summary;
  record.LINKED_ACCESS_ID = linkedAccessId;
  record.LINKED_REQUEST_ID = linkedRequestId;
  record.ASSIGNED_TO = '';
  record.STATUS = 'Open';
  record.NEXT_SLACK_TRIGGER = 'Notify IT support';
  record.RESPONSE_TARGET = responseTarget;
  record.ROOT_CAUSE = '';
  record.RESOLUTION = '';
  record.RESOLVED_AT = '';
  record.SLACK_THREAD = '';
  record.NOTES = buildIncidentNotes_(sourceCommand, firstSeen, ongoing, details);
  record._SUBMISSION_KEY = buildSlackInteractionKey_(interaction, 'incident');
  record._DETAILS = details;
  record._FIRST_SEEN = firstSeen;
  record._ONGOING = ongoing ? 'Yes' : 'No';
  record._SOURCE_COMMAND = sourceCommand;
  record._SOURCE_CHANNEL_ID = metadata.sourceChannelId;
  record._SOURCE_CHANNEL_NAME = metadata.sourceChannelName;
  return record;
}

function buildAccessRequestBlocks_(record, rowNumber) {
  return [
    buildSectionBlock_('*New GMP access request*'),
    buildSectionBlock_(
      '*Request:* `' + record.REQUEST_ID + '`\n' +
        '*Type:* `' + record.REQUEST_TYPE + '`\n' +
        '*Priority:* `' + record.QUEUE_PRIORITY + '`'
    ),
    buildSectionBlock_(
      '*Target:* ' + record.TARGET_USER + '\n' +
        '*Email:* `' + record.COMPANY_EMAIL + '`\n' +
        '*Dept:* `' + record.DEPT + '`'
    ),
    buildSectionBlock_(
      '*System:* `' + record.GMP_SYSTEM + '`\n' +
        '*Access Level:* `' + record.ACCESS_LEVEL + '`\n' +
        '*GMP Impact:* `' + record.GMP_IMPACT + '`'
    ),
    buildSectionBlock_('*Reason:* ' + record.REASON),
    buildSectionBlock_(
      '*Manager:* `' + record.MANAGER + '`\n' +
        '*Requester:* ' + record.SLACK_USER + '\n' +
        '*Sheet Row:* <' + getRowLink_(GMP_SCHEMA.tabs.ACCESS_CONTROL, rowNumber) + '|Open Access Control>'
    )
  ];
}

function buildIncidentBlocks_(record, rowNumber) {
  var title = record._SOURCE_COMMAND === GMP_SLASH_CONFIG.ISSUE_ALIAS_COMMAND ? 'New IT alert' : 'New GMP issue report';
  return [
    buildSectionBlock_('*' + title + '*'),
    buildSectionBlock_(
      '*Incident:* `' + record.INCIDENT_ID + '`\n' +
        '*Severity:* `' + record.SEVERITY + '`\n' +
        '*Type:* `' + record.ISSUE_TYPE + '`'
    ),
    buildSectionBlock_(
      '*System:* `' + record.SYSTEM + '`\n' +
        '*Reporter:* ' + record.SLACK_USER + '\n' +
        '*Response Target:* `' + normalizeText_(record.RESPONSE_TARGET || 'Slack thread') + '`\n' +
        '*Still Happening:* `' + normalizeText_(record._ONGOING || 'No') + '`'
    ),
    buildSectionBlock_('*Summary:* ' + record.SUMMARY),
    buildSectionBlock_(
      '*First Noticed:* ' + normalizeText_(record._FIRST_SEEN || 'Not specified') + '\n' +
        '*Details:* ' + normalizeText_(record._DETAILS || 'No extra details provided')
    ),
    buildSectionBlock_(
      '*Linked Access:* `' + normalizeText_(record.LINKED_ACCESS_ID) + '`\n' +
        '*Linked Request:* `' + normalizeText_(record.LINKED_REQUEST_ID) + '`\n' +
        '*Sheet Row:* <' + getRowLink_(GMP_SCHEMA.tabs.INCIDENTS, rowNumber) + '|Open Incidents>'
    )
  ];
}

function buildAccessRequestModal_(params) {
  var metadata = buildModalMetadata_(params, GMP_SLASH_CONFIG.ACCESS_COMMAND);
  return {
    type: 'modal',
    callback_id: GMP_SLASH_CONFIG.ACCESS_CALLBACK,
    private_metadata: metadata,
    title: { type: 'plain_text', text: 'GMP Access' },
    submit: { type: 'plain_text', text: 'Submit' },
    close: { type: 'plain_text', text: 'Cancel' },
    blocks: [
      {
        type: 'input',
        block_id: 'blk_request_type',
        label: { type: 'plain_text', text: 'Request Type' },
        element: {
          type: 'static_select',
          action_id: 'request_type',
          options: staticOptions_(['New Access', 'Change Access', 'Remove Access'])
        }
      },
      textInputBlock_('blk_target_user', 'target_user', 'Target User', false, false),
      textInputBlock_('blk_company_email', 'company_email', 'Company Email', false, false),
      textInputBlock_('blk_department', 'department', 'Department', false, false),
      textInputBlock_('blk_system', 'system_name', 'GMP System', false, false),
      {
        type: 'input',
        block_id: 'blk_access_level',
        label: { type: 'plain_text', text: 'Access Level' },
        element: {
          type: 'static_select',
          action_id: 'access_level',
          options: staticOptions_(['Admin', 'Full Access', 'Read-Write', 'Read-Only', 'Limited'])
        }
      },
      textInputBlock_('blk_reason', 'reason', 'Reason', true, false),
      textInputBlock_('blk_manager', 'manager', 'Manager', false, false),
      {
        type: 'input',
        block_id: 'blk_gmp_impact',
        label: { type: 'plain_text', text: 'GMP Impact' },
        element: {
          type: 'radio_buttons',
          action_id: 'gmp_impact',
          options: staticOptions_(['Yes', 'No'])
        }
      },
      textInputBlock_('blk_access_id', 'access_id', 'Current Access ID', false, true)
    ]
  };
}

function buildIssueModal_(params) {
  var metadata = buildModalMetadata_(params, GMP_SLASH_CONFIG.ISSUE_COMMAND);
  var parsedMetadata = parseModalMetadata_(metadata, GMP_SLASH_CONFIG.ISSUE_COMMAND);
  var sourceCommand = parsedMetadata.sourceCommand;
  var isItAlert = sourceCommand === GMP_SLASH_CONFIG.ISSUE_ALIAS_COMMAND;
  var title = isItAlert ? 'IT Alert' : 'GMP Issue';
  var introText = isItAlert ?
    ':rotating_light: Report one alert with enough context for IT to act immediately. This posts to the support thread and logs a single incident row.' :
    ':rotating_light: Report a GMP issue once with enough detail for follow-up. This posts to the support thread and logs a single incident row.';
  return {
    type: 'modal',
    callback_id: GMP_SLASH_CONFIG.ISSUE_CALLBACK,
    private_metadata: metadata,
    title: { type: 'plain_text', text: title },
    submit: { type: 'plain_text', text: 'Submit' },
    close: { type: 'plain_text', text: 'Cancel' },
    blocks: [
      {
        type: 'context',
        elements: [{ type: 'mrkdwn', text: introText }]
      },
      { type: 'divider' },
      {
        type: 'input',
        block_id: 'blk_system',
        label: { type: 'plain_text', text: 'System or Workflow' },
        hint: { type: 'plain_text', text: 'Name the automation, platform, or process that is failing.' },
        element: {
          type: 'plain_text_input',
          action_id: 'system_name',
          placeholder: { type: 'plain_text', text: 'Katana Inventory, Shopify Admin, Google Drive GMP Folder...' }
        }
      },
      {
        type: 'input',
        block_id: 'blk_severity',
        label: { type: 'plain_text', text: 'Severity' },
        hint: { type: 'plain_text', text: 'Pick the urgency level IT should respond to first.' },
        element: {
          type: 'radio_buttons',
          action_id: 'severity',
          options: staticOptions_([
            'High - urgent / people blocked',
            'Medium - degraded but usable',
            'Low - minor / informational'
          ])
        }
      },
      {
        type: 'input',
        block_id: 'blk_issue_type',
        label: { type: 'plain_text', text: 'Issue Type' },
        element: {
          type: 'static_select',
          action_id: 'issue_type',
          options: staticOptions_(['Access Problem', 'Permission Error', 'Sync Failure', 'Outage', 'Data Mismatch', 'Other'])
        }
      },
      {
        type: 'input',
        block_id: 'blk_summary',
        label: { type: 'plain_text', text: 'Short Summary' },
        hint: { type: 'plain_text', text: 'This should read well in the Slack incident post title.' },
        element: {
          type: 'plain_text_input',
          action_id: 'summary',
          placeholder: { type: 'plain_text', text: 'Orders are not syncing from Shopify to Katana.' }
        }
      },
      {
        type: 'input',
        block_id: 'blk_details',
        label: { type: 'plain_text', text: 'What Happened?' },
        hint: { type: 'plain_text', text: 'Include what is blocked, what you checked, and any pattern you noticed.' },
        element: {
          type: 'plain_text_input',
          action_id: 'details',
          multiline: true,
          placeholder: { type: 'plain_text', text: 'Example: changes save in Shopify but never appear in Katana. Last good sync was earlier today.' }
        }
      },
      {
        type: 'input',
        block_id: 'blk_ongoing',
        optional: true,
        label: { type: 'plain_text', text: 'Current State' },
        element: {
          type: 'checkboxes',
          action_id: 'ongoing',
          options: staticOptions_(['Still happening'])
        }
      },
      {
        type: 'input',
        block_id: 'blk_first_seen',
        optional: true,
        label: { type: 'plain_text', text: 'First Noticed' },
        hint: { type: 'plain_text', text: 'Relative times are fine.' },
        element: {
          type: 'plain_text_input',
          action_id: 'first_seen',
          placeholder: { type: 'plain_text', text: 'Today at 8:15 AM, after the 7 AM batch, yesterday afternoon...' }
        }
      },
      { type: 'divider' },
      {
        type: 'context',
        elements: [{ type: 'mrkdwn', text: '*Optional links and routing*' }]
      },
      textInputBlock_('blk_linked_access', 'linked_access', 'Linked Access ID', false, true),
      textInputBlock_('blk_linked_request', 'linked_request', 'Linked Request ID', false, true),
      {
        type: 'input',
        block_id: 'blk_response_target',
        optional: true,
        label: { type: 'plain_text', text: 'Who Needs The Next Update?' },
        hint: { type: 'plain_text', text: 'Person, team, or channel to mention in follow-up.' },
        element: {
          type: 'plain_text_input',
          action_id: 'response_target',
          placeholder: { type: 'plain_text', text: 'Warehouse lead, QA, requester, #it-support...' }
        }
      }
    ]
  };
}

function textInputBlock_(blockId, actionId, labelText, multiline, optional) {
  return {
    type: 'input',
    block_id: blockId,
    optional: !!optional,
    label: { type: 'plain_text', text: labelText },
    element: {
      type: 'plain_text_input',
      action_id: actionId,
      multiline: !!multiline
    }
  };
}

function staticOptions_(values) {
  var options = [];
  for (var i = 0; i < values.length; i++) {
    options.push({
      text: { type: 'plain_text', text: values[i] },
      value: values[i]
    });
  }
  return options;
}

function getInputValue_(stateValues, blockId, actionId) {
  var block = stateValues[blockId];
  if (!block || !block[actionId]) {
    return '';
  }
  return normalizeText_(block[actionId].value);
}

function getSelectedValue_(stateValues, blockId, actionId) {
  var block = stateValues[blockId];
  if (!block || !block[actionId]) {
    return '';
  }
  if (block[actionId].selected_option) {
    return normalizeSelectedOptionValue_(block[actionId].selected_option.value);
  }
  return '';
}

function hasCheckboxValue_(stateValues, blockId, actionId) {
  var block = stateValues[blockId];
  if (!block || !block[actionId] || !block[actionId].selected_options) {
    return false;
  }
  return block[actionId].selected_options.length > 0;
}

function normalizeSelectedOptionValue_(value) {
  var normalized = normalizeText_(value);
  if (normalized.indexOf('High - ') === 0) {
    return 'High';
  }
  if (normalized.indexOf('Medium - ') === 0) {
    return 'Medium';
  }
  if (normalized.indexOf('Low - ') === 0) {
    return 'Low';
  }
  return normalized;
}

function buildModalMetadata_(params, fallbackCommand) {
  if (typeof params === 'string') {
    return JSON.stringify({
      sourceCommand: normalizeText_(params) || fallbackCommand,
      sourceChannelId: '',
      sourceChannelName: ''
    });
  }

  return JSON.stringify({
    sourceCommand: normalizeText_(params && params.command) || fallbackCommand,
    sourceChannelId: normalizeText_(params && params.channel_id),
    sourceChannelName: normalizeText_(params && params.channel_name)
  });
}

function parseModalMetadata_(rawMetadata, fallbackCommand) {
  var value = normalizeText_(rawMetadata);
  if (!value) {
    return {
      sourceCommand: fallbackCommand,
      sourceChannelId: '',
      sourceChannelName: ''
    };
  }

  if (value.charAt(0) !== '{') {
    return {
      sourceCommand: value || fallbackCommand,
      sourceChannelId: '',
      sourceChannelName: ''
    };
  }

  try {
    var parsed = JSON.parse(value);
    return {
      sourceCommand: normalizeText_(parsed.sourceCommand) || fallbackCommand,
      sourceChannelId: normalizeText_(parsed.sourceChannelId),
      sourceChannelName: normalizeText_(parsed.sourceChannelName)
    };
  } catch (err) {
    return {
      sourceCommand: fallbackCommand,
      sourceChannelId: '',
      sourceChannelName: ''
    };
  }
}

function buildIncidentNotes_(sourceCommand, firstSeen, ongoing, details) {
  var notes = ['Created from ' + sourceCommand];
  notes.push('Still happening: ' + (ongoing ? 'Yes' : 'No'));
  if (firstSeen) {
    notes.push('First noticed: ' + firstSeen);
  }
  if (details) {
    notes.push('Details: ' + details);
  }
  return notes.join('\n');
}

function tryPostInitialAccessNotification_(requestRecord, rowNumber) {
  try {
    var slackMessage = postSlackMessageWithFallback_(
      getConfiguredAccessReviewChannels_().concat([
        requestRecord._SOURCE_CHANNEL_ID,
        requestRecord._SOURCE_CHANNEL_NAME
      ]),
      'New GMP access request ' + requestRecord.REQUEST_ID,
      buildAccessRequestBlocks_(requestRecord, rowNumber)
    );
    return {
      threadRef: buildSlackThreadRef_(slackMessage.channel, slackMessage.ts),
      error: ''
    };
  } catch (err) {
    Logger.log('Initial access notification failed: ' + err.message);
    return {
      threadRef: '',
      error: err.message
    };
  }
}

function tryPostInitialIncidentNotification_(incidentRecord, rowNumber) {
  try {
    var slackMessage = postSlackMessageWithFallback_(
      getConfiguredSupportChannels_().concat([
        incidentRecord._SOURCE_CHANNEL_ID,
        incidentRecord._SOURCE_CHANNEL_NAME
      ]),
      'New GMP issue ' + incidentRecord.INCIDENT_ID,
      buildIncidentBlocks_(incidentRecord, rowNumber)
    );
    return {
      threadRef: buildSlackThreadRef_(slackMessage.channel, slackMessage.ts),
      error: ''
    };
  } catch (err) {
    Logger.log('Initial incident notification failed: ' + err.message);
    return {
      threadRef: '',
      error: err.message
    };
  }
}

function buildSlackInteractionKey_(interaction, kind) {
  var viewId = normalizeText_(interaction && interaction.view && interaction.view.id);
  if (viewId) {
    return kind + '|view|' + viewId;
  }

  var userId = normalizeText_(interaction && interaction.user && interaction.user.id);
  var callbackId = normalizeText_(interaction && interaction.view && interaction.view.callback_id);
  var stateSignature = '';
  try {
    stateSignature = JSON.stringify(interaction && interaction.view && interaction.view.state && interaction.view.state.values || {});
  } catch (err) {
    stateSignature = '';
  }
  return kind + '|fallback|' + hashSecurityValue_([userId, callbackId, stateSignature].join('|'));
}

function reserveSecuritySubmissionKey_(submissionKey, recordId) {
  if (!normalizeText_(submissionKey)) {
    return true;
  }

  var props = PropertiesService.getScriptProperties();
  var rawCache = props.getProperty('SECURITY_SUBMISSION_CACHE');
  var cache = {};
  if (rawCache) {
    try {
      cache = JSON.parse(rawCache) || {};
    } catch (err) {
      cache = {};
    }
  }

  pruneSecuritySubmissionCache_(cache);

  var cacheKey = hashSecurityValue_(submissionKey);
  if (cache[cacheKey]) {
    return false;
  }

  cache[cacheKey] = {
    recordId: normalizeText_(recordId),
    at: Date.now()
  };
  props.setProperty('SECURITY_SUBMISSION_CACHE', JSON.stringify(cache));
  return true;
}

function pruneSecuritySubmissionCache_(cache) {
  var keys = Object.keys(cache || {});
  var cutoff = Date.now() - (14 * 24 * 60 * 60 * 1000);
  for (var i = 0; i < keys.length; i++) {
    var entry = cache[keys[i]];
    if (!entry || !entry.at || entry.at < cutoff) {
      delete cache[keys[i]];
    }
  }

  keys = Object.keys(cache || {});
  if (keys.length <= 500) {
    return;
  }

  keys.sort(function(left, right) {
    return (cache[right] && cache[right].at || 0) - (cache[left] && cache[left].at || 0);
  });

  for (var j = 500; j < keys.length; j++) {
    delete cache[keys[j]];
  }
}

function hashSecurityValue_(value) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    normalizeText_(value),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, '');
}

function getStoredInitialThreadRef_(entityPrefix, entityId) {
  if (!normalizeText_(entityId)) {
    return '';
  }
  return normalizeText_(
    PropertiesService.getScriptProperties().getProperty('SECURITY_THREAD_' + entityPrefix + '_' + entityId)
  );
}

function storeInitialThreadRef_(entityPrefix, entityId, threadRef) {
  if (!normalizeText_(entityId) || !normalizeText_(threadRef)) {
    return;
  }
  PropertiesService.getScriptProperties().setProperty(
    'SECURITY_THREAD_' + entityPrefix + '_' + entityId,
    normalizeText_(threadRef)
  );
}

function queueSecurityIntakeJob_(kind, record) {
  var props = PropertiesService.getScriptProperties();
  var queueKey = 'PENDING_SECURITY_JOB_' + kind.toUpperCase() + '_' + Date.now() + '_' + Utilities.getUuid().slice(0, 8);
  props.setProperty(queueKey, JSON.stringify({
    kind: kind,
    record: record,
    createdAt: new Date().toISOString()
  }));
}

function processPendingSecurityIntake_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return;
  }

  try {
    ensureThreeTabWorkbookReady_();

    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    var keys = Object.keys(allProps)
      .filter(function(key) { return key.indexOf('PENDING_SECURITY_JOB_') === 0; })
      .sort();

    for (var i = 0; i < keys.length; i++) {
      try {
        processQueuedSecurityJob_(keys[i], allProps[keys[i]]);
      } catch (err) {
        Logger.log('Queued security job failed: ' + keys[i] + ' | ' + err.message);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function processQueuedSecurityJob_(propertyKey, rawValue) {
  var props = PropertiesService.getScriptProperties();
  var job = JSON.parse(rawValue);
  if (!job || !job.kind || !job.record) {
    props.deleteProperty(propertyKey);
    return;
  }

  if (job.kind === 'access') {
    processQueuedAccessJob_(job.record);
    props.deleteProperty(propertyKey);
    return;
  }

  if (job.kind === 'incident') {
    processQueuedIncidentJob_(job.record);
    props.deleteProperty(propertyKey);
    return;
  }

  props.deleteProperty(propertyKey);
}

function processQueuedAccessJob_(rawRecord) {
  var requestRecord = hydrateQueuedRecord_(rawRecord);
  var existingRow = findRequestRowById_(requestRecord.REQUEST_ID);
  var rowNumber = upsertQueuedAccessRequestRecord_(requestRecord, existingRow);
  requestRecord.SLACK_THREAD = normalizeText_(requestRecord.SLACK_THREAD) ||
    normalizeText_(existingRow && existingRow.data && existingRow.data.SLACK_THREAD) ||
    getStoredInitialThreadRef_('REQ', requestRecord.REQUEST_ID);

  var notificationResult = requestRecord.SLACK_THREAD ?
    { threadRef: requestRecord.SLACK_THREAD, error: '' } :
    tryPostInitialAccessNotification_(requestRecord, rowNumber);

  requestRecord.SLACK_THREAD = notificationResult.threadRef || '';
  if (requestRecord.SLACK_THREAD) {
    storeInitialThreadRef_('REQ', requestRecord.REQUEST_ID, requestRecord.SLACK_THREAD);
  }
  requestRecord.NEXT_SLACK_TRIGGER = notificationResult.error ?
    'Retry Slack request notification' :
    'Wait for approval update';
  if (notificationResult.error) {
    requestRecord.NOTES = appendUniqueNoteLine_(
      requestRecord.NOTES,
      'Initial Slack notification failed: ' + notificationResult.error
    );
  }
  updateQueuedRequestRecord_(rowNumber, requestRecord, existingRow);
  sendHeartbeat('Slack Access Intake', requestRecord.REQUEST_ID);
}

function processQueuedIncidentJob_(rawRecord) {
  var incidentRecord = hydrateQueuedRecord_(rawRecord);
  var existingRow = findIncidentRowById_(incidentRecord.INCIDENT_ID);
  var rowNumber = upsertQueuedIncidentRecord_(incidentRecord, existingRow);
  incidentRecord.SLACK_THREAD = normalizeText_(incidentRecord.SLACK_THREAD) ||
    normalizeText_(existingRow && existingRow.data && existingRow.data.SLACK_THREAD) ||
    getStoredInitialThreadRef_('INC', incidentRecord.INCIDENT_ID);

  var notificationResult = incidentRecord.SLACK_THREAD ?
    { threadRef: incidentRecord.SLACK_THREAD, error: '' } :
    tryPostInitialIncidentNotification_(incidentRecord, rowNumber);

  incidentRecord.SLACK_THREAD = notificationResult.threadRef || '';
  if (incidentRecord.SLACK_THREAD) {
    storeInitialThreadRef_('INC', incidentRecord.INCIDENT_ID, incidentRecord.SLACK_THREAD);
  }
  incidentRecord.NEXT_SLACK_TRIGGER = notificationResult.error ?
    'Retry Slack incident notification' :
    'Wait for status update';
  if (notificationResult.error) {
    incidentRecord.NOTES = appendUniqueNoteLine_(
      incidentRecord.NOTES,
      'Initial Slack notification failed: ' + notificationResult.error
    );
  }
  updateQueuedIncidentRecord_(rowNumber, incidentRecord, existingRow);
  sendHeartbeat('Slack Incident Intake', incidentRecord.INCIDENT_ID);
}

function sanitizeRecordForQueue_(record) {
  return JSON.parse(JSON.stringify(record));
}

function hydrateQueuedRecord_(rawRecord) {
  var record = {};
  var keys = Object.keys(rawRecord);
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    record[key] = reviveQueuedValue_(rawRecord[key]);
  }
  return record;
}

function reviveQueuedValue_(value) {
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(value)) {
    var parsed = new Date(value);
    if (!isNaN(parsed)) {
      return parsed;
    }
  }
  return value;
}

function upsertQueuedAccessRequestRecord_(requestRecord, existingRow) {
  if (existingRow) {
    updateQueuedRequestRecord_(existingRow.rowNumber, requestRecord, existingRow);
    return existingRow.rowNumber;
  }
  return appendAccessRequestRecord_(requestRecord);
}

function updateQueuedRequestRecord_(rowNumber, requestRecord, existingRow) {
  updateRequestRow_(
    rowNumber,
    mergeQueuedRequestRecord_(requestRecord, existingRow && existingRow.data),
    existingRow && existingRow.sheetName
  );
}

function mergeQueuedRequestRecord_(requestRecord, existingData) {
  var merged = shallowCloneObject_(existingData || {});
  var incoming = shallowCloneObject_(requestRecord || {});
  var keys = Object.keys(incoming);
  for (var i = 0; i < keys.length; i++) {
    merged[keys[i]] = incoming[keys[i]];
  }

  merged.SLACK_THREAD = chooseNonBlankText_(incoming.SLACK_THREAD, merged.SLACK_THREAD);
  merged.NEXT_SLACK_TRIGGER = chooseNonBlankText_(incoming.NEXT_SLACK_TRIGGER, merged.NEXT_SLACK_TRIGGER);
  merged.IT_OWNER = chooseNonBlankText_(incoming.IT_OWNER, merged.IT_OWNER);
  merged.NOTES = chooseLongerText_(incoming.NOTES, merged.NOTES);
  merged.APPROVAL = chooseNonBlankText_(incoming.APPROVAL, merged.APPROVAL);
  merged.PROVISIONING = chooseNonBlankText_(incoming.PROVISIONING, merged.PROVISIONING);
  merged.ACCESS_ID = chooseNonBlankText_(incoming.ACCESS_ID, merged.ACCESS_ID);
  return merged;
}

function upsertQueuedIncidentRecord_(incidentRecord, existingRow) {
  if (existingRow) {
    updateQueuedIncidentRecord_(existingRow.rowNumber, incidentRecord, existingRow);
    return existingRow.rowNumber;
  }
  return appendIncidentRecord_(incidentRecord);
}

function updateQueuedIncidentRecord_(rowNumber, incidentRecord, existingRow) {
  updateIncidentRow_(rowNumber, mergeQueuedIncidentRecord_(incidentRecord, existingRow && existingRow.data));
}

function mergeQueuedIncidentRecord_(incidentRecord, existingData) {
  var merged = shallowCloneObject_(existingData || {});
  var incoming = shallowCloneObject_(incidentRecord || {});
  var keys = Object.keys(incoming);
  for (var i = 0; i < keys.length; i++) {
    merged[keys[i]] = incoming[keys[i]];
  }

  merged.SLACK_THREAD = chooseNonBlankText_(incoming.SLACK_THREAD, merged.SLACK_THREAD);
  merged.NEXT_SLACK_TRIGGER = chooseNonBlankText_(incoming.NEXT_SLACK_TRIGGER, merged.NEXT_SLACK_TRIGGER);
  merged.ASSIGNED_TO = chooseNonBlankText_(incoming.ASSIGNED_TO, merged.ASSIGNED_TO);
  merged.RESOLUTION = chooseNonBlankText_(incoming.RESOLUTION, merged.RESOLUTION);
  merged.NOTES = chooseLongerText_(incoming.NOTES, merged.NOTES);
  return merged;
}

function shallowCloneObject_(source) {
  var clone = {};
  var keys = Object.keys(source || {});
  for (var i = 0; i < keys.length; i++) {
    clone[keys[i]] = source[keys[i]];
  }
  return clone;
}

function appendUniqueNoteLine_(existingValue, noteLine) {
  var current = normalizeText_(existingValue);
  var note = normalizeText_(noteLine);
  if (!note) {
    return current;
  }
  if (!current) {
    return note;
  }
  if (current.indexOf(note) !== -1) {
    return current;
  }
  return current + '\n' + note;
}


function deriveRequestPriority_(requestType, gmpImpact) {
  if (normalizeText_(requestType) === 'Remove Access') {
    return 'High';
  }
  return normalizeText_(gmpImpact) === 'Yes' ? 'High' : 'Medium';
}

function defaultRequestSlaDue_(priority) {
  var now = new Date();
  if (normalizeText_(priority) === 'High') {
    return addHours_(now, 4);
  }
  if (normalizeText_(priority) === 'Medium') {
    return addDays_(now, 1);
  }
  return addDays_(now, 3);
}

function slackUserReference_(user) {
  var displayName = normalizeText_(user && (user.real_name || user.name || user.username));
  if (displayName) {
    return displayName;
  }
  return resolveSlackUserDisplayName_(user && user.id);
}

function slackAck_() {
  return ContentService.createTextOutput('');
}

function slackAckText_(text) {
  return ContentService
    .createTextOutput(JSON.stringify({ response_type: 'ephemeral', text: text }))
    .setMimeType(ContentService.MimeType.JSON);
}

function slackClearModal_() {
  return ContentService
    .createTextOutput(JSON.stringify({ response_action: 'clear' }))
    .setMimeType(ContentService.MimeType.JSON);
}
