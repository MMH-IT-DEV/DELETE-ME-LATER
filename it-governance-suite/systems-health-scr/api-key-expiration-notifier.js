/**
 * 2026_SCRIPT-V3-System-Health-Monitor
 * Container-bound script for Health Monitor Spreadsheet
 *
 * MyMagicHealer - January 2026 (updated Feb 2026)
 *
 * FEATURES:
 * - Webhook endpoint for heartbeats/errors from any system
 * - Missing heartbeat detection (systems that stopped running)
 * - Credential expiration alerts
 * - Slack notifications via Script Properties (no hardcoded URLs)
 * - Daily health report
 * - Custom menu in spreadsheet
 *
 * NOTE: This file works alongside health-check.gs in the same project.
 * health-check.gs handles: onOpen menu, triggers, daily checks, cleanup, standardize.
 * This file handles: webhook endpoints, heartbeat/error receiving, standalone checks.
 *
 * SETUP:
 * 1. Project Settings > Script Properties > add:
 *    SLACK_WEBHOOK_URL = your webhook URL
 */

// ============================================
// CONFIGURATION
// ============================================
var CONFIG = {
  // Tab names in spreadsheet
  TABS: {
    REGISTRY: 'System Registry',
    ERROR_LOG: 'Error Log',
    HEARTBEAT_LOG: 'Heartbeat Log',
    FIX_GUIDES: 'Fix Guides'
  },

  // Expected heartbeat intervals (hours) - alert if no heartbeat within this time
  HEARTBEAT_THRESHOLDS: {
    'FLOW': 48,      // Shopify Flows - alert if no activity in 48 hours
    'SCRIPT': 48,    // Google Apps Scripts - 48 hours
    'BOT': 192,      // Windows Bots - 8 days (may run weekly)
    'API': 72,       // API integrations - 72 hours
    'TASK': 48,      // Scheduled tasks - 48 hours
    'MAKE': 48       // Make.com scenarios - 48 hours
  },

  // Days before expiry to send alerts
  EXPIRY_ALERT_DAYS: [30, 14, 7, 3, 1]
};

// ============================================
// SHARED HELPER - Get webhook URL from Script Properties
// ============================================

/**
 * Reads webhook URL from Script Properties instead of hardcoding it.
 * Both this file and health-check.gs use this function.
 */
function getWebhookUrl() {
  return PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || '';
}

// ============================================
// WEBHOOK ENDPOINTS
// ============================================

/**
 * POST endpoint - receives heartbeats and errors from external systems
 *
 * Heartbeat payload:
 * { "system": "Flow Name", "status": "success", "trigger": "Order #123", "details": "optional" }
 *
 * Error payload:
 * { "system": "Flow Name", "error": "Error message", "severity": "critical|error|warning|info", "context": {} }
 */
function doPost(e) {
  try {
    if (typeof routeItAlertPost_ === 'function' &&
        e &&
        e.parameter &&
        (e.parameter.payload || e.parameter.command === '/it-alert')) {
      return routeItAlertPost_(e);
    }

    var payload = JSON.parse(e.postData.contents);

    // Route based on payload type
    if (payload.status === 'success' || payload.status === 'heartbeat') {
      return handleHeartbeat(payload);
    } else {
      return handleError(payload);
    }
  } catch (error) {
    Logger.log('Webhook error: ' + error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET endpoint - health check for the webhook itself
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'healthy',
      service: '2026_SCRIPT-V3-System-Health-Monitor',
      timestamp: new Date().toISOString(),
      message: 'Health Monitor webhook is running'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// HEARTBEAT HANDLING
// ============================================
function handleHeartbeat(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var heartbeatSheet = ss.getSheetByName(CONFIG.TABS.HEARTBEAT_LOG);
  var registrySheet = ss.getSheetByName(CONFIG.TABS.REGISTRY);

  var timestamp = new Date();
  var tz = Session.getScriptTimeZone();
  var standardTs = Utilities.formatDate(timestamp, tz, 'yyyy-MM-dd HH:mm:ss');
  var systemName = payload.system || payload.integration || 'Unknown System';

  // Log to Heartbeat Log (append at bottom)
  var logRow = heartbeatSheet.getLastRow() + 1;
  heartbeatSheet.getRange(logRow, 1, 1, 5).setValues([[
    standardTs,
    systemName,
    'Success',
    payload.trigger || payload.action || payload.event || '',
    payload.details || ''
  ]]);

  // Update Registry - plain text status + color coding
  updateRegistryField(registrySheet, systemName, 'Last Heartbeat', standardTs);
  updateRegistryField(registrySheet, systemName, 'Status', 'Healthy');
  SpreadsheetApp.flush();
  styleStatusByName(registrySheet, systemName);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'Heartbeat logged', system: systemName }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// ERROR HANDLING
// ============================================
function handleError(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errorSheet = ss.getSheetByName(CONFIG.TABS.ERROR_LOG);
  var heartbeatSheet = ss.getSheetByName(CONFIG.TABS.HEARTBEAT_LOG);
  var registrySheet = ss.getSheetByName(CONFIG.TABS.REGISTRY);

  var timestamp = new Date();
  var tz = Session.getScriptTimeZone();
  var standardTs = Utilities.formatDate(timestamp, tz, 'yyyy-MM-dd HH:mm:ss');
  var systemName = payload.system || payload.integration || 'Unknown System';
  var fixGuideId = getFixGuideId(registrySheet, systemName);

  // Map severity (plain text, no emoji)
  var severityMap = {
    'critical': 'Critical',
    'error': 'Error',
    'warning': 'Warning',
    'info': 'Info'
  };
  var severity = severityMap[payload.severity] || 'Error';

  // Log to Error Log (append at bottom)
  if (errorSheet) {
    var logRow = errorSheet.getLastRow() + 1;
    errorSheet.getRange(logRow, 1, 1, 10).setValues([[
      standardTs,
      systemName,
      payload.error || 'Unknown error',
      severity,
      payload.context ? JSON.stringify(payload.context) : '',
      fixGuideId,
      'Open',
      '',
      '',
      ''
    ]]);
  } else if (heartbeatSheet) {
    var hbRow = heartbeatSheet.getLastRow() + 1;
    heartbeatSheet.getRange(hbRow, 1, 1, 5).setValues([[
      standardTs,
      systemName,
      'Failed',
      payload.trigger || payload.action || 'Error',
      payload.error || 'Unknown error'
    ]]);
  }

  // Update Registry - plain text status + color coding
  updateRegistryField(registrySheet, systemName, 'Status', 'Down');
  styleStatusByName(registrySheet, systemName);

  // Send Slack alert
  sendSlackAlert({
    system: systemName,
    error: payload.error,
    severity: payload.severity || 'error',
    context: payload.context,
    fixGuideId: fixGuideId
  });

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'Error logged and alert sent', system: systemName }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// STANDALONE HEALTH CHECKS (callable from menu)
// ============================================

/**
 * Check for systems that haven't sent heartbeats within their threshold
 */
function checkMissingHeartbeats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var registry = ss.getSheetByName(CONFIG.TABS.REGISTRY);
  var data = registry.getDataRange().getValues();
  var headers = data[0];

  var typeCol = headers.indexOf('Type');
  var statusCol = headers.indexOf('Status');
  var heartbeatCol = headers.indexOf('Last Heartbeat');
  var heartbeatMethodCol = headers.indexOf('Heartbeat Method');
  var systemCol = 0;

  var now = new Date();
  var missingHeartbeats = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var systemName = row[systemCol];
    var type = row[typeCol];
    var heartbeatMethod = heartbeatMethodCol >= 0 ? row[heartbeatMethodCol] : '';
    var lastHeartbeat = row[heartbeatCol];
    var heartbeatMethodText = heartbeatMethod ? heartbeatMethod.toString().trim() : '';
    var isMonitored = heartbeatMethodCol >= 0
      ? (heartbeatMethodText !== '' && heartbeatMethodText.toUpperCase() !== 'N/A')
      : !!lastHeartbeat;
    var parsedHeartbeat = null;

    if (lastHeartbeat && lastHeartbeat instanceof Date && !isNaN(lastHeartbeat)) {
      parsedHeartbeat = lastHeartbeat;
    } else if (lastHeartbeat) {
      var parsedDate = new Date(lastHeartbeat.toString().trim());
      if (!isNaN(parsedDate)) {
        parsedHeartbeat = parsedDate;
      }
    }

    // Skip if no system name or heartbeat monitoring is not configured
    if (!systemName || !isMonitored) continue;

    // Get threshold for this type
    var thresholdHours = CONFIG.HEARTBEAT_THRESHOLDS[type] || 48;

    // Check if heartbeat is missing or too old
    if (parsedHeartbeat) {
      var hoursSinceHeartbeat = (now - parsedHeartbeat) / (1000 * 60 * 60);

      if (hoursSinceHeartbeat > thresholdHours) {
        missingHeartbeats.push({
          system: systemName,
          type: type,
          lastSeen: parsedHeartbeat,
          hoursSince: Math.round(hoursSinceHeartbeat),
          threshold: thresholdHours
        });

        // Update status to Degraded (plain text + color)
        registry.getRange(i + 1, statusCol + 1).setValue('Degraded');
        styleRegStatusCell(registry.getRange(i + 1, statusCol + 1), 'Degraded');
      }
    } else {
      // Heartbeat monitoring is configured but no valid heartbeat has been recorded
      missingHeartbeats.push({
        system: systemName,
        type: type,
        lastSeen: 'Never',
        hoursSince: 'N/A',
        threshold: thresholdHours
      });
    }
  }

  // Alert if any missing
  if (missingHeartbeats.length > 0) {
    sendMissingHeartbeatAlert(missingHeartbeats);
  }

  Logger.log('Heartbeat check complete. ' + missingHeartbeats.length + ' systems with missing heartbeats.');
  return missingHeartbeats;
}

/**
 * Check for credentials expiring soon
 */
function checkExpiringCredentials() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var registry = ss.getSheetByName(CONFIG.TABS.REGISTRY);
  var data = registry.getDataRange().getValues();
  var headers = data[0];

  var expiryCol = headers.indexOf('Expiry Date');
  var systemCol = 0;

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var expiringCredentials = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var systemName = row[systemCol];
    var expiryDate = row[expiryCol];

    if (!systemName || !expiryDate || !(expiryDate instanceof Date)) continue;

    var daysUntilExpiry = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));

    // Check if expiring soon
    if (CONFIG.EXPIRY_ALERT_DAYS.indexOf(daysUntilExpiry) >= 0) {
      expiringCredentials.push({
        system: systemName,
        expiryDate: expiryDate,
        daysLeft: daysUntilExpiry
      });
    } else if (daysUntilExpiry < 0) {
      expiringCredentials.push({
        system: systemName,
        expiryDate: expiryDate,
        daysLeft: daysUntilExpiry,
        expired: true
      });
    }
  }

  // Alert for each expiring credential
  for (var c = 0; c < expiringCredentials.length; c++) {
    sendExpiryAlert(expiringCredentials[c]);
  }

  Logger.log('Expiry check complete. ' + expiringCredentials.length + ' credentials expiring/expired.');
  return expiringCredentials;
}

// ============================================
// SLACK NOTIFICATIONS
// ============================================

function sendSlackAlert(payload) {
  var webhookUrl = getWebhookUrl();
  if (!webhookUrl) {
    Logger.log('Slack webhook not configured in Script Properties');
    return;
  }

  var fixSteps = getFixSteps(payload.system);
  var severityEmoji = { 'critical': ':rotating_light:', 'error': ':x:', 'warning': ':warning:', 'info': ':information_source:' };
  var emoji = severityEmoji[payload.severity] || ':x:';

  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var tz = Session.getScriptTimeZone();
  var ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

  var message = {
    channel: '#it-support',
    username: 'IT Health Monitor',
    icon_emoji: ':heartpulse:',
    text: 'System Error - ' + payload.system,
    blocks: [
      { type: 'section', text: { type: 'mrkdwn', text: emoji + '  *System Error Detected*' } },
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn', text: '*System:*  `' + payload.system + '`\n*Severity:*  `' + (payload.severity || 'error').toUpperCase() + '`' } },
      { type: 'section', text: { type: 'mrkdwn', text: '*Error:*\n```' + payload.error + '```' } },
      { type: 'section', text: { type: 'mrkdwn', text: '*Fix Guide:*  `' + (payload.fixGuideId || 'Not found') + '`\n' + fixSteps } },
      { type: 'section', text: { type: 'mrkdwn', text: '*Full Report:*  <https://docs.google.com/spreadsheets/d/' + ssId + '/edit|View Health Tracker>' } },
      { type: 'context', elements: [{ type: 'mrkdwn', text: ts + '  |  Error Alert' }] }
    ]
  };

  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(message),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Slack send error: ' + e);
  }
}

function sendMissingHeartbeatAlert(systems) {
  var webhookUrl = getWebhookUrl();
  if (!webhookUrl || systems.length === 0) return;

  var systemList = '';
  for (var s = 0; s < systems.length; s++) {
    var sys = systems[s];
    var lastSeen = (sys.lastSeen instanceof Date) ? sys.lastSeen.toLocaleString() : sys.lastSeen;
    systemList += ':warning:  `' + sys.system + '` (' + sys.type + ') - Last seen: ' + lastSeen + '\n';
  }

  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var tz = Session.getScriptTimeZone();
  var ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

  var message = {
    channel: '#it-support',
    username: 'IT Health Monitor',
    icon_emoji: ':heartpulse:',
    text: 'Missing Heartbeats - ' + systems.length + ' system(s)',
    blocks: [
      { type: 'section', text: { type: 'mrkdwn', text: ':warning:  *Missing Heartbeats Detected*' } },
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn', text: systemList } },
      { type: 'section', text: { type: 'mrkdwn', text: '*Action Required:*\nCheck if these systems are still running. They may have stopped or encountered silent failures.' } },
      { type: 'section', text: { type: 'mrkdwn', text: '*Full Report:*  <https://docs.google.com/spreadsheets/d/' + ssId + '/edit|View Health Tracker>' } },
      { type: 'context', elements: [{ type: 'mrkdwn', text: ts + '  |  Heartbeat Alert' }] }
    ]
  };

  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(message),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Slack send error: ' + e);
  }
}

function sendExpiryAlert(credential) {
  var webhookUrl = getWebhookUrl();
  if (!webhookUrl) return;

  var emoji = credential.expired ? ':rotating_light:' : (credential.daysLeft <= 7 ? ':warning:' : ':information_source:');
  var status = credential.expired ? 'EXPIRED' : 'expires in ' + credential.daysLeft + ' day' + (credential.daysLeft === 1 ? '' : 's');

  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var tz = Session.getScriptTimeZone();
  var ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var expiryStr = Utilities.formatDate(credential.expiryDate, tz, 'yyyy-MM-dd');

  var message = {
    channel: '#it-support',
    username: 'IT Health Monitor',
    icon_emoji: ':heartpulse:',
    text: 'Credential ' + (credential.expired ? 'Expired' : 'Expiring') + ' - ' + credential.system,
    blocks: [
      { type: 'section', text: { type: 'mrkdwn', text: emoji + '  *Credential ' + (credential.expired ? 'Expired' : 'Expiring Soon') + '*' } },
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn', text: '*System:*  `' + credential.system + '`\n*Status:*  `' + status + '`\n*Expiry Date:*  `' + expiryStr + '`' } },
      { type: 'section', text: { type: 'mrkdwn', text: '*Action Required:*\nRenew or regenerate the credential before it expires to avoid service interruption.' } },
      { type: 'section', text: { type: 'mrkdwn', text: '*Full Report:*  <https://docs.google.com/spreadsheets/d/' + ssId + '/edit|View Health Tracker>' } },
      { type: 'context', elements: [{ type: 'mrkdwn', text: ts + '  |  Expiry Alert' }] }
    ]
  };

  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(message),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Slack send error: ' + e);
  }
}

function testSlackAlert() {
  var webhookUrl = getWebhookUrl();
  if (!webhookUrl) {
    SpreadsheetApp.getUi().alert('Error', 'SLACK_WEBHOOK_URL not set in Script Properties.\n\nGo to: Project Settings > Script Properties > Add:\nKey: SLACK_WEBHOOK_URL\nValue: your webhook URL', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  sendSlackAlert({
    system: 'Test System',
    error: 'This is a test alert from the Health Monitor',
    severity: 'info',
    fixGuideId: 'TEST-001'
  });

  SpreadsheetApp.getUi().alert('Test alert sent to #it-support!');
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function updateRegistryField(sheet, systemName, fieldName, value) {
  if (!sheet) { Logger.log('updateRegistryField: sheet not found for ' + systemName); return; }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var fieldCol = headers.indexOf(fieldName);

  if (fieldCol === -1) return;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase().indexOf(systemName.toLowerCase()) >= 0) {
      sheet.getRange(i + 1, fieldCol + 1).setValue(value);
      break;
    }
  }
}

/**
 * Finds a system in the registry and color-codes its Status cell.
 * Uses styleRegStatusCell from health-check.gs (shared scope).
 */
function styleStatusByName(sheet, systemName) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var statusCol = headers.indexOf('Status');

  if (statusCol === -1) return;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase().indexOf(systemName.toLowerCase()) >= 0) {
      var cell = sheet.getRange(i + 1, statusCol + 1);
      var statusVal = cell.getValue().toString().trim();
      // styleRegStatusCell is defined in health-check.gs (same project scope)
      if (typeof styleRegStatusCell === 'function') {
        styleRegStatusCell(cell, statusVal);
      }
      break;
    }
  }
}

function getFixGuideId(sheet, systemName) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var fixGuideCol = headers.indexOf('Fix Guide ID');

  if (fixGuideCol === -1) return 'Not found';

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase().indexOf(systemName.toLowerCase()) >= 0) {
      return data[i][fixGuideCol] || 'Not set';
    }
  }
  return 'System not in registry';
}

function getFixSteps(systemName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fixSheet = ss.getSheetByName(CONFIG.TABS.FIX_GUIDES);

  if (!fixSheet) return 'Fix guide not found';

  var data = fixSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase().indexOf(systemName.toLowerCase()) >= 0) {
      var steps = [];
      for (var j = 3; j <= 7; j++) {
        if (data[i][j]) steps.push((j - 2) + '. ' + data[i][j]);
      }
      return steps.length > 0 ? steps.join('\n') : 'No steps available';
    }
  }
  return 'No fix guide available for this system';
}

// ============================================
// TEST FUNCTIONS
// ============================================

function testHeartbeatWebhook() {
  var testPayload = {
    postData: {
      contents: JSON.stringify({
        system: 'Shopify Hold - ShipStation Hold Sync',
        status: 'success',
        trigger: 'Test order #12345'
      })
    }
  };

  var result = doPost(testPayload);
  Logger.log('Test heartbeat result: ' + result.getContent());
}

function testErrorWebhook() {
  var testPayload = {
    postData: {
      contents: JSON.stringify({
        system: 'Shopify Hold - ShipStation Hold Sync',
        error: 'Test error - 401 Unauthorized',
        severity: 'error',
        context: { endpoint: 'https://ssapi.shipstation.com/orders/holduntil' }
      })
    }
  };

  var result = doPost(testPayload);
  Logger.log('Test error result: ' + result.getContent());
}
