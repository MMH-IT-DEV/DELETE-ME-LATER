  /**
   * /it-alert Slash Command Handler
   * Systems Health Monitor — Web App Entry Point
   *
   * Slash command: /it-alert
   * Opens a Slack modal (Layout B) where any team member can report a system failure.
   * On submit: posts to #it-support + logs to "Incident Reports" tab.
   *
   * SETUP (one-time):
   * 1. Script Properties > add:
   *      SLACK_BOT_TOKEN   = xoxb-your-bot-token
   *      SLACK_WEBHOOK_URL = your-incoming-webhook (already set)
   * 2. Deploy as Web App: Execute as "Me", Access "Anyone"
   * 3. Copy the Web App URL into:
   *      Slack App > Slash Commands > /it-alert > Request URL
   *      Slack App > Interactivity & Shortcuts > Request URL
   * 4. Run setupItAlertTrigger() once to keep GAS warm (prevents cold-start timeouts)
   *
   * ARCHITECTURE:
   * - /it-alert fires → doPost → opens Slack modal via views.open (bot token)
   * - User fills modal → Slack sends interaction → doPost → posts + logs
   * - Keepalive trigger fires every 5 min to keep GAS warm
   */

  var ALERT_CONFIG = {
    BOT_TOKEN_PROP:  'SLACK_BOT_TOKEN',
    WEBHOOK_PROP:    'SLACK_WEBHOOK_URL',
    CHANNEL:         '#it-support',
    BOT_NAME:        'IT Governance Bot',
    BOT_ICON:        ':shield:',
    MODAL_CALLBACK:  'it_alert_modal',
    INCIDENT_TAB:    'Incident Reports'
  };

  var INCIDENT_COLS = {
    TIMESTAMP:    1,   // A  When submitted
    SYSTEM:       2,   // B  Automation / system name
    SEVERITY:     3,   // C  Critical / High / Low
    ONGOING:      4,   // D  Yes / No
    TIME_NOTICED: 5,   // E  When first noticed (free text)
    DESCRIPTION:  6,   // F  What happened
    REPORTED_BY:  7,   // G  Reporter name
    STATUS:       8,   // H  Open / In Progress / Resolved
    ASSIGNED_TO:  9,   // I  Who is handling it
    RESOLUTION:   10   // J  Resolution notes
  };


  // ─── WEB APP ENTRY POINT ────────────────────────────────────────────────────

  /**
   * Handles both slash command POSTs and Slack interaction POSTs.
   * Slack sends slash commands as form-encoded params.
   * Slack sends interactions as form-encoded with a JSON `payload` param.
   */
  function routeItAlertPost_(e) {
    try {
      // Interaction payload (button clicks, modal submissions)
      if (e.parameter && e.parameter.payload) {
        var interaction = JSON.parse(e.parameter.payload);
        return handleInteraction_(interaction);
      }

      // Slash command
      if (e.parameter && e.parameter.command === '/it-alert') {
        return handleSlashCommand_(e.parameter);
      }

      // JSON body POST — heartbeat or error from a script
      if (e.postData && e.postData.contents) {
        try {
          var jsonPayload = JSON.parse(e.postData.contents);
          if (jsonPayload.status === 'success' || jsonPayload.status === 'heartbeat') {
            return handleHeartbeat(jsonPayload);
          } else if (jsonPayload.error) {
            return handleError(jsonPayload);
          }
        } catch (parseErr) { /* not JSON — fall through */ }
      }

      // Unknown — ack anyway
      return _ack();

    } catch (err) {
      Logger.log('doPost error: ' + err + '\n' + err.stack);
      return _ack();
    }
  }

  /** GET handler — lets Slack verify the endpoint */
  function routeItAlertGet_(e) {
    return ContentService.createTextOutput('IT Alert endpoint active.');
  }


  // ─── SLASH COMMAND ──────────────────────────────────────────────────────────

  function handleSlashCommand_(params) {
    var triggerId = params.trigger_id;
    var botToken  = PropertiesService.getScriptProperties().getProperty(ALERT_CONFIG.BOT_TOKEN_PROP);

    if (!botToken) {
      Logger.log('/it-alert: SLACK_BOT_TOKEN not set in Script Properties');
      return _ackText(':warning: Bot token not configured. Contact IT admin.');
    }

    if (!triggerId) {
      Logger.log('/it-alert: No trigger_id received');
      return _ackText(':warning: Could not open form. Please try again.');
    }

    // Open the modal — must happen within 3 seconds of slash command
    // Keep GAS warm (setupItAlertTrigger) to avoid cold-start timeout
    var result = openItAlertModal_(triggerId, botToken);

    if (!result.ok) {
      Logger.log('/it-alert modal open failed: ' + JSON.stringify(result));
      return _ackText(':warning: Could not open form (' + (result.error || 'unknown') + '). Try again.');
    }

    // Return empty 200 — Slack requires this to not show an error to the user
    return _ack();
  }


  // ─── MODAL OPEN ─────────────────────────────────────────────────────────────

  function openItAlertModal_(triggerId, botToken) {
    // Pull automation names from System Registry for the dropdown
    var options = getAutomationOptions_();

    var modal = {
      trigger_id: triggerId,
      view: {
        type: 'modal',
        callback_id: ALERT_CONFIG.MODAL_CALLBACK,
        title:  { type: 'plain_text', text: 'Report System Failure' },
        submit: { type: 'plain_text', text: '🚨 Report' },
        close:  { type: 'plain_text', text: 'Cancel' },
        blocks: [
          // Header context
          {
            type: 'context',
            elements: [{ type: 'mrkdwn', text: ':red_circle:  Your alert will be posted to ' + ALERT_CONFIG.CHANNEL + ' and logged in the Health Monitor.' }]
          },
          { type: 'divider' },

          // Automation dropdown
          {
            type: 'input',
            block_id: 'blk_automation',
            label: { type: 'plain_text', text: 'Automation / System' },
            element: {
              type: 'static_select',
              action_id: 'automation',
              placeholder: { type: 'plain_text', text: 'Select the system that broke...' },
              options: options
            }
          },

          // Severity radio buttons (side by side via section would be ideal but Slack doesn't allow it in modals — use input block)
          {
            type: 'input',
            block_id: 'blk_severity',
            label: { type: 'plain_text', text: 'Severity' },
            element: {
              type: 'radio_buttons',
              action_id: 'severity',
              options: [
                { text: { type: 'plain_text', text: '🔴  Critical — system fully down, immediate action needed' }, value: 'Critical' },
                { text: { type: 'plain_text', text: '🟠  High — major degradation, urgent but functional' },      value: 'High'     },
                { text: { type: 'plain_text', text: '🟡  Low — minor issue, can wait until business hours' },     value: 'Low'      }
              ]
            }
          },

          // Description
          {
            type: 'input',
            block_id: 'blk_description',
            label: { type: 'plain_text', text: 'What did you observe?' },
            element: {
              type: 'plain_text_input',
              action_id: 'description',
              multiline: true,
              min_length: 10,
              placeholder: { type: 'plain_text', text: 'e.g. No Slack alerts received since 9 AM. Checked the sheet — heartbeat log is empty for today.' }
            }
          },

          // Still ongoing checkbox
          {
            type: 'input',
            block_id: 'blk_ongoing',
            label: { type: 'plain_text', text: 'Status' },
            optional: true,
            element: {
              type: 'checkboxes',
              action_id: 'ongoing',
              options: [
                { text: { type: 'plain_text', text: 'Still happening (issue is ongoing)' }, value: 'yes' }
              ]
            }
          },

          // When noticed
          {
            type: 'input',
            block_id: 'blk_time',
            label: { type: 'plain_text', text: 'When did you first notice this?' },
            optional: true,
            element: {
              type: 'plain_text_input',
              action_id: 'time_noticed',
              placeholder: { type: 'plain_text', text: 'e.g. 10:30 AM today, yesterday afternoon' }
            }
          }
        ]
      }
    };

    try {
      var resp = UrlFetchApp.fetch('https://slack.com/api/views.open', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + botToken },
        payload: JSON.stringify(modal),
        muteHttpExceptions: true
      });
      return JSON.parse(resp.getContentText());
    } catch (err) {
      Logger.log('openItAlertModal_ error: ' + err);
      return { ok: false, error: err.message };
    }
  }

  /**
   * Reads automation names from System Registry tab.
   * Falls back to a hardcoded list if the sheet is unavailable.
   */
  function getAutomationOptions_() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var reg = ss.getSheetByName(HB_CONFIG.REGISTRY_TAB);
      if (reg && reg.getLastRow() > 1) {
        var names = reg.getRange(2, REG_COLS.NAME, reg.getLastRow() - 1, 1).getValues();
        var opts = [];
        for (var i = 0; i < names.length; i++) {
          var n = sv(names[i][0]);
          if (n) {
            opts.push({ text: { type: 'plain_text', text: n }, value: n });
          }
        }
        if (opts.length > 0) {
          opts.push({ text: { type: 'plain_text', text: 'Other (describe below)' }, value: 'Other' });
          return opts;
        }
      }
    } catch (e) {
      Logger.log('getAutomationOptions_ fallback: ' + e);
    }

    // Fallback list
    return [
      { text: { type: 'plain_text', text: 'Health Monitor (Daily Check)' }, value: 'Health Monitor' },
      { text: { type: 'plain_text', text: 'IT Security Tracker' },          value: 'IT Security Tracker' },
      { text: { type: 'plain_text', text: 'Katana PO Sync' },               value: 'Katana PO Sync' },
      { text: { type: 'plain_text', text: 'FedEx Dispute Bot' },            value: 'FedEx Dispute Bot' },
      { text: { type: 'plain_text', text: 'Shopify Flows' },                value: 'Shopify Flows' },
      { text: { type: 'plain_text', text: 'Other (describe below)' },       value: 'Other' }
    ];
  }


  // ─── INTERACTION HANDLER ────────────────────────────────────────────────────

  function handleInteraction_(interaction) {
    // Modal submission
    if (interaction.type === 'view_submission' &&
        interaction.view && interaction.view.callback_id === ALERT_CONFIG.MODAL_CALLBACK) {

      var data = parseModalSubmission_(interaction);
      postAlertToSlack_(data);
      logAlertToSheet_(data);

      // Return {"response_action":"clear"} to close the modal
      return ContentService
        .createTextOutput(JSON.stringify({ response_action: 'clear' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return _ack();
  }


  // ─── PARSE MODAL SUBMISSION ─────────────────────────────────────────────────

  function parseModalSubmission_(interaction) {
    var v   = interaction.view.state.values;
    var usr = interaction.user;

    var automation  = v.blk_automation.automation.selected_option
                        ? v.blk_automation.automation.selected_option.value : 'Unknown';
    var severity    = v.blk_severity.severity.selected_option
                        ? v.blk_severity.severity.selected_option.value : 'Unknown';
    var description = v.blk_description.description.value || '';
    var ongoingOpts = v.blk_ongoing && v.blk_ongoing.ongoing
                        ? v.blk_ongoing.ongoing.selected_options : [];
    var ongoing     = ongoingOpts.length > 0;
    var timeNoticed = v.blk_time && v.blk_time.time_noticed
                        ? v.blk_time.time_noticed.value || '' : '';

    var reporter    = usr.real_name || usr.name || usr.username || 'Unknown';
    var reporterId  = usr.id || '';

    var tz = Session.getScriptTimeZone();
    var ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    return {
      automation:  automation,
      severity:    severity,
      description: description,
      ongoing:     ongoing,
      timeNoticed: timeNoticed || 'Not specified',
      reporter:    reporter,
      reporterId:  reporterId,
      timestamp:   ts
    };
  }


  // ─── POST TO SLACK ──────────────────────────────────────────────────────────

  function postAlertToSlack_(d) {
    var url = PropertiesService.getScriptProperties().getProperty(ALERT_CONFIG.WEBHOOK_PROP);
    if (!url) { Logger.log('postAlertToSlack_: no webhook URL'); return; }

    var sevEmoji = d.severity === 'Critical' ? ':rotating_light:' :
                  d.severity === 'High'     ? ':warning:' : ':information_source:';
    var ongoingText = d.ongoing ? ':red_circle: Still ongoing' : ':white_check_mark: May be resolved';

    var blocks = [
      { type: 'section', text: { type: 'mrkdwn',
        text: sevEmoji + '  *IT Alert — ' + d.severity + ' Severity*' }
      },
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn',
        text: '*System:*  `' + d.automation + '`\n' +
              '*Severity:*  `' + d.severity + '`\n' +
              '*Status:*  ' + ongoingText
      }},
      { type: 'section', text: { type: 'mrkdwn',
        text: '*What happened:*\n' + d.description
      }},
      { type: 'section', text: { type: 'mrkdwn',
        text: '*First noticed:*  ' + d.timeNoticed + '\n' +
              '*Reported by:*  <@' + d.reporterId + '> (' + d.reporter + ')'
      }},
      { type: 'context', elements: [{ type: 'mrkdwn',
        text: d.timestamp + '  |  /it-alert  |  Manual Report'
      }]}
    ];

    try {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          channel:    ALERT_CONFIG.CHANNEL,
          username:   ALERT_CONFIG.BOT_NAME,
          icon_emoji: ALERT_CONFIG.BOT_ICON,
          text:       'IT Alert: ' + d.automation + ' [' + d.severity + '] reported by ' + d.reporter,
          blocks:     blocks
        }),
        muteHttpExceptions: true
      });
    } catch (err) {
      Logger.log('postAlertToSlack_ error: ' + err);
    }
  }


  // ─── LOG TO INCIDENT REPORTS TAB ────────────────────────────────────────────

  /**
   * Logs the alert to the "Incident Reports" tab.
   * Creates the tab automatically if it doesn't exist.
   */
  function logAlertToSheet_(d) {
    try {
      var ss  = SpreadsheetApp.getActiveSpreadsheet();
      var log = ss.getSheetByName(ALERT_CONFIG.INCIDENT_TAB);

      // Auto-create the tab if it doesn't exist yet
      if (!log) {
        log = setupIncidentReportsTab_(ss);
      }

      var row = new Array(10).fill('');
      row[INCIDENT_COLS.TIMESTAMP    - 1] = d.timestamp;
      row[INCIDENT_COLS.SYSTEM       - 1] = d.automation;
      row[INCIDENT_COLS.SEVERITY     - 1] = d.severity;
      row[INCIDENT_COLS.ONGOING      - 1] = d.ongoing ? 'Yes' : 'No';
      row[INCIDENT_COLS.TIME_NOTICED - 1] = d.timeNoticed;
      row[INCIDENT_COLS.DESCRIPTION  - 1] = d.description;
      row[INCIDENT_COLS.REPORTED_BY  - 1] = d.reporter;
      row[INCIDENT_COLS.STATUS       - 1] = 'Open';

      log.appendRow(row);

      // Apply severity color to the new row
      var newRow  = log.getLastRow();
      var sevCell = log.getRange(newRow, INCIDENT_COLS.SEVERITY);
      applyIncidentSeverityColor_(sevCell, d.severity);

      SpreadsheetApp.flush();
      Logger.log('IT Alert logged to Incident Reports: ' + d.automation + ' [' + d.severity + '] by ' + d.reporter);
    } catch (err) {
      Logger.log('logAlertToSheet_ error: ' + err);
    }
  }


  // ─── TAB SETUP ───────────────────────────────────────────────────────────────

  /**
   * Creates and formats the "Incident Reports" tab.
   * Called automatically on first alert, or manually via setupIncidentReportsTab().
   */
  function setupIncidentReportsTab_(ss) {
    var s = ss || SpreadsheetApp.getActiveSpreadsheet();
    var sheet = s.getSheetByName(ALERT_CONFIG.INCIDENT_TAB);

    if (!sheet) {
      sheet = s.insertSheet(ALERT_CONFIG.INCIDENT_TAB);
    } else {
      // Only re-style, don't wipe data
      sheet.getRange(1, 1, 1, 13).clearContent();
    }

    // ── Header row ──
    var headers = [
      'Timestamp', 'System', 'Severity', 'Ongoing', 'Time Noticed',
      'Description', 'Reported By', 'Status', 'Assigned To', 'Resolution Notes'
    ];

    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setValues([headers]);
    hRange.setBackground('#263238');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    hRange.setFontSize(10);
    hRange.setWrap(false);

    // ── Freeze header ──
    sheet.setFrozenRows(1);

    // ── Column widths ──
    //    A:Timestamp B:System  C:Severity D:Ongoing E:TimeNoticed F:Description G:Reporter H:Status I:AssignedTo J:Resolution
    var widths = [160, 180, 90, 70, 130, 340, 130, 110, 130, 260];
    for (var i = 0; i < widths.length; i++) {
      sheet.setColumnWidth(i + 1, widths[i]);
    }

    // ── Status dropdown (col H) — rows 2–500 ──
    var statusRange = sheet.getRange(2, INCIDENT_COLS.STATUS, 499, 1);
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Open', 'In Progress', 'Resolved'], true)
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusRule);

    // ── Severity conditional formatting (col C) ──
    var sevRange = sheet.getRange(2, INCIDENT_COLS.SEVERITY, 499, 1);
    var rules = [];

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Critical')
      .setBackground('#FFCDD2').setFontColor('#B71C1C')
      .setRanges([sevRange]).build());

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('High')
      .setBackground('#FFE0B2').setFontColor('#E65100')
      .setRanges([sevRange]).build());

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Low')
      .setBackground('#FFF9C4').setFontColor('#F57F17')
      .setRanges([sevRange]).build());

    sheet.setConditionalFormatRules(rules);

    // ── Status conditional formatting (col H) ──
    var statusFmtRange = sheet.getRange(2, INCIDENT_COLS.STATUS, 499, 1);
    var statusRules = [];

    statusRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Open')
      .setBackground('#FFCDD2').setFontColor('#B71C1C')
      .setRanges([statusFmtRange]).build());

    statusRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('In Progress')
      .setBackground('#FFF9C4').setFontColor('#F57F17')
      .setRanges([statusFmtRange]).build());

    statusRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Resolved')
      .setBackground('#C8E6C9').setFontColor('#1B5E20')
      .setRanges([statusFmtRange]).build());

    // Merge both rule sets
    var allRules = sheet.getConditionalFormatRules().concat(statusRules);
    sheet.setConditionalFormatRules(allRules);

    SpreadsheetApp.flush();
    Logger.log('✓ Incident Reports tab created/updated');
    return sheet;
  }

  /** Public wrapper — run manually from the Apps Script editor to create/reset the tab */
  function setupIncidentReportsTab() {
    setupIncidentReportsTab_(null);
    try {
      SpreadsheetApp.getUi().alert('Done', '"Incident Reports" tab is ready.', SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
  }

  /** Applies severity background directly to a cell (used when appending rows) */
  function applyIncidentSeverityColor_(cell, severity) {
    var colors = { Critical: '#FFCDD2', High: '#FFE0B2', Low: '#FFF9C4' };
    var fg     = { Critical: '#B71C1C', High: '#E65100', Low: '#F57F17' };
    if (colors[severity]) {
      cell.setBackground(colors[severity]).setFontColor(fg[severity]);
    }
  }


  // ─── KEEPALIVE (prevents cold-start timeouts) ────────────────────────────────

  /**
   * No-op function triggered every 5 minutes to keep GAS warm.
   * A warm instance responds in < 1 second — cold instances take 5-10 seconds,
   * causing Slack to time out before views.open can be called.
   *
   * Run setupItAlertTrigger() once from the Apps Script editor to activate.
   */
  function keepAliveWarmup_() {
    // Intentionally empty — just wakes the GAS runtime
    Logger.log('keepAlive: ' + new Date().toISOString());
  }

  /**
   * Sets up the 5-minute keepalive trigger.
   * Run this manually once after deploying.
   * Safe to run multiple times — removes existing keepalive triggers first.
   */
  function setupItAlertTrigger() {
    // Remove any existing keepalive triggers
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'keepAliveWarmup_') {
        ScriptApp.deleteTrigger(t);
      }
    });

    // Create new 5-minute keepalive
    ScriptApp.newTrigger('keepAliveWarmup_')
      .timeBased()
      .everyMinutes(5)
      .create();

    Logger.log('✓ Keepalive trigger set: every 5 minutes (keepAliveWarmup_)');
    try {
      SpreadsheetApp.getUi().alert('Done',
        'Keepalive trigger set — GAS will stay warm.\n\n' +
        'Next: deploy as Web App and configure Slack.',
        SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      Logger.log('Keepalive setup complete.');
    }
  }


  // ─── TEST FUNCTION ───────────────────────────────────────────────────────────

  /**
   * Simulates a form submission to test posting + logging without Slack interaction.
   * Run from Apps Script editor.
   */
  function testItAlertFlow() {
    var testData = {
      automation:  'Health Monitor',
      severity:    'High',
      description: 'TEST: No Slack summary received by 9 AM. Heartbeat Log is empty.',
      ongoing:     true,
      timeNoticed: '9:15 AM today',
      reporter:    'Test User',
      reporterId:  'U00000000',
      timestamp:   Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };

    postAlertToSlack_(testData);
    logAlertToSheet_(testData);
    Logger.log('✓ Test IT alert sent to ' + ALERT_CONFIG.CHANNEL + ' and logged to "Incident Reports" tab.');
  }


  // ─── HELPERS ────────────────────────────────────────────────────────────────

  /** Return empty HTTP 200 — required by Slack for slash command ack */
  function _ack() {
    return ContentService.createTextOutput('');
  }

  /** Return ephemeral text message (only visible to user who ran the command) */
  function _ackText(text) {
    return ContentService
      .createTextOutput(JSON.stringify({ response_type: 'ephemeral', text: text }))
      .setMimeType(ContentService.MimeType.JSON);
  }
