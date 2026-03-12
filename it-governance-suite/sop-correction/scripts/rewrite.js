/**
 * rewrite.js
 * Rewrites IT-014 and IT-015 into the existing Google Docs (same IDs — no new docs created).
 * body.clear() wipes the content of the SAME document, then rewrites it.
 *
 * Standard template:
 *   Title → Metadata (right-aligned) → 1.Purpose → 2.Scope → 3.Responsibilities
 *   → 4.Definitions → 5.Procedure → 6.References → 7.Attachments → 8.Revision History
 *
 * Design spec (from LOG-006 / QA-012 example SOPs):
 *   Font:  Arial throughout
 *   Title: 26pt, #000000, bold, centered
 *   H3:    14pt, #434343, bold, 6pt spacing after, line-height 1.15
 *   Body:  11pt, #000000, line-height 1.15
 *   Meta:  11pt, right-aligned — label plain, value underlined
 *   Table: 1pt #000000 border, 5pt padding, white bg, header row bold
 *   Link:  #1155CC, underlined
 *   Image placeholder: 10pt, #7F6000, background #FFF2CC, bold italic
 */


// ─── KNOWN URLS ────────────────────────────────────────────────────────────────
// Update these if sheets/scripts are moved or redeployed.

var URLS = {
  // Google Sheets
  SECURITY_TRACKER_SHEET:  'https://docs.google.com/spreadsheets/d/1WBzed_RMtPC3kii-ybJ3lE4Fo1pAqWBt9OJyCSVR-mQ/edit',
  HEALTH_MONITOR_SHEET:    'https://docs.google.com/spreadsheets/d/1jnWtdBPzR7DreihCHQASiN7splmRS7HYJTR77GpfI5w/edit',
  SOP_REGISTRY_SHEET:      'https://docs.google.com/spreadsheets/d/1AUKnAMy1mTshnik3QYm8IwGN1NKPQ7ERX19LxlFlk-Q/edit',

  // Apps Script editors
  SECURITY_TRACKER_SCRIPT: 'https://script.google.com/u/0/home/projects/1Fp0ooeKm028-0XYu5X5CCFrxzILbDEPxwX6Si_I7c3A3GcwmEnANNb4k/edit',
  HEALTH_MONITOR_SCRIPT:   'https://script.google.com/u/0/home/projects/1TBkee_JgNKnHxxeCbWSp3uJyh5Wg5o_GLowbp0k51DP98GjzzDWN4eS5/edit',

  // SOP Google Docs
  IT014_DOC: 'https://docs.google.com/document/d/1yxhYFNNBy9pxtWwjrt8Kgpsq_zTs6hJv-J-xXGMDMEo/edit',
  IT015_DOC: 'https://docs.google.com/document/d/1oGRzFrKKGAxZ2WVl6QhikuKLijb9Uq_JrS88YyApAE8/edit',
  IT008_DOC: 'https://docs.google.com/document/d/1KvqBxBo0bpkqGtFrwPWpYTWdb8TQdnY1TTuqT5iZ6Go/edit',
  IT010_DOC: 'https://docs.google.com/document/d/13PHfqd7Z9Ued0jdjTiTS1A2UTTS3K-h0AMX6ujFkAZg/edit',

  // External
  FEDEX_BOT_GITHUB: 'https://github.com/MMH-IT-DEV/MMH_FEDEX_DISPUTE_BOT',
  SHOPIFY_FLOWS:    'https://admin.shopify.com/store/mymagichealer/apps/flow'
};


// ─── CONSTANTS ────────────────────────────────────────────────────────────────

var ARIAL      = 'Arial';
var H3_COLOR   = '#434343';
var BODY_COLOR = '#000000';
var LINK_COLOR = '#1155CC';


// ─── STYLE HELPERS ────────────────────────────────────────────────────────────

/** Document title: 26pt, bold, centered */
function _title(body, text) {
  var para = body.appendParagraph(text);
  para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  if (text.length > 0) {
    var t = para.editAsText();
    t.setFontFamily(ARIAL);
    t.setFontSize(26);
    t.setForegroundColor(BODY_COLOR);
    t.setBold(true);
  }
  return para;
}

/** H3 section heading: 14pt, #434343, bold */
function _h3(body, text) {
  var para = body.appendParagraph(text);
  para.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  para.setSpacingAfter(6);
  para.setLineSpacing(1.15);
  if (text.length > 0) {
    var t = para.editAsText();
    t.setFontFamily(ARIAL);
    t.setFontSize(14);
    t.setForegroundColor(H3_COLOR);
    t.setBold(true);
  }
  return para;
}

/** Body paragraph: 11pt, #000000 */
function _p(body, text) {
  var para = body.appendParagraph(text);
  para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  para.setLineSpacing(1.15);
  if (text.length > 0) {
    var t = para.editAsText();
    t.setFontFamily(ARIAL);
    t.setFontSize(11);
    t.setForegroundColor(BODY_COLOR);
    t.setBold(false);
  }
  return para;
}

/** Bold body paragraph */
function _pBold(body, text) {
  var para = _p(body, text);
  if (text.length > 0) para.editAsText().setBold(true);
  return para;
}

/**
 * Right-aligned metadata line.
 * Label is plain; value is underlined.
 */
function _meta(body, label, value) {
  var full = label + value;
  var para = body.appendParagraph(full);
  para.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  para.setLineSpacing(1.15);
  var t = para.editAsText();
  t.setFontFamily(ARIAL);
  t.setFontSize(11);
  t.setForegroundColor(BODY_COLOR);
  t.setBold(false);
  if (value.length > 0) t.setUnderline(label.length, full.length - 1, true);
  return para;
}

/**
 * Right-aligned metadata line where the value is also a hyperlink.
 */
function _metaLink(body, label, value, url) {
  var full = label + value;
  var para = body.appendParagraph(full);
  para.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  para.setLineSpacing(1.15);
  var t = para.editAsText();
  t.setFontFamily(ARIAL);
  t.setFontSize(11);
  t.setForegroundColor(BODY_COLOR);
  t.setBold(false);
  if (value.length > 0) {
    var s = label.length;
    var e = full.length - 1;
    t.setUnderline(s, e, true);
    t.setForegroundColor(s, e, LINK_COLOR);
    t.setLinkUrl(s, e, url);
  }
  return para;
}

/** Bullet list item: 11pt */
function _bullet(body, text, level) {
  var item = body.appendListItem(text);
  item.setNestingLevel(level || 0);
  item.setGlyphType(DocumentApp.GlyphType.BULLET);
  item.setLineSpacing(1.15);
  if (text.length > 0) {
    var t = item.editAsText();
    t.setFontFamily(ARIAL);
    t.setFontSize(11);
    t.setForegroundColor(BODY_COLOR);
  }
  return item;
}

/**
 * Table: 1pt black border, 5pt padding, white bg, header row bold.
 */
function _tbl(body, data) {
  var table = body.appendTable(data);
  table.setBorderColor(BODY_COLOR);
  for (var r = 0; r < table.getNumRows(); r++) {
    var row   = table.getRow(r);
    var isHdr = (r === 0);
    for (var c = 0; c < row.getNumCells(); c++) {
      var cell = row.getCell(c);
      cell.setPaddingTop(5);
      cell.setPaddingBottom(5);
      cell.setPaddingLeft(5);
      cell.setPaddingRight(5);
      cell.setBackgroundColor('#ffffff');
      if (cell.getText().length > 0) {
        var ct = cell.editAsText();
        ct.setFontFamily(ARIAL);
        ct.setFontSize(11);
        ct.setForegroundColor(BODY_COLOR);
        if (isHdr) ct.setBold(true);
      }
    }
  }
  return table;
}

/** Empty paragraph spacer */
function _blank(body) {
  var para = body.appendParagraph('');
  para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  return para;
}

/**
 * Body paragraph with one or more inline hyperlinks.
 * links = [{ phrase: 'text to link', url: 'https://...' }, ...]
 * Each phrase must appear exactly once in text.
 */
function _pWithLinks(body, text, links) {
  var para = _p(body, text);
  var t    = para.editAsText();
  for (var i = 0; i < links.length; i++) {
    var phrase = links[i].phrase;
    var url    = links[i].url;
    var idx    = text.indexOf(phrase);
    if (idx === -1) continue;
    var end    = idx + phrase.length - 1;
    t.setLinkUrl(idx, end, url);
    t.setForegroundColor(idx, end, LINK_COLOR);
    t.setUnderline(idx, end, true);
  }
  return para;
}

/**
 * Bullet item with one or more inline hyperlinks.
 * links = [{ phrase: 'text to link', url: 'https://...' }, ...]
 */
function _bulletWithLinks(body, text, links, level) {
  var item = _bullet(body, text, level);
  var t    = item.editAsText();
  for (var i = 0; i < links.length; i++) {
    var phrase = links[i].phrase;
    var url    = links[i].url;
    var idx    = text.indexOf(phrase);
    if (idx === -1) continue;
    var end    = idx + phrase.length - 1;
    t.setLinkUrl(idx, end, url);
    t.setForegroundColor(idx, end, LINK_COLOR);
    t.setUnderline(idx, end, true);
  }
  return item;
}

/**
 * Reference line in the References section — the entire line is linked.
 * Format: "• IT-015: IT Security Tracker Maintenance" where the whole line links to the doc.
 */
function _refLink(body, displayText, url) {
  var item = body.appendListItem(displayText);
  item.setNestingLevel(0);
  item.setGlyphType(DocumentApp.GlyphType.BULLET);
  item.setLineSpacing(1.15);
  if (displayText.length > 0) {
    var t = item.editAsText();
    t.setFontFamily(ARIAL);
    t.setFontSize(11);
    t.setForegroundColor(LINK_COLOR);
    t.setUnderline(true);
    t.setLinkUrl(url);
  }
  return item;
}

/**
 * Image placeholder — amber highlighted block.
 * Marks exactly where a screenshot or diagram should be inserted.
 * To insert: click just before this line in the doc, then Insert → Image.
 */
function _img(body, description) {
  var text = '[ \uD83D\uDCF7 INSERT IMAGE: ' + description + ' ]';
  var para = body.appendParagraph(text);
  para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  para.setLineSpacing(1.15);
  var t = para.editAsText();
  t.setFontFamily(ARIAL);
  t.setFontSize(10);
  t.setForegroundColor('#7F6000');
  t.setBackgroundColor('#FFF2CC');
  t.setBold(true);
  t.setItalic(true);
  return para;
}


// ─── IT-014: Automation Maintenance ──────────────────────────────────────────
// Opens the EXISTING document by ID — body.clear() wipes content, does NOT create a new doc.

function rewriteIT014() {
  var doc  = DocumentApp.openById(SOP_TO_FIX[0].id);
  var body = doc.getBody();
  body.clear();

  // ── Title ──
  _title(body, 'IT-014: Automation Maintenance');
  _blank(body);

  // ── Metadata ──
  _meta(body,     'Department: ',               'Information Technology');
  _meta(body,     'SOP ID: ',                   'IT-014');
  _meta(body,     'Version: ',                  '1.4');
  _meta(body,     'Effective Date: ',           'March 11, 2026');
  _meta(body,     'Review Cycle: ',             'Every 6 months');
  _metaLink(body, 'Systems Health Monitor: ',   'Open spreadsheet \u2192', URLS.HEALTH_MONITOR_SHEET);
  _metaLink(body, 'Systems Health Apps Script: ','Open editor \u2192', URLS.HEALTH_MONITOR_SCRIPT);
  _meta(body,     'Prepared by: ',              '________________________________');
  _meta(body,     'Reviewed by (IT Manager): ', '________________________________');
  _meta(body,     'Approved by (QA Manager): ', '________________________________');
  _blank(body);

  // ── 1. Purpose ──
  _h3(body, '1. Purpose');
  _p(body, 'To define procedures for maintaining, monitoring, and troubleshooting all registered systems in the Systems Health Monitor, in compliance with FDA 21 CFR Part 11 requirements applicable to GMP-critical electronic records and systems.');
  _blank(body);

  // ── 2. Scope ──
  _h3(body, '2. Scope');
  _p(body, 'This SOP applies to all IT Department personnel performing maintenance on the Systems Health Monitor and its registered systems. It covers:');
  _bullet(body, 'Managing entries in the System Registry tab');
  _bullet(body, 'Responding to automated health check alerts posted to #it-support in Slack');
  _bullet(body, 'Renewing expired or near-expiry credentials and authentication tokens');
  _bullet(body, 'Updating or patching automation scripts');
  _bullet(body, 'Manually reporting observed system failures');
  _bullet(body, 'Periodic validation of GMP-critical systems');
  _blank(body);

  // ── 3. Responsibilities ──
  _h3(body, '3. Responsibilities');
  _tbl(body, [
    ['Role',            'Responsibility'],
    ['IT Team Members', 'Execute all maintenance tasks per this SOP; monitor daily health check alerts in #it-support; update System Registry entries; document all actions in the Heartbeat Log'],
    ['IT Manager',      'First escalation point for unresolved issues; approve all changes to GMP-critical systems; review this SOP every 6 months; sign off on corrective actions'],
    ['QA Manager',      'Approve this SOP and any revisions; oversee validation of GMP-critical systems; confirm FDA 21 CFR Part 11 compliance']
  ]);
  _blank(body);
  _pBold(body, 'Escalation Path');
  _tbl(body, [
    ['Role',                          'Name',                    'Contact',            'Escalate When'],
    ['IT Administrator (primary)',     'Erik Demchuk',            'Slack: @erik',        'Day-to-day maintenance, routine issues'],
    ['IT Manager (escalation 1)',      '[IT Manager Name]',       'Slack: [handle]',     'Issue unresolved > 1 business week; any GMP-critical change; webhook or script change approval'],
    ['QA Manager (escalation 2)',      '[QA Manager Name]',       'Email: [email]',      'SOP deviation; compliance concern; GMP validation sign-off; QA approval required']
  ]);
  _blank(body);

  // ── 4. Definitions ──
  _h3(body, '4. Definitions');
  _tbl(body, [
    ['Term',                    'Definition'],
    ['Systems Health Monitor',  'Google Sheets workbook (2026_Systems-Health-Tracker) with bound Apps Script that runs a daily health check at 8:30 AM and posts results to #it-support in Slack'],
    ['System Registry',         'Tab in the Systems Health Monitor spreadsheet listing every monitored system — including GMP classification, authentication details, owner, heartbeat method, and current status'],
    ['Heartbeat Log',           'Tab in the Systems Health Monitor spreadsheet recording every health check event, heartbeat ping, and system status update with a timestamp'],
    ['Heartbeat',               'A timestamped entry logged by each system to confirm it ran successfully; absence of a heartbeat beyond the expected interval indicates a potential failure'],
    ['GMP Critical',            'Any system marked "Yes" in the GMP Critical column of the System Registry; changes require IT Manager approval before implementation'],
    ['Credential Expiry',       'The date after which an authentication token or API key becomes invalid; the health check warns 30 days before expiry and re-alerts at 14, 7, 3, and 1 days'],
    ['Health Check',            'The automated daily process (8:30 AM) that verifies URL reachability, checks for missing heartbeats, flags expiring credentials, and sends a summary to #it-support'],
    ['Log Cleanup',             'Automated process at 7:00 PM daily that removes Heartbeat Log entries older than today to maintain spreadsheet performance'],
    ['Status values',           'Healthy = running normally; Degraded = warning (heartbeat overdue or credential expiring); Down = failure detected; Unknown = no data yet']
  ]);
  _blank(body);

  // ── 5. Procedure ──
  _h3(body, '5. Procedure');
  _blank(body);

  // 5.1
  _pBold(body, '5.1  Daily Health Check Response');
  _p(body, 'The Systems Health Monitor runs automatically at 8:30 AM each day and posts a status summary to #it-support in Slack. Each IT team member is responsible for reviewing this summary on days they are working.');
  _blank(body);
  _img(body, 'Daily Slack health check summary posted to #it-support — showing total systems, passed, warnings, and failures');
  _blank(body);
  _tbl(body, [
    ['Alert Type',                      'Meaning',                                                     'Required Action',                                                                  'Deadline'],
    ['Missing heartbeat',               'System did not log a successful run within expected interval', 'Investigate root cause and restore; escalate to IT Manager if unresolved',          'Within 1 business week'],
    ['Credential expiring (≤ 30 days)', 'Auth token or API key nearing expiry date',                  'Renew credential before expiry — see Step 5.4',                                   'Before expiry date'],
    ['URL unreachable',                 'HTTP endpoint health check returned error or non-200 code',   'Verify endpoint, connectivity, and system status; escalate if unresolved',         'Within 1 business week'],
    ['Health check did not run',        'No Slack summary received by 9:00 AM',                        'Escalate to IT Manager immediately; investigate trigger failure in Apps Script',   'Same day']
  ]);
  _blank(body);

  // 5.2
  _pBold(body, '5.2  Adding or Updating a System in the Registry');
  _pWithLinks(body, 'When a new system is deployed or an existing one is modified, open the Systems Health Monitor spreadsheet and navigate to the System Registry tab.', [
    { phrase: 'Systems Health Monitor spreadsheet', url: URLS.HEALTH_MONITOR_SHEET }
  ]);
  _blank(body);
  _img(body, 'System Registry tab showing all columns: System Name, Type, Platform, GMP Critical, Auth Type, Expiry Date, Status, Last Heartbeat, etc.');
  _blank(body);
  _bullet(body, 'Add a new row or update the existing row. All required fields must be completed: System Name, Description, Type, Platform, GMP Critical, Run Frequency, Maintenance Guide, Auth Type, Auth Location, Expiry Date, Owner, Validated, Heartbeat Method.');
  _bullet(body, 'If GMP Critical = "Yes", obtain IT Manager approval before saving the entry.');
  _bullet(body, 'Set Validated = "No" and clear Last Validated for new entries pending validation.');
  _bullet(body, 'Run Health Monitor menu → Format Registry Sheet to re-apply consistent formatting and dropdowns.');
  _bullet(body, 'Document the addition or change in the Heartbeat Log: date, description of change, and your name.');
  _blank(body);

  // 5.3
  _pBold(body, '5.3  Removing a System from the Registry');
  _bullet(body, 'Confirm the system has been fully decommissioned (triggers deleted, script disabled or removed).');
  _bullet(body, 'For GMP-critical systems (GMP Critical = "Yes"), obtain IT Manager written approval before proceeding.');
  _bullet(body, 'Delete or mark the row as decommissioned in the System Registry tab.');
  _bullet(body, 'Document the removal in the Heartbeat Log with the date, system name, reason for removal, and your name.');
  _blank(body);

  // 5.4
  _pBold(body, '5.4  Credential Renewal');
  _p(body, 'The health check flags credentials expiring within 30 days (re-alerts at 14, 7, 3, and 1 days). Upon receiving this alert:');
  _bullet(body, 'Identify the credential type and storage location from the Auth Type and Auth Location columns in the System Registry.');
  _bullet(body, 'Renew the credential through the appropriate platform (e.g., Google OAuth re-authorization, API key regeneration, external vendor renewal).');
  _bulletWithLinks(body, 'If the credential is stored in Apps Script Properties: open the Apps Script editor → Project Settings → Script Properties and update the stored value.', [
    { phrase: 'Apps Script editor', url: URLS.HEALTH_MONITOR_SCRIPT }
  ]);
  _bullet(body, 'Update the Expiry Date column in the System Registry. The Days Left column recalculates automatically.');
  _bullet(body, 'Verify the affected system runs successfully and logs a heartbeat after renewal.');
  _bullet(body, 'Document the renewal in the Heartbeat Log: credential name, old expiry, new expiry, date renewed, and your name.');
  _blank(body);

  // 5.5
  _pBold(body, '5.5  Script Updates and Patching');
  _p(body, 'When changes to a system script are required:');
  _bullet(body, 'For GMP-critical systems: obtain IT Manager approval before making any code changes.');
  _bulletWithLinks(body, 'Open the relevant spreadsheet → Extensions → Apps Script. For the Health Monitor itself, open the Apps Script editor directly.', [
    { phrase: 'Apps Script editor directly', url: URLS.HEALTH_MONITOR_SCRIPT }
  ]);
  _bullet(body, 'Make the required changes to the script.');
  _bullet(body, 'Test manually: Run → select the relevant function and verify it executes without errors.');
  _bullet(body, 'Confirm the system logs a heartbeat or produces expected output after the change.');
  _bullet(body, 'Document in the Heartbeat Log: date, function(s) modified, reason for change, test outcome, and your name.');
  _blank(body);

  // 5.6
  _pBold(body, '5.6  Manual Incident Reporting');
  _p(body, 'If an IT team member observes a system failure not captured by the automated health check:');
  _bullet(body, 'Report the incident in #it-support with: system name, observed symptom, time first noticed, and any visible error messages.');
  _bullet(body, 'Log the incident in the Heartbeat Log tab: date, system affected, description of failure, steps taken, and outcome.');
  _bullet(body, 'If the issue is not resolved within 1 business week, escalate to the IT Manager.');
  _blank(body);

  // 5.7
  _pBold(body, '5.7  Periodic Validation of GMP-Critical Systems');
  _p(body, 'Every 6 months, coinciding with the SOP review cycle, all systems with GMP Critical = "Yes" must be validated:');
  _bullet(body, 'Review the Last Run OK column to confirm the system has been running successfully.');
  _bullet(body, 'Confirm heartbeats are logged at the expected Run Frequency.');
  _bullet(body, 'Manually trigger the system (if safe to do so) and verify output.');
  _bullet(body, 'Update Validated = "Yes" and Last Validated = validation date in the System Registry.');
  _bullet(body, 'Obtain IT Manager sign-off on validation results.');
  _bullet(body, 'Document validation in the Heartbeat Log: system name, validation date, outcome, and IT Manager name.');
  _blank(body);

  // 5.8
  _pBold(body, '5.8  Reformatting the System Registry');
  _pWithLinks(body, 'If the System Registry tab loses formatting, open the Systems Health Monitor Apps Script editor and run Health Monitor menu → Format Registry Sheet. The script re-applies all headers, color coding, dropdowns, and column widths while preserving live Status, Last Heartbeat, and Last Run OK data.', [
    { phrase: 'Systems Health Monitor Apps Script editor', url: URLS.HEALTH_MONITOR_SCRIPT }
  ]);
  _blank(body);

  // 5.9
  _pBold(body, '5.9  Connecting a New System to the Health Monitor');
  _p(body, 'When a new automation or script is deployed, connect it to the Health Monitor so failures are detected automatically. Complete all steps in order:');
  _blank(body);
  _pBold(body, 'Step A — Deploy the Health Monitor web app (or confirm it is deployed)');
  _bulletWithLinks(body, 'Open the Health Monitor Apps Script editor.', [
    { phrase: 'Health Monitor Apps Script editor', url: URLS.HEALTH_MONITOR_SCRIPT }
  ]);
  _bullet(body, 'Click Deploy → Manage Deployments. If no active deployment exists, click New Deployment → Web App.');
  _bullet(body, 'Set Execute as: Me, Who has access: Anyone. Click Deploy and copy the web app URL.');
  _bullet(body, 'NOTE: Re-deploy (New Version) every time the Health Monitor script code is changed — the URL stays the same, but the code updates only after re-deployment.');
  _blank(body);
  _pBold(body, 'Step B — Add HEALTH_MONITOR_URL to the new system\u2019s Script Properties');
  _bullet(body, 'Open the Apps Script editor for the system being connected (not the Health Monitor).');
  _bullet(body, 'Go to Project Settings (gear icon) \u2192 Script Properties \u2192 Add property.');
  _bullet(body, 'Key: HEALTH_MONITOR_URL   |   Value: [paste the web app URL from Step A]');
  _blank(body);
  _pBold(body, 'Step C \u2014 Add sendHeartbeatToMonitor_() to the new system\u2019s script');
  _p(body, 'Paste the following helper function into the script (copy from heartbeat-plan.txt or the shared function below):');
  _blank(body);
  _tbl(body, [
    ['Function',                  'Where to call it',                                                    'Payload fields required'],
    ['sendHeartbeatToMonitor_()', 'At the END of the main trigger function (both success and catch paths)', 'system (must match System Registry name exactly), status ("success" or "error"), details (optional string)'],
    ['Success path',              'After all main logic completes successfully',                           '{ system: "System Name", status: "success", details: "e.g. 5 records processed" }'],
    ['Error path (catch block)',  'Inside the catch(e) block',                                            '{ system: "System Name", error: e.message, severity: "error" }']
  ]);
  _blank(body);
  _p(body, 'For Shopify Flows (no Apps Script): add a "Send HTTP Request" action at the END of the flow on both the success and error branches. POST to the HEALTH_MONITOR_URL with the JSON payload above.');
  _blank(body);
  _pBold(body, 'Step D \u2014 Test and register the heartbeat');
  _bullet(body, 'Run the main function once manually and verify a new row appears in the Heartbeat Log tab of the Systems Health Monitor.');
  _bullet(body, 'Confirm the System Name in the payload matches the System Registry exactly (case-sensitive).');
  _bullet(body, 'In the System Registry, update: Heartbeat Method (e.g., "Script sendHeartbeat()"), Last Heartbeat (auto-updates), Last Run OK = "Yes".');
  _bullet(body, 'Document the connection in the Heartbeat Log: date, system name, step completed, and your name.');
  _blank(body);

  // 5.10
  _pBold(body, '5.10  Troubleshooting');
  _blank(body);

  _p(body, '5.10.1  Health Check Did Not Run — No Slack Summary by 9:00 AM');
  _p(body, 'Symptom: No daily health check message in #it-support after 9:00 AM.');
  _bulletWithLinks(body, 'Open the Health Monitor Apps Script editor \u2192 Triggers (clock icon on left sidebar).', [
    { phrase: 'Health Monitor Apps Script editor', url: URLS.HEALTH_MONITOR_SCRIPT }
  ]);
  _bullet(body, 'Verify the runHealthCheck trigger is set to time-based, daily, near 8:30 AM. If missing, run Health Monitor menu \u2192 Setup Triggers.');
  _bullet(body, 'Check Executions (play icon) for any failed execution of runHealthCheck and expand to read the error message.');
  _bullet(body, 'If the trigger exists but failed silently, check that both the System Registry and Heartbeat Log tabs still exist in the spreadsheet.');
  _blank(body);

  _p(body, '5.10.2  Slack Webhook Not Delivering');
  _p(body, 'Symptom: Health check runs (Heartbeat Log is updating) but no messages arrive in #it-support.');
  _bulletWithLinks(body, 'Open the Health Monitor Apps Script editor \u2192 Project Settings \u2192 Script Properties. Verify SLACK_WEBHOOK_URL is set.', [
    { phrase: 'Health Monitor Apps Script editor', url: URLS.HEALTH_MONITOR_SCRIPT }
  ]);
  _bullet(body, 'Run Health Monitor menu \u2192 Test Slack Connection. If this fails, the webhook URL has expired or been revoked.');
  _bullet(body, 'To renew: go to Slack \u2192 workspace Apps \u2192 Incoming Webhooks \u2192 create new webhook for #it-support. Paste the new URL into SLACK_WEBHOOK_URL Script Property.');
  _blank(body);

  _p(body, '5.10.3  System Showing "Unknown" Status Despite Running');
  _p(body, 'Symptom: A system is operational but the System Registry shows Status = Unknown and Last Heartbeat is blank or old.');
  _bullet(body, 'The system is not connected to the Health Monitor. Follow Step 5.9 to add heartbeat integration.');
  _bullet(body, 'If heartbeat integration was set up previously: verify HEALTH_MONITOR_URL in the target system\u2019s Script Properties still points to the current deployed web app URL.');
  _bullet(body, 'Re-deploy the Health Monitor web app (Step 5.9A) if it was recently modified, then re-run the target system to send a fresh heartbeat.');
  _blank(body);

  _p(body, '5.10.4  System Registry Formatting Lost');
  _p(body, 'Symptom: Columns are missing, dropdowns are gone, or color coding has been removed.');
  _bulletWithLinks(body, 'Run Health Monitor menu \u2192 Format Registry Sheet. This re-applies all headers, colors, dropdowns, and column widths from the System Registry template defined in sheet-setup.js.', []);
  _bullet(body, 'If a column was accidentally deleted (not just formatted), check the Heartbeat Log for the last known good data before the deletion. Restore manually or from a Google Sheets version history (File \u2192 Version history \u2192 See version history).');
  _bullet(body, 'Do NOT add, remove, or reorder columns manually \u2014 the health check script reads columns by fixed position numbers. Any structural change requires an update to the script.');
  _blank(body);

  // ── 6. References ──
  _h3(body, '6. References');
  _refLink(body, 'Systems Health Monitor (spreadsheet)', URLS.HEALTH_MONITOR_SHEET);
  _refLink(body, 'Systems Health Monitor Apps Script', URLS.HEALTH_MONITOR_SCRIPT);
  _refLink(body, 'Security & GMP Connection Tracker (spreadsheet)', URLS.SECURITY_TRACKER_SHEET);
  _refLink(body, 'Security & GMP Connection Tracker Apps Script', URLS.SECURITY_TRACKER_SCRIPT);
  _refLink(body, 'IT-015: IT Security Tracker Maintenance', URLS.IT015_DOC);
  _refLink(body, 'FDA 21 CFR Part 11 — Electronic Records; Electronic Signatures', 'https://www.ecfr.gov/current/title-21/chapter-I/subchapter-A/part-11');
  _blank(body);

  // ── 7. Attachments ──
  _h3(body, '7. Attachments');
  _p(body, 'Attachment A — System Registry Column Definitions');
  _blank(body);
  _tbl(body, [
    ['Column',            'Description'],
    ['System Name',       'Unique name identifying the system or automation'],
    ['Description',       'Brief description of what the system does'],
    ['Type',              'System type: FLOW (Shopify Flow), SCRIPT (Apps Script), BOT (standalone bot), DASHBOARD (Google Sheet)'],
    ['Platform',          'Primary platform (e.g., Shopify, Google, Katana, FedEx, Wasp/Katana)'],
    ['GMP Critical',      '"Yes" = GMP-critical, requires IT Manager approval for all changes; "No" = non-critical'],
    ['Run Frequency',     'How often the system runs (e.g., Daily 8:30 AM, Per order, Weekly, On demand)'],
    ['Maintenance Guide', 'SOP ID or hyperlink to the maintenance guide for this system'],
    ['Auth Type',         'Authentication type required (e.g., API Key, Bearer Token, OAuth, None)'],
    ['Auth Location',     'Where credentials are stored (e.g., Apps Script → Script Properties → KEY_NAME)'],
    ['Expiry Date',       'Credential expiry date; "NO EXPIRY" if the credential does not expire; blank if no credential required'],
    ['Days Left',         'Days until credential expires — auto-calculated daily from Expiry Date. Red if expired, yellow if \u2264 30 days'],
    ['Owner',             'Name of the person responsible for this system'],
    ['Validated',         '"Yes" = GMP validation complete; "No" = pending validation; "Pending" = validation in progress'],
    ['Last Validated',    'Date of the most recent GMP validation'],
    ['Heartbeat Method',  'How this system reports its status (e.g., Script sendHeartbeat(), Flow HTTP step, N/A)'],
    ['Status',            'Current operational status: Healthy / Degraded / Down / Unknown'],
    ['Last Heartbeat',    'Timestamp of the most recent successful heartbeat (yyyy-MM-dd HH:mm:ss)'],
    ['Last Run OK',       '"Yes" = last run succeeded; "No" = last run failed; "N/A" = no run data available'],
    ['Notes',             'Additional context, known issues, maintenance history, or external links (e.g., GitHub, Drive)']
  ]);
  _blank(body);
  _img(body, 'Full System Registry tab in the Systems Health Monitor — showing all 19 columns with color-coded status, platform chips, and expiry dates');
  _blank(body);

  // ── 8. Revision History ──
  _h3(body, '8. Revision History');
  _tbl(body, [
    ['Version', 'Effective Date', 'Description of Change',                                                                                          'Change Control #'],
    ['1.4',     'March 11, 2026', 'Updated live spreadsheet and Apps Script links, corrected connection steps, and aligned references with the current tracker suite', 'n/a'],
    ['1.3',     'March 4, 2026',  'Full SOP rewrite \u2014 GMP-compliant 8-section format; added troubleshooting (5.10), heartbeat integration guide (5.9), and escalation contacts table', 'n/a'],
    ['',        '',               '',                                                                                                                ''],
    ['',        '',               '',                                                                                                                '']
  ]);

  Logger.log('\u2713 IT-014 rewrite complete: ' + doc.getUrl());
}


// ─── IT-015: IT Security Tracker Maintenance ─────────────────────────────────
// Opens the EXISTING document by ID — body.clear() wipes content, does NOT create a new doc.

function rewriteIT015() {
  var doc  = DocumentApp.openById(SOP_TO_FIX[1].id);
  var body = doc.getBody();
  body.clear();

  // ── Title ──
  _title(body, 'IT-015: IT Security Tracker Maintenance');
  _blank(body);

  // ── Metadata ──
  _meta(body,     'Department: ',               'Information Technology');
  _meta(body,     'SOP ID: ',                   'IT-015');
  _meta(body,     'Version: ',                  '1.1');
  _meta(body,     'Effective Date: ',           'March 11, 2026');
  _meta(body,     'Review Cycle: ',             'Every 6 months');
  _metaLink(body, 'Security & GMP Connection Tracker: ', 'Open spreadsheet \u2192', URLS.SECURITY_TRACKER_SHEET);
  _metaLink(body, 'Security Tracker Apps Script: ',      'Open editor \u2192',      URLS.SECURITY_TRACKER_SCRIPT);
  _metaLink(body, 'Systems Health Monitor: ',            'Open spreadsheet \u2192', URLS.HEALTH_MONITOR_SHEET);
  _meta(body,     'Prepared by: ',              'Erik Demchuk');
  _meta(body,     'Reviewed by (IT Manager): ', '________________________________');
  _meta(body,     'Approved by (QA Manager): ', '________________________________');
  _blank(body);

  // ── 1. Purpose ──
  _h3(body, '1. Purpose');
  _p(body, 'To define procedures for maintaining, monitoring, and troubleshooting the Security & GMP Connection Tracker, a GMP-oriented Google Sheets and Apps Script workflow used to manage access requests, active GMP accounts, review checks, and security incidents.');
  _blank(body);

  // ── 2. Scope ──
  _h3(body, '2. Scope');
  _p(body, 'This SOP applies to IT personnel who administer the tracker workbook, its Slack intake, and its related Apps Script automations. It covers:');
  _bullet(body, 'Slack intake for access requests and incident reports');
  _bullet(body, 'Maintenance of the Access Control, Review Checks, and Incidents tabs');
  _bullet(body, 'Sync of approved requests into the ACTIVE GMP ACCOUNTS section');
  _bullet(body, 'Refresh of Katana and Shopify activity evidence');
  _bullet(body, 'Generation and completion of periodic review checks');
  _bullet(body, 'Verification of Slack, heartbeat, and Apps Script trigger integrations');
  _blank(body);

  // ── 3. Responsibilities ──
  _h3(body, '3. Responsibilities');
  _tbl(body, [
    ['Role',             'Responsibility'],
    ['IT Administrator', 'Review new requests, maintain workbook structure, run sync and refresh jobs, investigate automation failures, and keep documentation current'],
    ['IT Manager',       'Approve material workflow or script changes, resolve escalated access issues, and review periodic access-review outcomes'],
    ['QA Manager',       'Approve SOP revisions and oversee GMP-control expectations for access tracking and incident documentation']
  ]);
  _blank(body);

  // ── 4. Definitions ──
  _h3(body, '4. Definitions');
  _tbl(body, [
    ['Term',                            'Definition'],
    ['Access Control tab',          'Primary sheet tab containing ACCESS REQUESTS in the upper section and ACTIVE GMP ACCOUNTS in the lower section'],
    ['ACCESS REQUESTS',             'Request queue where Slack or manual intake rows first appear with approval and provisioning workflow fields'],
    ['ACTIVE GMP ACCOUNTS',         'Operational account inventory grouped by system category (Katana, Wasp, Shopify, ShipStation, 1Password, Other Systems)'],
    ['Review Checks tab',           'Sheet tab holding OPEN REVIEW QUEUE and COMPLETED REVIEWS generated from active-account risk and activity data'],
    ['Incidents tab',               'Sheet tab for security issue intake, response ownership, root cause, and closure tracking'],
    ['GMP Security menu',           'Custom spreadsheet menu used to set up the workbook, install triggers, sync approved requests, refresh account signals, generate reviews, and test Slack'],
    ['Slack intake',                'Slash-command workflow using /gmp-access, /gmp-request-access, and /gmp-issue to create records through Slack modals'],
    ['Request sync',                'Automation that converts approved request rows into active-account records or revokes matching accounts when removal is completed'],
    ['HEALTH_MONITOR_URL',          'Script Property storing the Systems Health Monitor web app endpoint used for heartbeat reporting'],
    ['Slack thread',                'Stored permalink used so follow-up messages remain attached to the original Slack request or issue conversation']
  ]);
  _blank(body);

  // ── 5. Procedure ──
  _h3(body, '5. Procedure');
  _blank(body);

  // 5.1
  _pBold(body, '5.1  Workbook Structure and Category Layout');
  _pWithLinks(body, 'Open the Security & GMP Connection Tracker spreadsheet to maintain the live workbook structure.', [
    { phrase: 'Security & GMP Connection Tracker spreadsheet', url: URLS.SECURITY_TRACKER_SHEET }
  ]);
  _bullet(body, 'Access Control contains ACCESS REQUESTS first, followed by ACTIVE GMP ACCOUNTS.');
  _bullet(body, 'ACTIVE GMP ACCOUNTS must remain grouped under these category headers in order: Katana, Wasp, Shopify, ShipStation, 1Password, Other Systems.');
  _bullet(body, 'Review Checks contains OPEN REVIEW QUEUE and COMPLETED REVIEWS.');
  _bullet(body, 'Incidents is used for GMP-related security issue intake and follow-up.');
  _bullet(body, 'Do not manually reorder, rename, or delete workflow columns. The Apps Script code reads sections and columns by fixed schema.');
  _blank(body);
  _img(body, 'Access Control tab showing ACCESS REQUESTS at the top and ACTIVE GMP ACCOUNTS grouped by Katana, Wasp, Shopify, ShipStation, 1Password, and Other Systems');
  _blank(body);

  // 5.2
  _pBold(body, '5.2  Slack Intake and Request Creation');
  _p(body, 'Access intake is designed to start in Slack so the requestor, target user, system, access level, and business reason are captured consistently.');
  _bullet(body, 'Use /gmp-access or /gmp-request-access to open the access modal.');
  _bullet(body, 'Use /gmp-issue to open the incident modal.');
  _bullet(body, 'Submitted access requests are written into ACCESS REQUESTS with Approval = Submitted and Provisioning = Open.');
  _bullet(body, 'The queue processor runs every minute, so new Slack submissions may appear shortly after modal submission rather than instantly.');
  _bullet(body, 'Each intake row should retain its Slack Thread value so status updates can post back into the same conversation.');
  _blank(body);
  _img(body, 'Slack access modal feeding a new row into ACCESS REQUESTS with request type, target user, company email, GMP system, access level, manager, and reason');
  _blank(body);

  // 5.3
  _pBold(body, '5.3  Approval, Provisioning, and Sync to ACTIVE GMP ACCOUNTS');
  _p(body, 'After reviewing a request row, update the approval and provisioning fields so the tracker can convert the request into an account record.');
  _bullet(body, 'For new or changed access, set Approval = Approved once the request is authorized.');
  _bullet(body, 'Use Provisioning values to reflect progress: Open, Provisioning, Provisioned, Removal Pending, or Closed.');
  _bullet(body, 'Run GMP Security menu → Sync Approved Requests after approvals or provisioning changes are entered.');
  _bullet(body, 'The sync process creates or updates an ACTIVE GMP ACCOUNTS row using Access ID first, then Company Email + GMP System if no Access ID exists yet.');
  _bullet(body, 'New and changed accounts are stored under the correct system category header automatically.');
  _bullet(body, 'Remove Access requests mark the matching active account as Revoked once the removal is provisioned or closed.');
  _bullet(body, 'Confirm the synced row has the expected Access ID, owner, review status, and notes before closing the request work item.');
  _blank(body);

  // 5.4
  _pBold(body, '5.4  Refreshing Account Evidence and Daily Maintenance');
  _p(body, 'Use the GMP Security menu to refresh evidence and maintain the account inventory.');
  _tbl(body, [
    ['Menu Action',                 'Purpose'],
    ['Refresh Katana Accounts',     'Pull Katana account presence data and update account rows'],
    ['Refresh Activity Signals',    'Recalculate generic activity bands and review indicators from current evidence'],
    ['Refresh Shopify Activity',    'Update Shopify-specific activity evidence and then refresh broader activity signals'],
    ['Generate Review Checks',      'Create or update review rows in Review Checks based on active-account risk and staleness'],
    ['Run Daily Maintenance',       'Run the standard daily sequence for account refresh and review generation']
  ]);
  _blank(body);
  _bullet(body, 'Run Refresh Katana Accounts whenever Katana users are added, removed, or materially changed.');
  _bullet(body, 'Run Refresh Shopify Activity when Shopify evidence or order-flow usage needs to be refreshed.');
  _bullet(body, 'Use Run Daily Maintenance for the normal end-to-end routine rather than running each refresh manually unless you are troubleshooting.');
  _blank(body);

  // 5.5
  _pBold(body, '5.5  Review Queue Maintenance');
  _p(body, 'Review Checks is the working queue for periodic access review decisions.');
  _bullet(body, 'Run GMP Security menu → Generate Review Checks after account refreshes or before a formal review cycle.');
  _bullet(body, 'Work the OPEN REVIEW QUEUE section first. Use Reviewer Action values Keep, Reduce, Remove, or Need Info.');
  _bullet(body, 'Set Decision Status to Done only after the review outcome is documented and any required access changes have been submitted.');
  _bullet(body, 'Completed items remain in COMPLETED REVIEWS as the audit trail for prior review actions.');
  _bullet(body, 'If a review results in access removal or reduction, create or update the corresponding ACCESS REQUESTS item and re-run the request sync after approval.');
  _blank(body);

  // 5.6
  _pBold(body, '5.6  Incident Intake and Follow-up');
  _p(body, 'Use the Incidents tab for security issues that affect GMP systems, access control, or auditability.');
  _bullet(body, 'Preferred intake is /gmp-issue from Slack so the Slack thread is captured automatically.');
  _bullet(body, 'Populate Severity, System, Summary, Assigned To, Status, Response Target, and any linked access or request IDs.');
  _bullet(body, 'Use Status values Open, Investigating, and Resolved.');
  _bullet(body, 'Document Root Cause and Resolution before setting the incident to Resolved.');
  _bullet(body, 'If the incident requires an access correction, cross-reference the affected Access ID or Request ID and maintain both records together.');
  _blank(body);
  _img(body, 'Incidents tab showing Severity, Status, Response Target, linked access/request IDs, and Slack thread tracking');
  _blank(body);

  // 5.7
  _pBold(body, '5.7  Trigger and Integration Verification');
  _blank(body);

  _pWithLinks(body, 'Open the Security Tracker Apps Script project when verifying automation settings.', [
    { phrase: 'Security Tracker Apps Script project', url: URLS.SECURITY_TRACKER_SCRIPT }
  ]);
  _bullet(body, 'Run GMP Security menu → Setup Triggers if triggers are missing or were deleted.');
  _bullet(body, 'Verify these installable triggers exist: onEditInstallable, runDailySecurityMaintenance, keepSecurityCommandWarm_, and processPendingSecurityIntake_.');
  _bullet(body, 'The daily maintenance trigger should run near 8:00 AM local time. The keepalive trigger runs every 5 minutes and the intake processor runs every 1 minute.');
  _bullet(body, 'Run GMP Security menu → Test Slack Connection after webhook or bot changes.');
  _bullet(body, 'In Script Properties, confirm SLACK_WEBHOOK_URL, SLACK_BOT_TOKEN, and HEALTH_MONITOR_URL are populated when troubleshooting integrations.');
  _bullet(body, 'Use Executions in Apps Script to inspect failed runs before making corrective changes.');
  _blank(body);

  // 5.8
  _pBold(body, '5.8  Troubleshooting');
  _blank(body);

  _p(body, '5.8.1  Approved request did not appear in ACTIVE GMP ACCOUNTS');
  _bullet(body, 'Verify Approval = Approved on the request row.');
  _bullet(body, 'For New Access or Change Access, re-run GMP Security menu → Sync Approved Requests.');
  _bullet(body, 'For Remove Access, confirm Provisioning is Provisioned or Closed before expecting the active-account row to change to Revoked.');
  _bullet(body, 'Check that Company Email and GMP System are populated on the request row because those fields are used for fallback matching when Access ID is blank.');
  _blank(body);

  _p(body, '5.8.2  Account landed in the wrong section or category');
  _bullet(body, 'Use the canonical system names Katana, Wasp, Shopify, ShipStation, and 1Password whenever possible.');
  _bullet(body, 'Run Sync Approved Requests again, then re-run the account sort through the same workflow if the grouping still looks stale.');
  _bullet(body, 'Do not manually drag account rows between category headers because sorting and sync logic will overwrite manual placement.');
  _blank(body);

  _p(body, '5.8.3  Slack modal or follow-up message is not working');
  _bullet(body, 'Check Apps Script Executions for errors in doPost, Slack intake processing, or Slack notification helpers.');
  _bullet(body, 'Run GMP Security menu → Test Slack Connection.');
  _bullet(body, 'Verify the deployed web app is current after code changes to slash-command handling.');
  _blank(body);

  _p(body, '5.8.4  Daily maintenance or review generation did not run');
  _bullet(body, 'Open the Apps Script Triggers page and confirm runDailySecurityMaintenance still exists.');
  _bullet(body, 'Inspect the latest execution record for the exact failing function.');
  _bullet(body, 'If needed, recreate triggers with GMP Security menu → Setup Triggers and document the corrective action.');
  _blank(body);

  // ── 6. References ──
  _h3(body, '6. References');
  _refLink(body, 'IT-014: Automation Maintenance',                    URLS.IT014_DOC);
  _refLink(body, 'IT-008: IT Security & Access Control',              URLS.IT008_DOC);
  _refLink(body, 'IT-010: Cyber Security Incident Response',          URLS.IT010_DOC);
  _refLink(body, 'Security & GMP Connection Tracker (spreadsheet)',   URLS.SECURITY_TRACKER_SHEET);
  _refLink(body, 'Security Tracker Apps Script',                      URLS.SECURITY_TRACKER_SCRIPT);
  _refLink(body, 'Systems Health Monitor (spreadsheet)',              URLS.HEALTH_MONITOR_SHEET);
  _refLink(body, 'Systems Health Monitor Apps Script',                URLS.HEALTH_MONITOR_SCRIPT);
  _refLink(body, 'FDA 21 CFR Part 11 — Electronic Records; Electronic Signatures', 'https://www.ecfr.gov/current/title-21/chapter-I/subchapter-A/part-11');
  _blank(body);

  // ── 7. Attachments ──
  _h3(body, '7. Attachments');
  _p(body, 'Attachment A — ACCESS REQUESTS key fields');
  _blank(body);
  _tbl(body, [
    ['Field',               'Description / Values'],
    ['Request ID',          'Stable request identifier generated by the workbook'],
    ['Queue Priority',      'High / Medium / Low'],
    ['Request Type',        'New Access / Change Access / Remove Access'],
    ['Target User',         'Person whose account is being created, changed, or removed'],
    ['Company Email',       'Primary matching key for fallback sync logic'],
    ['GMP System',          'System name used for sync and category placement'],
    ['Access Level',        'Admin / Full Access / Read-Write / Read-Only / Limited'],
    ['Approval',            'Submitted / Approved / Denied'],
    ['Provisioning',        'Open / Provisioning / Provisioned / Removal Pending / Closed'],
    ['Access ID',           'Populated after sync or provisioning when an account record exists'],
    ['Slack Thread / Notes','Links the request to Slack follow-up and audit comments']
  ]);
  _blank(body);

  _p(body, 'Attachment B — ACTIVE GMP ACCOUNTS key fields');
  _blank(body);
  _tbl(body, [
    ['Field',                  'Description / Values'],
    ['Access ID',              'Primary account identifier in the tracker'],
    ['GMP System',             'Canonical category-driving system value'],
    ['Person / Company Email', 'Human identity attached to the account row'],
    ['Platform Account',       'Actual account or login identifier when known'],
    ['Access Level',           'Current granted access level'],
    ['Presence Status',        'Provisioning / Present / Missing / Unmanaged / Revoked'],
    ['Activity Status',        'Verified Active / Some Evidence / No Evidence / Unknown'],
    ['Last Activity Evidence', 'Latest login or usage evidence used for account review'],
    ['Next Review Due',        'Date driving review queue generation'],
    ['Review Status',          'Open / Waiting / Done'],
    ['Owner',                  'Internal owner responsible for maintaining the access'],
    ['Notes',                  'Audit trail of sync, provisioning, and review actions']
  ]);
  _blank(body);

  _p(body, 'Attachment C — REVIEW CHECKS key fields');
  _blank(body);
  _tbl(body, [
    ['Field',                    'Description / Values'],
    ['Review ID',                'Stable review identifier'],
    ['Review Cycle / Priority',  'Scheduling and priority controls'],
    ['Person / System / Access Level', 'The account under review'],
    ['Presence Status / Activity Status', 'Signals used to decide if access is still justified'],
    ['Why In Review',            'Reason the row entered the queue'],
    ['Reviewer Action',          'Keep / Reduce / Remove / Need Info'],
    ['Decision Status',          'Open / Waiting / Done'],
    ['Completed At / Next Review Due', 'Review completion and forward scheduling'],
    ['Linked Request / Slack Thread', 'Cross-reference to access changes and communication']
  ]);
  _blank(body);

  _p(body, 'Attachment D — INCIDENTS key fields');
  _blank(body);
  _tbl(body, [
    ['Field',                 'Description / Values'],
    ['Incident ID',           'Stable incident identifier'],
    ['Reported At / Severity','Event date and High / Medium / Low impact level'],
    ['System / Summary',      'Affected system and concise description of the issue'],
    ['Assigned To / Status',  'Current owner and Open / Investigating / Resolved state'],
    ['Response Target',       'Internal escalation or response destination'],
    ['Resolution / Root Cause', 'Documented fix and underlying cause'],
    ['Linked Access ID / Linked Request ID', 'Connection to relevant access-control records'],
    ['Slack Thread / Notes',  'Conversation trace and supporting audit notes']
  ]);
  _blank(body);

  // ── 8. Revision History ──
  _h3(body, '8. Revision History');
  _tbl(body, [
    ['Version', 'Effective Date',   'Description of Change', 'Change Control #'],
    ['1.1',     'March 11, 2026',   'Updated SOP to match the live 3-tab workbook, GMP Security menu, Slack slash-command intake, grouped account categories, and current trigger set', 'n/a'],
    ['1.0',     'January 30, 2026', 'Initial release', 'n/a'],
    ['',        '',                 '', '']
  ]);

  Logger.log('\u2713 IT-015 rewrite complete: ' + doc.getUrl());
}


/**
 * rewriteBoth — rewrites both SOPs in one go.
 * Both docs are modified IN PLACE (existing document IDs, no new docs created).
 */
function rewriteBoth() {
  rewriteIT014();
  rewriteIT015();
  Logger.log('\u2713 Both SOPs rewritten. Check the Google Docs to review.');
}
