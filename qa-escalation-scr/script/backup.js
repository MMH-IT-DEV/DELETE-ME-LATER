/**
 * QA Escalation — Self-Backup Module v2.4
 * Saves script source code to Google Drive via Apps Script API.
 * Access: QA Escalation menu → "💾 Save Backup"
 */

// ============ SAVE BACKUP ============
function saveBackup() {
  var ui = SpreadsheetApp.getUi();
  var scriptId = ScriptApp.getScriptId();
  var token = ScriptApp.getOAuthToken();

  logHeartbeat(BACKUP.SYSTEM_NAME, '⏳ Started', 'Backup initiated');

  try {
    var url = 'https://script.googleapis.com/v1/projects/' + scriptId + '/content';
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      var errorMsg = 'API Error: ' + response.getResponseCode();
      logError(BACKUP.SYSTEM_NAME, errorMsg);
      ui.alert('❌ ' + errorMsg);
      return;
    }

    var result = JSON.parse(response.getContentText());
    var content = '// ═══════════════════════════════════════════════════════════════\n';
    content += '// BACKUP: ' + BACKUP.NAME + ' v' + BACKUP.VERSION + '\n';
    content += '// Date: ' + new Date().toLocaleString() + '\n';
    content += '// Type: ' + BACKUP.TYPE + '\n';
    content += '// Script: https://script.google.com/d/' + scriptId + '/edit\n';
    content += '// ═══════════════════════════════════════════════════════════════\n\n';

    for (var i = 0; i < result.files.length; i++) {
      content += '// ─────────────────────────────────────────────────────────────────\n';
      content += '// FILE: ' + result.files[i].name + ' (' + result.files[i].type + ')\n';
      content += '// ─────────────────────────────────────────────────────────────────\n\n';
      content += result.files[i].source + '\n\n';
    }

    var rootFolderId = getOrCreateFolder('Apps Script Backups', 'root', token);
    var typeFolderId = getOrCreateFolder(BACKUP.TYPE, rootFolderId, token);

    var props = PropertiesService.getScriptProperties();
    var existingFileId = props.getProperty('BACKUP_FILE_ID');
    var filename = BACKUP.NAME.replace(/\s/g, '_') + '_v' + BACKUP.VERSION + '.txt';

    if (existingFileId) {
      var updated = updateFile(existingFileId, filename, content, token);
      if (updated) {
        logHeartbeat(BACKUP.SYSTEM_NAME, '✓ Success', 'Backup updated: ' + filename);
        ui.alert('✓ Backup Updated!\n\n' + filename + '\n\nFolder: Apps Script Backups/' + BACKUP.TYPE);
        return;
      }
    }

    var fileId = createFileInFolder(filename, content, typeFolderId, token);
    props.setProperty('BACKUP_FILE_ID', fileId);

    logHeartbeat(BACKUP.SYSTEM_NAME, '✓ Success', 'Backup created: ' + filename);
    ui.alert('✓ Backup Created!\n\n' + filename + '\n\nFolder: Apps Script Backups/' + BACKUP.TYPE);

  } catch(e) {
    logError(BACKUP.SYSTEM_NAME, e.message);
    ui.alert('❌ Error: ' + e.message);
    Logger.log(e.stack);
  }
}

// ============ BACKUP INFO ============
function backupInfo() {
  var ui = SpreadsheetApp.getUi();
  var fileId = PropertiesService.getScriptProperties().getProperty('BACKUP_FILE_ID');

  if (!fileId) {
    ui.alert('No backup yet.\n\nRun "💾 Save Backup" first.');
    return;
  }

  ui.alert('📁 Backup Info\n\n' +
    'Project: ' + BACKUP.NAME + '\n' +
    'Version: ' + BACKUP.VERSION + '\n' +
    'Type: ' + BACKUP.TYPE + '\n\n' +
    'URL: https://drive.google.com/file/d/' + fileId);
}

// ============ DRIVE HELPERS ============
function getOrCreateFolder(name, parentId, token) {
  var query = "name='" + name + "' and mimeType='application/vnd.google-apps.folder' and trashed=false";
  if (parentId !== 'root') {
    query += " and '" + parentId + "' in parents";
  }

  var searchUrl = 'https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(query);
  var searchResponse = UrlFetchApp.fetch(searchUrl, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  var searchResult = JSON.parse(searchResponse.getContentText());
  if (searchResult.files && searchResult.files.length > 0) {
    return searchResult.files[0].id;
  }

  var metadata = { name: name, mimeType: 'application/vnd.google-apps.folder' };
  if (parentId !== 'root') { metadata.parents = [parentId]; }

  var createResponse = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(metadata),
    muteHttpExceptions: true
  });

  return JSON.parse(createResponse.getContentText()).id;
}

function createFileInFolder(filename, content, folderId, token) {
  var boundary = '-------314159265358979323846';
  var metadata = JSON.stringify({ name: filename, mimeType: 'text/plain', parents: [folderId] });

  var body = '--' + boundary + '\r\n';
  body += 'Content-Type: application/json; charset=UTF-8\r\n\r\n';
  body += metadata + '\r\n';
  body += '--' + boundary + '\r\n';
  body += 'Content-Type: text/plain\r\n\r\n';
  body += content + '\r\n';
  body += '--' + boundary + '--';

  var response = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'multipart/related; boundary=' + boundary },
    payload: body,
    muteHttpExceptions: true
  });

  return JSON.parse(response.getContentText()).id;
}

function updateFile(fileId, filename, content, token) {
  try {
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId, {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ name: filename }),
      muteHttpExceptions: true
    });

    var boundary = '-------314159265358979323846';
    var body = '--' + boundary + '\r\n';
    body += 'Content-Type: text/plain\r\n\r\n';
    body += content + '\r\n';
    body += '--' + boundary + '--';

    var response = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + fileId + '?uploadType=multipart', {
      method: 'PATCH',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'multipart/related; boundary=' + boundary },
      payload: body,
      muteHttpExceptions: true
    });

    return response.getResponseCode() === 200;
  } catch(e) {
    return false;
  }
}
