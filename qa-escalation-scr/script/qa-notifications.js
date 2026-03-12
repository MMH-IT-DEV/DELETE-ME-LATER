/**
 * QA Escalation — Slack Notifications
 * Handles Slack message sending for QA escalations and error alerts.
 */

// ============ QA ESCALATION NOTIFICATION ============
function sendSlackNotification(data, row) {
  var webhookUrl = getSlackWebhookUrl_();
  if (!webhookUrl) {
    Logger.log('Slack webhook not configured. Set SLACK_WEBHOOK_URL in Script Properties.');
    return false;
  }

  var photoLink = '-';
  if (data.photo && data.photo.toString().indexOf('http') === 0) {
    photoLink = '<' + data.photo + '|View>';
  }

  var messageText = '*QA Escalation* <!subteam^' + CONFIG.QA_SUBTEAM_ID + '>\n\n' +
    '*Order:* `' + data.orderNumber + '`  |  *Lot:* `' + data.lotNumber + '`\n' +
    '*Complaint:* ' + data.complaint + '\n' +
    '*Resolution:* ' + data.resolution + '\n' +
    '*Photo:* ' + photoLink;

  var message = {
    text: 'QA Escalation - Order ' + data.orderNumber,
    blocks: [
      { type: 'section', text: { type: 'mrkdwn', text: messageText } },
      { type: 'context', elements: [{ type: 'mrkdwn', text: 'Row ' + row + ' | Product Quality & Safety Log' }] }
    ]
  };

  try {
    var response = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(message),
      muteHttpExceptions: true
    });
    return response.getResponseCode() === 200;
  } catch (error) {
    Logger.log('Slack error: ' + error.message);
    return false;
  }
}

// ============ ERROR ALERT ============
function sendErrorSlackAlert(systemName, errorMessage) {
  var webhookUrl = getSlackWebhookUrl_();
  if (!webhookUrl) {
    Logger.log('Slack webhook not configured. Set SLACK_WEBHOOK_URL in Script Properties.');
    return;
  }

  var message = {
    text: '🔴 System Error: ' + systemName,
    blocks: [
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: '*🔴 System Error*\n\n*System:* ' + systemName + '\n*Error:* ' + errorMessage + '\n*Time:* ' + new Date().toLocaleString()
        }
      }
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
    Logger.log('Error alert failed: ' + e.message);
  }
}
