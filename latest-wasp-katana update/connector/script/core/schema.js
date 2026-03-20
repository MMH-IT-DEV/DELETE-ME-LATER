var GMP_CONFIG = {
  SCHEMA_VERSION: '2026-03-06-review-cycle',
  SCHEMA_VERSION_PROPERTY: 'GMP_SECURITY_SCHEMA_VERSION',
  BOT_TOKEN_PROPERTY: 'SLACK_BOT_TOKEN',
  VERIFICATION_TOKEN_PROPERTY: 'SLACK_VERIFICATION_TOKEN',
  ACCESS_REVIEW_CHANNEL_PROPERTY: 'SLACK_ACCESS_REVIEW_CHANNEL',
  SUPPORT_CHANNEL_PROPERTY: 'SLACK_SUPPORT_CHANNEL',
  HEALTH_MONITOR_URL_PROPERTY: 'HEALTH_MONITOR_URL',
  SHOPIFY_TOKEN_PROPERTY: 'SHOPIFY_TOKEN',
  SHOPIFY_STORE_PROPERTY: 'SHOPIFY_STORE',
  KATANA_API_KEY_PROPERTY: 'KATANA_API_KEY',
  DEFAULT_ACCESS_REVIEW_CHANNEL: '#it-access-review',
  DEFAULT_SUPPORT_CHANNEL: '#it-support',
  DAILY_TRIGGER_HOUR: 8,
  REVIEW_SUMMARY_LIMIT: 10,
  ISSUE_COMMAND_GROUP_PROPERTY: 'SLACK_ISSUE_COMMAND_GROUP_ID'
};

function buildColumnMap_(keys) {
  var map = {};
  for (var i = 0; i < keys.length; i++) {
    map[keys[i]] = i + 1;
  }
  return map;
}

var REQUEST_KEYS = [
  // Visible columns (1-11)
  'REQUEST_ID',
  'SUBMITTED_AT',
  'REQUEST_TYPE',
  'SLACK_USER',
  'TARGET_USER',
  'COMPANY_EMAIL',
  'GMP_SYSTEM',
  'ACCESS_LEVEL',
  'REASON',
  'APPROVAL',
  'APPROVED_BY',
  // Hidden columns (12-25) — used by automation
  'QUEUE_PRIORITY',
  'DEPT',
  'GMP_IMPACT',
  'MANAGER',
  'IT_OWNER',
  'PROVISIONING',
  'NEXT_SLACK_TRIGGER',
  'SLA_DUE',
  'DECISION_DATE',
  'ACCESS_ID',
  'LINKED_REVIEW_ID',
  'LINKED_INCIDENT_ID',
  'SLACK_THREAD',
  'NOTES'
];

var ACCOUNT_KEYS = [
  'ACCESS_ID',
  'GMP_SYSTEM',
  'PERSON',
  'COMPANY_EMAIL',
  'PLATFORM_ACCOUNT',
  'ACCESS_LEVEL',
  'CURRENT_STATE',
  'MFA',
  'LAST_VERIFIED_AT',
  'LAST_REVIEW_DATE',
  'NEXT_REVIEW_DUE',
];

var REVIEW_KEYS = [
  'REVIEW_ID',
  'REVIEW_CYCLE',
  'REVIEW_PRIORITY',
  'PERSON',
  'SYSTEM',
  'ACCESS_LEVEL',
  'PRESENCE_STATUS',
  'ACTIVITY_STATUS',
  'WHY_FLAGGED',
  'REVIEWER_ACTION',
  'DECISION_STATUS',
  'DECISION_DATE',
  'NEXT_REVIEW_DUE',
  'NOTES',
  'TRIGGER_TYPE',
  'ACCESS_ID',
  'LINKED_REQUEST',
  'SLACK_THREAD',
  'NEXT_SLACK_TRIGGER',
  'LAST_LOGIN',
  'LOGINS_30D',
  'DAYS_SINCE_LOGIN',
  'ACTIVITY_SCORE',
  'COMPANY_EMAIL',
  'REVIEWED_BY'
];

var INCIDENT_KEYS = [
  'INCIDENT_ID',
  'REPORTED_AT',
  'SEVERITY',
  'SYSTEM',
  'SLACK_USER',
  'SUMMARY',
  'ASSIGNED_TO',
  'STATUS',
  'RESPONSE_TARGET',
  'RESOLUTION',
  'RESOLVED_AT',
  'NOTES',
  'ISSUE_TYPE',
  'LINKED_ACCESS_ID',
  'LINKED_REQUEST_ID',
  'ROOT_CAUSE',
  'SLACK_THREAD',
  'NEXT_SLACK_TRIGGER'
];

var GMP_SCHEMA = {
  tabs: {
    ACCESS_CONTROL: 'Access Control',
    REVIEW_CHECKS: 'Review Checks',
    INCIDENTS: 'Incidents'
  },
  sectionLabels: {
    REQUESTS: 'ACCESS REQUESTS',
    ACCOUNTS: 'ACTIVE SYSTEM ACCOUNTS',
    ARCHIVED_REQUESTS: 'ARCHIVED REQUESTS',
    REVIEWS: 'REVIEW CHECKS',
    OPEN_REVIEWS: 'NEEDS REVIEW',
    WAITING_REVIEWS: 'UPCOMING REVIEWS',
    PENDING_REMOVAL: 'PENDING REMOVAL',
    COMPLETED_REVIEWS: 'COMPLETED REVIEWS',
    INCIDENTS: 'INCIDENT INTAKE'
  },
  layout: {
    TITLE_ROW: 1,
    NOTES_ROW: 2,
    REQUEST_SECTION_ROW: 4,
    REQUEST_HEADER_ROW: 5,
    ACCOUNT_SECTION_ROW: 7,
    ACCOUNT_HEADER_ROW: 8,
    ARCHIVE_SECTION_ROW: 10,
    ARCHIVE_HEADER_ROW: 11,
    REVIEW_SECTION_ROW: 4,
    REVIEW_HEADER_ROW: 5,
    REVIEW_CONTROL_ROW: 6,
    INCIDENT_SECTION_ROW: 4,
    INCIDENT_HEADER_ROW: 5
  },
  requests: {
    keys: REQUEST_KEYS,
    headers: [
      'Request ID',
      'Submitted At',
      'Request Type',
      'Requested By',
      'Target User',
      'Company Email',
      'System',
      'Access Level',
      'Reason',
      'Approval',
      'Approved By',
      'Queue Priority',
      'Dept',
      'System Impact',
      'Manager',
      'IT Owner',
      'Access Setup Status',
      'Next Slack Trigger',
      'SLA Due',
      'Decision Date',
      'Access ID',
      'Linked Review ID',
      'Linked Incident ID',
      'Slack Thread',
      'Notes'
    ],
    columns: buildColumnMap_(REQUEST_KEYS),
    widths: [16, 18, 14, 14, 18, 22, 18, 14, 34, 14, 16, 12, 14, 11, 14, 14, 16, 22, 18, 18, 16, 18, 18, 22, 34],
    dropdowns: [
      { key: 'REQUEST_TYPE', values: ['New Access', 'Change Access', 'Remove Access'] },
      { key: 'ACCESS_LEVEL', values: ['Admin', 'Full Access', 'Read-Write', 'Read-Only', 'Limited', 'Remove'] },
      { key: 'APPROVAL', values: ['Submitted', 'Approved', 'Denied', 'Escalated'] },
      { key: 'QUEUE_PRIORITY', values: ['High', 'Medium', 'Low'] },
      { key: 'GMP_IMPACT', values: ['Yes', 'No'] },
      { key: 'PROVISIONING', values: ['Open', 'In Progress', 'Completed', 'Removing Access', 'Closed'] }
    ]
  },
  accounts: {
    keys: ACCOUNT_KEYS,
    headers: [
      'Access ID',
      'System',
      'Person',
      'Company Email',
      'Platform Account',
      'Access Level',
      'Access Status',
      'MFA',
      'Last Verified',
      'Last Reviewed',
      'Next Review Due'
    ],
    columns: buildColumnMap_(ACCOUNT_KEYS),
    widths: [16, 18, 18, 24, 24, 14, 16, 10, 16, 16, 16],
    dropdowns: [
      { key: 'ACCESS_LEVEL', values: ['Admin', 'Full Access', 'Read-Write', 'Read-Only', 'Limited'] },
      { key: 'CURRENT_STATE', values: ['Setting Up', 'Access Granted', 'Unmanaged', 'Revoked'] },
      { key: 'MFA', values: ['Yes', 'No'] }
    ]
  },
  reviews: {
    keys: REVIEW_KEYS,
    headers: [
      'Review ID',
      'Review Cycle',
      'Priority',
      'Person',
      'System',
      'Access Level',
      'Access Status',
      'Activity Status',
      'Check Needed',
      'Reviewer Action',
      'Review Status',
      'Reviewed At',
      'Next Check Date',
      'Notes',
      'Trigger Type',
      'Access ID',
      'Linked Request',
      'Slack Thread',
      'Next Slack Trigger',
      'Last Activity Evidence',
      '30d Logins',
      'Days Since Login',
      'Activity Score',
      'Company Email',
      'Reviewed By'
    ],
    columns: buildColumnMap_(REVIEW_KEYS),
    widths: [14, 12, 10, 16, 16, 12, 12, 14, 22, 14, 14, 16, 18, 22, 16, 14, 18, 18, 20, 16, 10, 14, 12, 22, 20],
    dropdowns: [
      { key: 'REVIEW_PRIORITY', values: ['High', 'Medium', 'Low'] },
      { key: 'ACCESS_LEVEL', values: ['Admin', 'Full Access', 'Read-Write', 'Read-Only', 'Limited'] },
      { key: 'REVIEWER_ACTION', values: ['Keep', 'Reduce', 'Remove', 'Need Info'] },
      { key: 'DECISION_STATUS', values: ['Open', 'Waiting', 'Done'] }
    ]
  },
  incidents: {
    keys: INCIDENT_KEYS,
    headers: [
      'Incident ID',
      'Reported At',
      'Severity',
      'System',
      'Reported By',
      'Summary',
      'Assigned To',
      'Status',
      'Response Target',
      'Resolution',
      'Resolved At',
      'Notes',
      'Issue Type',
      'Linked Access ID',
      'Linked Request ID',
      'Root Cause',
      'Slack Thread',
      'Next Slack Trigger'
    ],
    columns: buildColumnMap_(INCIDENT_KEYS),
    widths: [14, 18, 10, 22, 14, 34, 16, 14, 18, 28, 18, 34, 18, 16, 16, 24, 22, 24],
    dropdowns: [
      { key: 'SEVERITY', values: ['High', 'Medium', 'Low'] },
      { key: 'STATUS', values: ['Open', 'Investigating', 'Resolved'] }
    ]
  },
  platforms: ['Katana', 'Shopify', 'ShipStation', '1Password', 'Wasp', 'Other']
};
