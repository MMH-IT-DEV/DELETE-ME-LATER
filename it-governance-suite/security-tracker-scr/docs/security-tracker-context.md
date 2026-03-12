# 2026 Security & Connection Tracker — Context

## What it is
A Google Sheets-based IT access tracking system for GMP compliance.
Tracks user access across systems, monthly reviews, and incidents.

## Google Sheet
**Sheet ID:** `1WBzed_RMtPC3kii-ybJ3lE4Fo1pAqWBt9OJyCSVR-mQ`
**Sheet Name:** 2026_Security & Connection Tracker

## Apps Script Project
**Script ID:** `1Fp0ooeKm028-0XYu5X5CCFrxzILbDEPxwX6Si_I7c3A3GcwmEnANNb4k`
**Local files:** `C:\Users\Admin\Documents\claude-projects\wasp-katana\it-governance-suite\security-tracker-scr\`
**Push command:** `cd it-governance-suite/security-tracker-scr && clasp push`
**Web App URL:** `https://script.google.com/macros/s/AKfycbymfMn3C7OZYM4HfD-06gROylesHMUhXpVIwwjlUvjhhH2hcjPM2xe23XktKEUL-rVOCA/exec`

## Slack App
**App:** IT Support | **Workspace:** MyMagicHealer (MMH)
**Bot Token:** stored as Script Property `SLACK_BOT_TOKEN`
**Slash command:** `/gmp-request-access`
**Channel:** `#it-support`

## Script Properties (in GAS editor)
| Property | Value |
|----------|-------|
| SLACK_WEBHOOK_URL | hooks.slack.com/... |
| SLACK_BOT_TOKEN | xoxb-... |
| SHEET_ID | 1WBzed_RMtPC3kii-ybJ3lE4Fo1pAqWBt9OJyCSVR-mQ |
| HEALTH_TRACKER_URL | https://script.google.com/macros/s/AKfycb... |

## Files in security-tracker-scr/
| File | Purpose |
|------|---------|
| `slack-notifications.js` | onEdit trigger → Slack notification when Status = Pending |
| `slash-command.js` | `/gmp-request-access` slash command → modal → writes to sheet |
| `sheet-setup.js` | Run `setupSheet()` once to apply headers, colors, dropdowns, date pickers |
| `active-directory-sync.js` | AD sync (pulled from original script) |
| `migrate-log-ad.js` | Migration helper (pulled from original script) |

Current local folder structure:

| Path | Purpose |
|------|---------|
| `slack/slack-notifications.js` | Pending-status Slack notification |
| `slack/slash-command.js` | Slash command intake modal |
| `sheet/sheet-setup.js` | Sheet formatting setup |
| `integrations/active-directory-sync.js` | Active Directory sync |
| `integrations/migrate-log-ad.js` | Migration helper |
| `shopify-sync.js` | Connected platform refresh logic |

## Three Tabs + Column Layout

### Tab 1 — Log (19 columns)
`Current Access | MFA Enabled | Last Review | Date | Person | Email | Role/Dept | Type | Service/Platform | Access Level | Access Reason | Approved By | Status ⚡(col 13) | Expires | Notes | Request By | Date Removed | Username/Account ID | Employment Status`

### Tab 2 — Reviews (15 columns)
`Review ID | Date | Reviewer | Review Type | Systems Reviewed | Access OK | Security OK | Findings | Actions Taken | Sign-off | Next Review ⚡(col 11) | Review Scope | Status | Completed Date | Users Reviewed`

### Tab 3 — Incidents (16 columns)
`Incident ID | Date | Reported By | Type | Severity | Description | Containment Actions | Root Cause | Corrective Actions | Status ⚡(col 10) | Closed Date | Policy Violation | Assigned To | Systems Affected | Impact | Detection Method`

## Slack Notification Triggers
- **Log tab:** Status (col 13) = `Pending` → fires notification
- **Reviews tab:** Next Review (col 11) = within 7 days → daily 8AM check
- **Incidents tab:** Status (col 10) changes → fires notification

## Slack Command Flow
1. User types `/gmp-request-access` in any Slack chat
2. Modal opens with 9 fields (Full Name + Email pre-populated from Slack profile)
3. Submit → row appended to Log tab with Status = Pending
4. Slack notification fires immediately (not via onEdit — called directly from script)

## Warmup Trigger
`setupWarmupTrigger()` was run — pings script every 5 min to prevent GAS cold start timeouts.

## Pending Work (to return to)
- Collect existing users from Shopify, ShipStation, Katana and populate the Log tab
- Felippe needs to provide Shopify user list (no admin access)
- Run `setupSheet()` if headers/colors/dropdowns need to be applied to the sheet
- Test full end-to-end: submit form → row in sheet → Slack notification received
