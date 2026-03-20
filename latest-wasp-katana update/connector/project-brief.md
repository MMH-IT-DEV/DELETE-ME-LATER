# IT Governance & Automation Control — Project Brief

## Project #20 | Owner: Felippe | Status: 🔄 In Progress

## What
A formal IT request process so no automation, script, integration, or AI tool gets built without IT approval first.

## Why
Departments are building automations on company systems without IT oversight — creating security risks, data integrity issues, and maintenance gaps.

## What Already Exists
- **Systems Health Tracker** — registers and monitors all automations (status, heartbeats, auth, owners)
- **IT Security Connector** — Slack slash commands (`/system-access`, `/system-issue`, `/it-alert`) that write to the tracker sheets
- **Slack integration** — modals, notifications, and webhook alerts already wired up

## Plan

### 1. New Slack Command: `/it-request`
- Opens a Slack modal for submitting new automation/tool/integration requests
- Fields: what they need, what systems it touches, why, urgency
- Follows the same pattern as existing `/system-access` and `/system-issue` commands
- Add to the connector script (slash-command.js)

### 2. New Sheet Tab: "IT Requests"
- New tab in the Health Tracker spreadsheet
- Columns: Date, Requested By, Description, Systems Involved, Reason, Urgency, Status (Pending/Approved/Denied), Reviewed By, Review Date, Notes
- IT reviews and updates status — Slack notification sent on status change

### 3. Policy Communication
- Announce to all departments via Slack that `/it-request` is the formal process
- No automation gets built without submitting through this command first

## Implementation
- All code changes happen in the **connector** script project
- Add new command routing to `slash-command.js`
- Add modal handler for the request form
- Add sheet writer for the new IT Requests tab
- Push via `clasp push`

## Pending Slack Message
**#20 — IT Governance & Automation Control**

We already have the Systems Health Tracker monitoring all automations, and Slack commands for system access and issue reporting.

To complete the governance process, we're adding:
- `/it-request` — a Slack command for anyone to submit new automation/tool requests
- **IT Requests** tab in the Health Tracker — to log, review, and approve/deny requests

Once live, all new automation or tool requests must go through `/it-request` before any development starts.
