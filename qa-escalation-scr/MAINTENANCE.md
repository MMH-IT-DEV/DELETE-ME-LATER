# QA Escalation Bot — Maintenance Guide

## What It Does
Watches the **Product Quality & Safety Log** sheet. When the `QA Escalation`
column is set to **"Escalate to QA"**, it sends a Slack message to `#qa-ops`
tagging the QA subteam, then marks `QA Sent` as ✓ or ✗.

---

## Key Links

| Resource | Link |
|----------|------|
| **Script editor** | https://script.google.com/d/10Jdnp-PuMVZP_p8teerW5Mr4tkHiGfv-i9Twfql-RvnFzmH0wEBZs4g0/edit |
| **Google Sheet** | https://docs.google.com/spreadsheets/d/1Jz1PLVPPtLQxWHG-6vboT1cL-vmylQl5N20W-ORPDVY/edit |
| **Drive project folder** | https://drive.google.com/drive/folders/1YkMQvwUwhLn6Bo5rJfTBY0iXRdt3LVBU |
| **Health Tracker** | https://docs.google.com/spreadsheets/d/1jnWtdBPzR7DreihCHQASiN7spImRS7HYjTR77Gpfi5w/edit |

---

## Slack

| | |
|-|-|
| **Channel** | `#qa-ops` |
| **Channel ID** | `C0AAFBN0YBC` |
| **Who gets tagged** | QA subteam — ID `S0ABE9GKD7T` |
| **Webhook URL** | Script Property `SLACK_WEBHOOK_URL` |

---

## How the Trigger Works

1. User opens the **Product Quality & Safety Log** sheet
2. Sets the `QA Escalation` column on any row to **"Escalate to QA"**
3. Bot reads: Order #, Lot #, Complaint, Resolution, Photo link from that row
4. Sends a Slack message to `#qa-ops` tagging `@qa-team`
5. Writes ✓ (success) or ✗ (failed) to the `QA Sent` column

Header row is **row 10**. Columns are found by name — safe to reorder them.

---

## Script Properties (GAS editor → Project Settings → Script Properties)

| Property | Required | What it is |
|----------|----------|-----------|
| `HEALTH_MONITOR_URL` | Yes | Systems Health Monitor web app URL |
| `BACKUP_FILE_ID` | Auto-set | Set after first backup run — do not edit |

---

## Common Issues

### Bot doesn't fire when "Escalate to QA" is set
Trigger is missing. Open the sheet → **QA Escalation menu → Setup Trigger** → approve authorization.

### Slack message not received / QA Sent shows ✗
Webhook is broken or revoked.
1. Generate new webhook in Slack App settings
2. Update Script Property `SLACK_WEBHOOK_URL`
3. Push: `cd qa-escalation-scr && clasp push`
4. Test: **QA Escalation menu → Test Slack**

### Row already ✓ but team didn't get a message
Bot skips rows already marked ✓. Clear the `QA Sent` cell and re-set the dropdown.

### Column not found / wrong data in Slack message
Column header text changed in the sheet.
Update the matching value in `CONFIG.COLUMN_HEADERS` in `Code.js`, then `clasp push`.

---

## Deploy Changes

```bash
cd C:\Users\Admin\Documents\claude-projects\wasp-katana\qa-escalation-scr
clasp push
```

Then: **QA Escalation menu → Setup Trigger** to reinstall the trigger.

---

## Test Without Touching the Sheet

Open the sheet → **QA Escalation menu → Test Slack**
Sends a dummy message to `#qa-ops` with Order `TEST-12345`. Does not write to the sheet.
