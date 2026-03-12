# GMP System Access Request Flow
**Project**: 2026_Security & GMP Connection Tracker
**Sheet**: https://docs.google.com/spreadsheets/d/1WBzed_RMtPC3kii-ybJ3lE4Fo1pAqWBt9OJyCSVR-mQ/edit
**Script**: https://script.google.com/u/0/home/projects/1Fp0ooeKm028-0XYu5X5CCFrxzILbDEPxwX6Si_I7c3A3GcwmEnANNb4k/edit
**Health Monitor**: https://docs.google.com/spreadsheets/d/1jnWtdBPzR7DreihCHQASiN7spImRS7HYjTR77Gpfi5w/edit

---

## Overview

End-to-end process for granting, recording, and maintaining GMP system access.
Covers: request → approval → account creation → active monitoring → revocation.

---

## Phase 1 — Access Request Submitted

**Who**: Employee, manager, or onboarding coordinator
**Where**: Access Requests tab in the Security & GMP Connection Tracker sheet

### What gets filled in (columns):
| Column | Field | Notes |
|--------|-------|-------|
| A | Date | Auto or manual |
| B | Submitted Date | When row was added |
| C | Requested By | Person submitting |
| D | User (Full Name) | Person needing access |
| E | Email | Their work email |
| F | System / Platform | e.g. Katana, Shopify, WASP, NetSuite |
| G | Access Level | e.g. View, Edit, Admin |
| H | Reason / Justification | Why they need it |
| I | Status | Set to **Pending** to trigger automation |
| J | Approved By | Filled in on approval |
| K | Approval Date | Filled in on approval |
| L | Notes | Any context for IT |
| M | Synced to Register | Auto-filled "Yes" on approval |

### Trigger (automated):
When column I (Status) is changed to `Pending`:
- **Slack alert fires** → `#it-alerts` channel
- Message includes: User, System, Access Level, Requested By, Reason
- Link to sheet for reviewer to act on
- Function: `handleAccessRequestStatus()` in `it-security-tracker.js`

### Daily staleness check (automated):
- `checkPendingRequests()` runs at 9 AM daily
- Flags any request still `Pending` after 24h with a Slack alert
- So nothing gets forgotten in the queue

---

## Phase 2 — IT Review & Decision

**Who**: IT person or designated approver
**Where**: Access Requests tab — update Status column

### Decision options:

#### Option A — Approve
1. Set Status → **Approved**
2. Fill in: Approved By (your name), Approval Date
3. **What happens automatically**:
   - Slack notification: "Access Approved" → User, System, Level
   - Row is copied to **Access Register** tab (`syncToAccessRegister()`)
   - Column M "Synced to Register" → marked `Yes`
   - Health Monitor logs the event (automation ID: SEC-004)

#### Option B — Deny
1. Set Status → **Denied**
2. Add reason in Notes column
3. **What happens automatically**:
   - Slack notification: "Access Denied" → User, System
   - Health Monitor logs the event (automation ID: SEC-005)
   - No entry written to Access Register

---

## Phase 3 — Account Creation

**Who**: IT person (with or alongside the user)
**When**: After approval, before or during onboarding to the system

### Steps:
1. Log into the target system (Katana, Shopify, WASP, etc.)
2. Create the account using the user's work email
3. Set access level to match what was approved
4. Enable MFA if required by system policy
5. Send credentials / invite link to the user
6. Document the account creation in the **Log tab**

### What to record in the Log tab:
| Field | Value |
|-------|-------|
| Date | Today |
| Person | User's full name |
| Email | Work email |
| Role / Dept | Their department |
| Type | `Access Grant` |
| Service / Platform | System name |
| Access Level | As approved |
| Access Reason | From the request |
| Approved By | Your name |
| Status | `Complete` |
| Username | Their login (if applicable) |
| Employment Status | Active |

When Status is set to `Complete` on an Access Grant:
- `handleLogADSync()` triggers → runs `syncAccessStatus()`
- **Current Access** column on that log row → set to `Active`
- Green chip applied to the cell

---

## Phase 4 — Access Register (Live Record)

**Tab**: Access Register
**Purpose**: Single source of truth for who has active access to what

Auto-populated when access is approved (via `syncToAccessRegister()`):

| Column | Data |
|--------|------|
| A | User Full Name |
| B | Email |
| C | System |
| D | Access Level |
| E | Status (`Active`) |
| F | Approval Date |
| G | Approved By |
| H | Reason |
| I | MFA Enabled — fill in manually |
| J | Last Login — update periodically |
| K | Next Review Date — set when created |
| L | Notes |

### MFA policy:
- Mark MFA Enabled: Yes/No once confirmed with the user
- If No → follow up with the user to enable it before access goes live (recommended)

### Review schedule:
- Set **Next Review Date** at time of provisioning
- Recommended cadence: 90 days for Admin, 180 days for Edit/View
- Overdue reviews are flagged daily at 9 AM via `checkOverdueReviews()`

---

## Phase 5 — Ongoing Account Monitoring

### Periodic Access Reviews (scheduled)
- Daily at 9 AM: `checkOverdueReviews()` scans the Reviews tab
- Sends Slack alert when Next Review Date has passed
- IT follows up with: Is this person still active? Do they still need this access?

### Log Tab (activity history)
- Every significant change gets a new Log row:
  - `Access Grant` — initial provisioning
  - `Access Modify` — level change (e.g. View → Edit)
  - `Access Revoke` — removal
  - `Review` — periodic access review completed
- `refreshLogView()` groups Log tab by system with colored section headers
  - Run from menu: IT Notifications → Refresh Log View

### Refresh Log View output:
Each system gets a colored section header showing:
```
Katana   (4 accounts  ·  3 active  ·  1 inactive)
```
Data rows show access status color-coded:
- Green chip = Active
- Red chip = Revoked / Denied
- Orange chip = Pending
- Blue chip = In Progress

---

## Phase 6 — Access Revocation

**Triggers**: Employee offboarding, role change, project end, periodic review failure

### Steps:
1. Add a new row to the **Log tab**:
   - Type: `Access Revoke`
   - Status: `Complete`
2. Update the **Access Register** row:
   - Status → `Revoked`
   - Date Removed column → today
3. Log into the target system and disable/delete the account
4. Confirm MFA devices are removed if applicable

### What happens automatically:
- `handleLogADSync()` fires on Status = `Complete`
- `syncAccessStatus()` derives status as `Revoked`
- **Current Access** cell → updated to `Revoked` (red chip)

---

## Summary: Who Does What

| Step | Actor | Where | Automated? |
|------|-------|--------|------------|
| Submit request | Employee / Manager | Access Requests tab | Slack alert on Pending |
| Review & decide | IT Person | Access Requests tab | Slack on Approved/Denied |
| Create account | IT Person | Target system | No — manual |
| Record in Log | IT Person | Log tab | Access status auto-synced |
| Monitor reviews | IT Person | Reviews tab | Slack alert on overdue |
| Revoke access | IT Person | Log + Register + system | Access status auto-synced |

---

## Automation IDs (Health Monitor)

| ID | Event |
|----|-------|
| SEC-001 | Incident Opened |
| SEC-002 | Incident Closed |
| SEC-003 | Access Request Pending |
| SEC-004 | Access Approved |
| SEC-005 | Access Denied |
| SEC-006 | Overdue Review flagged |
| SEC-007 | Stale request (>24h Pending) |
| SEC-HB | Daily heartbeat |

---

## Key Files

| File | Purpose |
|------|---------|
| `systems-health-scr/it-security-tracker.js` | Health monitor bridge, tracker logging, and shared security helpers |
| `security-tracker-scr/integrations/active-directory-sync.js` | Access status sync, Log refresh, menu |
| `security-tracker-scr/shopify-sync.js` | refreshLogView(), dashboard refresh |
| `security-tracker-scr/slack/slack-notifications.js` | Slack sendHeartbeat, CONFIG columns |
| `security-tracker-scr/slack/slash-command.js` | `/gmp-request-access` modal and request intake |
