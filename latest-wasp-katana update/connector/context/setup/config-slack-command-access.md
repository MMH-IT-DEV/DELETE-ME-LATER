# Slack Command Access Control
**Last updated:** 2026-03-18
**Status:** Active

---

## Restricted Commands

| Command | Purpose |
|---------|---------|
| `/system-issue` | Report a system or workflow issue |
| `/it-alert` | Alias for `/system-issue` |

Anyone not on the access list gets a private "You don't have access" message when they try to use these commands. All other commands remain open to everyone.

---

## Who Has Access

| Name | Email | Role |
|------|-------|------|
| Erik Demchuk | erik@mymagichealer.com | IT |
| Robyn Bepple | robyn@mymagichealer.com | Logistics & Warehouse Supervisor |
| Felippe Fernandes | felippe@mymagichealer.com | General Manager |

Access is managed through the Slack user group **@management**.
To add or remove someone: go to Slack → People & User Groups → Management → Edit Members.
No code changes or redeployment needed.

---

## Configuration

| Setting | Value |
|---------|-------|
| Slack workspace | MyMagicHealer (MMH) |
| User group | `@management` |
| Group ID | `S0AN90UPE4Q` |
| Script property key | `SLACK_ISSUE_COMMAND_GROUP_ID` |
| Script property value | `S0AN90UPE4Q` |
| Slack bot scope | `usergroups:read` |
| Slack app | IT Support |
| Apps Script project | Security & System Connection Tracker (connector) |

---

## How It Works

When `/system-issue` or `/it-alert` is triggered, the script calls the Slack API to check if the user is a member of group `S0AN90UPE4Q`. If yes, the modal opens. If no, a private error message is shown only to that user.

**Fail-open:** If the Slack API is unreachable or the group property is not set, the command stays open to everyone so no one gets accidentally locked out.

---

## To Change the Access List

**Add someone:** Slack → People & User Groups → Management → Edit → add member → Save.
**Remove someone:** same path → remove member → Save.
**Replace the group entirely:** Update `SLACK_ISSUE_COMMAND_GROUP_ID` in Apps Script Project Settings → Script Properties with the new group's ID.
