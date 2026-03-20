# Inbox Auto-Deletion Script — Project Brief

## Project #23 | Owner: Felippe | Status: ✅ Ready for Review

## Summary
Automated email cleanup for **contact@mymagichealer.com**. A Google Apps Script (`purgeOldMail30d`) moves all emails older than 30 days to Trash across the entire mailbox.

## What Was Done
- Added **dry-run mode** (`DRY_RUN = true`) — logs what would be deleted without touching emails
- Enabled **starred email protection** (`PROTECT_STARRED = true`) — starred emails are kept
- No exclusion rules needed (confirmed)
- Dry-run tested successfully — 3,000 threads identified, nothing deleted
- Pushed to Google Apps Script via clasp

## Current State
- `DRY_RUN = true` — script only logs, does not delete
- `PROTECT_STARRED = true` — starred emails are safe
- No trigger set — script only runs when manually executed

## To Go Live (after Felippe confirms)
1. Flip `DRY_RUN = false`
2. Set up daily time-driven trigger
3. Push via `clasp push`

## Script Location
- **Google Apps Script**: [Project Link](https://script.google.com/home/projects/1HX-zuuiW4v4uuNq3vhLZw8l7f5bcRA3kcr6lLJIBIJgqEZbK1kESLrMu/edit)
- **Script ID**: `1HX-zuuiW4v4uuNq3vhLZw8l7f5bcRA3kcr6lLJIBIJgqEZbK1kESLrMu`

## Pending Slack Message
**#23 — Inbox Auto-Deletion Script (contact@mymagichealer.com)**

Script is ready. Dry-run tested — 3,000 threads older than 30 days identified, nothing deleted. Starred emails are protected.

Before proceeding, need to confirm:
- Is 30-day retention OK for this inbox?

Once confirmed, we flip it on and set the daily trigger.
