# Katana PO Calendar Sync — Maintenance Guide

## What It Does
Fetches open Purchase Orders from Katana MRP and creates/updates all-day events
in Google Calendar. Runs daily on a time-driven trigger.

- **Creates** a calendar event for each open PO (status: Not Received or Partially Received)
- **Updates** events when the expected arrival date changes
- **Removes** events for POs that have been fully received
- Looks ahead **90 days** from today

---

## Key Links

| Resource | Link |
|----------|------|
| **Script editor** | https://script.google.com/home/projects/1mjraAHwRu5Nfel1yRzjuTK8I4uZPrl_QyUsYnZ67Bkmp-GTLFdUEMC1o/edit |
| **Drive project folder** | https://drive.google.com/drive/folders/190tw277_odKMPwKdO5CsFNpHFKX9tEwT |
| **Health Tracker** | https://docs.google.com/spreadsheets/d/1jnWtdBPzR7DreihCHQASiN7spImRS7HYjTR77Gpfi5w/edit |
| **Katana MRP** | https://katanamrp.com |

---

## Script Properties (GAS editor → Project Settings → Script Properties)

| Property | Required | What it is |
|----------|----------|-----------|
| `KATANA_API_KEY` | Yes | Katana API key — get from Katana → Settings → API |
| `CALENDAR_ID` | Yes | Google Calendar ID to write PO events to |
| `HEALTH_MONITOR_URL` | Yes | Systems Health Monitor web app URL |
| `SLACK_WEBHOOK_URL` | No | Legacy — not used by active code |

To find the `CALENDAR_ID`: open Google Calendar → Settings → select the calendar → scroll to "Calendar ID" (looks like `abc123@group.calendar.google.com`)

---

## Trigger

**Function:** `syncPurchaseOrdersToCalendar`
**Type:** Time-driven — runs daily
**Reinstall:** GAS editor → Triggers (clock icon) → Add Trigger → select `syncPurchaseOrdersToCalendar` → Time-driven → Day timer

---

## How to Read the Calendar Events

Each event is titled: `🚚 PO-492 • SUPPLIER NAME`
Description includes: PO number, status, item list with quantities.
Hidden in the description: `<!-- PO_ID:12345 -->` — used by the script to match events to POs.

---

## Common Issues

### Events not appearing / sync not running
Trigger is missing or deleted.
GAS editor → Triggers → Add Trigger → `syncPurchaseOrdersToCalendar` → Time-driven → Day timer.

### `KATANA_API_KEY not set` error in logs
Script Property is missing.
GAS editor → Project Settings → Script Properties → add `KATANA_API_KEY`.

### Supplier shows as "Supplier #12345" instead of name
Supplier ID is not in the `SUPPLIER_MAP` in `getConfig()` and the API lookup failed.
Run `listSupplierIDs()` in the script editor to get current supplier IDs and names, then update `SUPPLIER_MAP`.

### `Calendar not found` error
`CALENDAR_ID` Script Property is wrong or the calendar was deleted.
Verify the calendar ID in Google Calendar settings and update the Script Property.

### Old verbose events not updating
Events in the old format (containing `═══`) get auto-updated on next sync — no manual action needed.

---

## Run Manually

Open GAS editor → select `syncPurchaseOrdersToCalendar` → Run.
Check execution log for results (created / updated / skipped / removed counts).

## Test Functions (run from GAS editor)

| Function | What it does |
|----------|-------------|
| `testFetchPurchaseOrders` | Lists open POs from Katana API |
| `testCalendarAccess` | Confirms calendar ID is valid |
| `listSupplierIDs` | Prints all supplier IDs from open POs (use to update SUPPLIER_MAP) |
| `deleteAllPOEvents` | ⚠️ Deletes all PO events from calendar — use with caution |
