# FIXED — F5: Concurrent Sheet Entries (Stacked Blue Rows)

## Status
FIXED — deployed 2026-03-18

## What Was Happening
When a user prints multiple ShipStation labels at once (batch print), ShipStation fires
one SHIP_NOTIFY webhook per label simultaneously. Each webhook triggered a separate
Google Apps Script execution, all calling processShipment() concurrently. Since
appendActivityBlockToSheet_() used getLastRow() + 1 without a lock, concurrent
executions read the same row number and wrote on top of each other — producing
multiple consecutive blue "new entry" header rows stacked together with no spacing.

Visible symptom: multiple WK-XXX blue rows at the same timestamp (e.g. 8:22) with no
gap between them, all appearing as new entries instead of being properly separated.

## Root Cause
handleShipNotify() in 06_ShipStation.js fetched the resource URL and looped through
all shipments immediately, writing each one to the Activity sheet. When 10 webhooks
fire at once, 10 GAS executions do this simultaneously — race condition on getLastRow().

## Fix Applied
File: debug/06_ShipStation.js — handleShipNotify()

Changed handleShipNotify() from ~100 lines of fetch + loop + processShipment() to 3
lines: log receipt and delegate to processWebhookQueue().

processWebhookQueue() already has a CacheService guard (ss_poll_running, 60s TTL) that
ensures only ONE execution processes at a time. Concurrent webhook calls hit the guard
and return immediately. The 1-minute time trigger on processWebhookQueue() remains as
a true safety net.

```javascript
function handleShipNotify(resourceUrl) {
  logToSheet('SS_SHIP_NOTIFY_RECEIVED', { resourceUrl: resourceUrl },
    { timestamp: new Date().toISOString() });
  return processWebhookQueue();
}
```

## Side Fix Applied — Void Race Condition
File: debug/06_ShipStation.js — processVoid()

Separate issue discovered: if a label is printed and voided within ~60 seconds (before
processWebhookQueue runs), the shipment is skipped by the poll (voidDate already set)
but pollVoidedLabels() would still try to add inventory back — over-counting WASP.

Fix: added guard in processVoid() per-item loop. When useLedgerRetryGuard is true and
ledgerLine exists but previousLedgerStatus is empty (no ledger row found = shipment was
never deducted), skip the reversal. Works for voids at any time — 4-day-old labels
still work correctly because the F5 Ledger sheet is permanent (not cache).

## Flows Affected
- F5 only (ShipStation shipment processing)
- F1, F2, F3, F4, F6 — unchanged

## Key Files
- debug/06_ShipStation.js — handleShipNotify(), processVoid(), processWebhookQueue()
- debug/16_F5Ledger.js — F5 Shipment Ledger (persistent, used for void eligibility check)

## Known Remaining Considerations
- processWebhookQueue polls by date (shipDateStart=YYYY-MM-DD) not by time, so it
  fetches all of today's shipments each run. The ss_shipped_ dedup cache (86400s TTL)
  prevents reprocessing. Poll volume increases through the day but is functionally fine.
- The ss_poll_running CacheService flag has a 60s TTL. If processing exceeds 60s, a
  second poll could start. Pre-existing issue, not introduced by this fix.
