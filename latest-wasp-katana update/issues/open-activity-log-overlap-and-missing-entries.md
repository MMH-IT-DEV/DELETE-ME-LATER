# OPEN — Activity Log: Overlapping Entries + Missing Entries

## Status
INVESTIGATING — 2026-03-19

## Two Related Problems

### Problem 1: Overlapping entries (stacking)
When multiple webhooks fire simultaneously (F5 batch prints, concurrent F1/F3/F4),
Activity log entries overlap — header rows from different entries interleave with
each other's sub-rows. This happens because `logActivity` acquires and releases
`LockService.getScriptLock()` per entry, allowing other flows to slip in between.

Visible in: F5 shipping entries at 14:20-14:22 on 2026-03-19.

### Problem 2: Missing entries (logActivity silently fails)
MO-7417 processed successfully (webhook queue shows "processed", WASP calls made,
snapshot saved) but NO Activity log entry was written.

Visible in: MO-7417 at 15:12 on 2026-03-19.

## What We Tried

### Attempt 1: Switch to document lock (@415)
Changed `runWithScriptWriteLock_` from `LockService.getScriptLock()` to
`LockService.getDocumentLock()` to avoid contention with other operations
using the script lock.

**Result:** Likely BROKE logActivity completely. The debug script is a standalone
web app (not container-bound to the spreadsheet). `getDocumentLock()` may not
work for standalone scripts that open sheets by ID. Reverted at @416.

**Evidence:**
- All Activity entries BEFORE @415 worked fine (F1, F3, F5)
- MO-7417 processed at 15:12 (AFTER @415) — no Activity entry
- No Activity entries confirmed working between @415 and @416

### Attempt 2: Extend ss_poll_running TTL (@415)
Changed F5's `ss_poll_running` CacheService guard from 60s to 300s (5 min).

**Result:** Should help reduce concurrent processWebhookQueue runs. Still in place.

## Root Cause Analysis

### Why entries overlap:
1. `logActivity` uses `runWithScriptWriteLock_` with `getScriptLock()`
2. Each `logActivity` call acquires lock → writes header + sub-rows → releases lock
3. Between two entries from the same flow (e.g., two F5 shipments in a loop),
   another concurrent execution (different flow or SHIP_NOTIFY webhook) can
   acquire the lock and write ITS entry in between
4. Result: entries from different sources interleave

### Why entries go missing:
1. `logActivity` has a try-catch that returns 'WK-ERR' on any exception
2. If the lock acquisition fails (20s timeout), nothing is written
3. The calling handler doesn't know logActivity failed — it continues and
   returns { status: 'processed' }
4. The MO's snapshot is saved, blocking retries, but no Activity entry exists

## What We Need to Investigate

1. **Can `getDocumentLock()` work in standalone scripts?**
   Test: create a simple function in the debug script that tries
   `LockService.getDocumentLock().waitLock(5000)` and logs the result.

2. **Is `getScriptLock()` contention the real cause of overlaps?**
   Test: add timing logs to `runWithScriptWriteLock_` to measure how long
   the lock is held and how long callers wait.

3. **Would `getUserLock()` be a better alternative?**
   `getUserLock()` is per-user and might avoid contention with the script lock
   used by other operations. But web app requests might all run as the same user.

4. **Should logActivity failures propagate to the caller?**
   Currently the handler doesn't know logActivity failed. If it did, it could
   avoid saving the snapshot (so the MO can be retried).

5. **Is the 5-minute cache dedup (`mo_done_`) too aggressive?**
   It's set BEFORE processing. If the handler fails partway through, the dedup
   blocks retries for 5 minutes. Consider setting it AFTER success only.

## Potential Fixes (Not Yet Implemented)

### Fix A: Hold lock for entire batch (within one flow)
In `processWebhookQueue`, acquire the lock ONCE for the entire shipment loop
instead of per-shipment. Pro: no interleaving within F5. Con: blocks other
flows for the entire batch duration.

### Fix B: Queue-based logging
Instead of writing directly to the sheet, push entries to a queue (CacheService
or PropertiesService). A separate function drains the queue and writes entries
sequentially. Pro: completely eliminates contention. Con: adds complexity and delay.

### Fix C: logActivity returns success/failure to caller
If logActivity fails, the handler should NOT save the snapshot/state.
This allows the webhook to be retried and the entry to be written on retry.
Pro: self-healing. Con: might cause duplicate WASP calls on retry.

### Fix D: Retry logActivity internally
If the lock times out, retry once after a short delay instead of giving up.
Pro: simple. Con: adds latency.

## Current State (@416)

- `runWithScriptWriteLock_` uses `getScriptLock()` (reverted from document lock)
- `ss_poll_running` TTL = 300s (5 min)
- F5 ship dedup: `cache.put(cacheKey, 'processing', 86400)` before WASP calls
- F5 void dedup: same early cache set
- Overlap issue: NOT fixed (same as before @415)
- Missing entries: fixed by reverting document lock (needs confirmation)

## Key Files

- `debug/script/08_Logging.js` — `runWithScriptWriteLock_`, `logActivity`, `appendActivityBlockToSheet_`
- `debug/script/06_ShipStation.js` — `processWebhookQueue`, `processShipment`, `processVoid`
- `debug/script/02_Handlers.js` — `handleManufacturingOrderDone`, all flow handlers that call logActivity

## Timeline

| Time | Event |
|------|-------|
| 14:20-14:22 | F5 entries overlap (stacking visible) |
| ~14:30 | Deployed @415: document lock + 5min poll guard |
| 15:12 | MO-7417 processed but NO Activity entry |
| ~15:25 | Deployed @416: reverted document lock |
| Pending | Need to test MO-7417 again to confirm script lock works |
