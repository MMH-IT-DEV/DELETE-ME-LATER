# DEPLOYED — F2: Lock Timeout Causing Slack Error Alerts

## Status
DEPLOYED @406 on 2026-03-19 — monitoring for recurrence

## The Alert
Slack message from "IT Support APP" at 8:53 AM on 2026-03-18:
  System:   2026-Katana-WASP_DebugLog
  Severity: ERROR
  Error:    Lock timeout: another process was holding the lock for too long.
  Fix Guide: Not found

## What Is Happening
Google Apps Script's LockService.getScriptLock().waitLock(10000) throws
"Lock timeout: another process was holding the lock for too long" when a GAS execution
cannot acquire the script lock within 10 seconds.

The error bubbles up uncaught through the handler → routeWebhook() → doPost() catch
block → result = { status: 'error', message: 'Lock timeout...' } → heartbeat fires
→ Slack alert.

## Root Cause (Two-Part)

### Part 1 — F2 holds the script lock too long
processBatchIfReady() in 03_WaspCallouts.js acquires LockService.getScriptLock() and
holds it for the ENTIRE duration of batch processing — including every Katana API call
it makes per item. With a batch of 5–8 WASP callout items, each Katana stock
adjustment call taking 3–8 seconds, the lock can be held for 15–40+ seconds.

```javascript
// 03_WaspCallouts.js — processBatchIfReady()
var lock = LockService.getScriptLock();
try {
  lock.waitLock(10000);
} catch (e) {
  return false;  // catches its OWN timeout, but holds lock during API calls below
}
try {
  ...
  processBatchAdd(items, batchId);    // ← Katana API calls happen HERE
  // or
  processBatchRemove(items, batchId); // ← while lock is still held
  ...
} finally {
  lock.releaseLock();
}
```

### Part 2 — acquireExecutionGuard_ / releaseExecutionGuard_ don't catch the timeout
These two functions in 02_Handlers.js also use LockService.getScriptLock().waitLock(10000)
but unlike processBatchIfReady(), they do NOT catch the timeout exception. When they
time out waiting for the lock (held by F2), the exception propagates uncaught to doPost.

```javascript
// 02_Handlers.js — acquireExecutionGuard_()
function acquireExecutionGuard_(guardKey, ttlMs) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);  // ← throws on timeout, NOT caught here
  ...
}

// 02_Handlers.js — releaseExecutionGuard_()
function releaseExecutionGuard_(guardKey) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);  // ← same problem
  ...
}
```

## When It Triggers
F2 WASP callout batch arrives AND concurrent Katana webhooks (F4 MO completions, F6 SO
deliveries) arrive at the same time. F2 holds the lock while making API calls. The
concurrent F4/F6 handlers try acquireExecutionGuard_() and time out.

Confirmed from webhook queue log (2026-03-18 08:52–08:53): burst of 8–10 concurrent
SO/MO webhooks (sales_order.updated DELIVERED, manufacturing_order.done) all landing
within the same minute. F2 WASP callouts are not visible in the Katana webhook queue
log (separate entry point) but are suspected to be running concurrently.

## Evidence from Webhook Queue Log
Lines 149–225 of context/webhook-queue-log.txt show:
- 8:52:24 onward: multiple sales_order.updated (DELIVERED) logged as "pending"
  simultaneously — meaning multiple GAS executions are running concurrently
- 8:52:33: manufacturing_order.done MO-15852489 "skipped: MO completion already in
  progress" — F4 execution running
- 8:53:00–8:53:19: another burst of 5 sales_order.updated (DELIVERED) all "pending"
- 8:53:54: manufacturing_order.done MO-15852519 "skipped: MO completion already in
  progress" — another concurrent F4 execution
- Slack alert fires at 8:53

## Proposed Fix

### Fix A — Reduce time lock is held in processBatchIfReady() (the real fix)
Move Katana API calls OUTSIDE the script lock in processBatchIfReady(). The lock should
only wrap the deduplication check and the "mark as processed" cache write — not the
actual API work.

Pseudocode:
  1. Acquire lock → check if already processed → mark as processing → release lock
  2. Do ALL Katana API calls outside the lock
  3. Acquire lock → mark as done → release lock

This eliminates the 15–40 second lock hold window.

### Fix B — Catch the timeout in acquireExecutionGuard_ / releaseExecutionGuard_ (band-aid)
Add try-catch around waitLock() in both functions and handle gracefully (return false /
log a warning) instead of letting it propagate to doPost and trigger a Slack alert.

This silences the alert but does NOT fix the underlying lock contention. Other
executions would still be delayed or blocked during F2 batch processing.

### Recommended approach: Fix A + Fix B
Fix A addresses the root cause. Fix B ensures that even if contention happens for other
reasons in future, it doesn't produce noise in Slack.

## Key Files to Modify
- debug/03_WaspCallouts.js — processBatchIfReady(), processBatchAdd(), processBatchRemove()
- debug/02_Handlers.js — acquireExecutionGuard_(), releaseExecutionGuard_()

## F2 Architecture Context
F2 handles WASP callouts (quantity_added / quantity_removed events fired by WASP when
inventory changes). It uses a 10-second batch window to group concurrent callouts from
the same bulk operation into one Katana stock adjustment.

Flow:
  WASP callout → doPost → handleWaspQuantityAdded/Removed → addToBatchQueue()
  → processBatchIfReady() → [waits for 10s batch window] → acquires script lock
  → processBatchAdd() or processBatchRemove() → one Katana SA per item → releases lock
  → logActivity() → Activity sheet

The batch window (BATCH_WINDOW_MS = 10000ms) is implemented via CacheService sliding
window. The script lock is acquired after the window closes to serialize SA creation.

## Fix Applied — 2026-03-18

### Fix A — debug/03_WaspCallouts.js (processBatchIfReady)
Script lock now released BEFORE any Katana API calls. Lock is only held for fast
cache operations (dedup check → mark processed). All slow work (SA creation,
logActivity, WASP API calls) happens outside the lock. Lock hold time reduced from
15–40 seconds to milliseconds.

### Fix B — debug/02_Handlers.js (acquireExecutionGuard_ / releaseExecutionGuard_)
Both functions now catch waitLock() timeout with try-catch. On timeout: logs locally,
returns false/returns gracefully. Exception no longer propagates to doPost → no Slack
alert even if contention occurs for other reasons in future.

Status: pushed to @HEAD 2026-03-18. Observing to confirm Slack alerts stop.
Note: fix is at @HEAD only — needs clasp deploy + callout URL update to go live at @405.

## Related Issues
- See: fixed-f5-concurrent-sheet-entries.md
  processWebhookQueue() (F5) deliberately uses CacheService instead of LockService
  to avoid adding to this contention. Comment in the code explains this explicitly.
