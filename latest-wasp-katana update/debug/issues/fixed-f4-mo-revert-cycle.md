# FIXED — F4: MO Revert Cycle (Complete → Revert → Re-complete → Re-revert)

## Status
FIXED — deployed @423 on 2026-03-19. Tested and confirmed working.

## What Was Happening
Two bugs prevented the MO revert cycle from working reliably:

### Bug 1 — Revert blocked by slow Katana API confirmation
When a user set an MO from DONE back to NOT_STARTED, the webhook arrived with
status NOT_STARTED. The F4_CONFIRM_STATUS_REVERT hotfix re-fetched the MO from
Katana API to confirm the status change. But Katana's API propagation was slower
than the webhook — the re-fetch still returned DONE after 2 seconds of retries.
The code treated this as a "transient edit webhook" and blocked the revert.

### Bug 2 — Re-completion blocked after revert
After a successful revert, `reverseMOSnapshot` cleared the dedup cache key
(`mo_done_{moId}`) and the snapshot, but did NOT clear the execution guard
(`mo_done_guard_{moId}`, 120s TTL). Re-completions within 120 seconds were
blocked with "MO completion already in progress". Without a new completion,
no new snapshot was saved. Subsequent reverts then failed with "no action
unless reverting a completed MO" (no snapshot found).

## Fixes Applied

### Fix 1 — Trust webhook over slow API (02_Handlers.js)
- Increased confirmation delays from [0, 1000, 2000] to [0, 2000, 4000, 8000]
  (14 seconds total instead of 2)
- If API still returns DONE after all retries, proceed with revert instead of
  blocking. Log MO_REVERT_API_SLOW_PROCEEDING for observability.
- Reasoning: webhook explicitly says NOT_STARTED + snapshot exists = genuine
  status change. Transient edit webhooks carry DONE status, never NOT_STARTED.

### Fix 2 — Clear execution guards on revert (02_Handlers.js)
Added to `reverseMOSnapshot` cleanup (lines 5107-5110):
```
moCache.remove('mo_done_guard_' + cacheMoId);
moCache.remove('mo_revert_guard_' + cacheMoId);
```
MO can now be immediately re-completed and re-reverted with no waiting period.

### Additional — Revert window extended
`REVERT_WINDOW_DAYS` changed from 5 to 14 days.

## Test Results (2026-03-19 21:47-21:53)
MO-7464 full cycle verified:
1. Complete → Activity entry (consumed + produced) ✓
2. Revert to WIP → Activity entry (restored + removed) ✓
3. Re-complete → Activity entry (consumed + produced) ✓
4. Revert to NOT_STARTED → Activity entry (restored + removed) ✓
5. Set to WIP (no snapshot) → no action (correct) ✓

## Key Files
- debug/script/02_Handlers.js — handleManufacturingOrderUpdated, reverseMOSnapshot
