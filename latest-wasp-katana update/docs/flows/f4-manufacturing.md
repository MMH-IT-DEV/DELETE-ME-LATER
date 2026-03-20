# F4 — Manufacturing Order Flow: Setup, Rules & History

## What F4 Does

F4 processes Manufacturing Order completions from Katana. When an MO is marked as
Done, F4 removes ingredients from WASP at PRODUCTION and adds the finished/intermediate
output product. When an MO is reverted (set to Not Started or Work in Progress),
F4 reverses the WASP changes.

## Webhook Events

| Katana Event | Handler | Purpose |
|---|---|---|
| `manufacturing_order.done` | `handleManufacturingOrderDone` | Main completion — removes ingredients, adds output |
| `manufacturing_order.updated` (status=DONE) | Delegates to above | Fallback for missed .done webhooks |
| `manufacturing_order.updated` (non-DONE) | Revert check | If snapshot exists → reverse WASP changes |
| `manufacturing_order.deleted` | `handleManufacturingOrderDeleted` | Checks for snapshot to reverse |

## Location Mapping

All F4 operations use **PRODUCTION** at **MMH Kelowna**:

| Operation | WASP Site | WASP Location |
|---|---|---|
| Remove ingredients | MMH Kelowna | PRODUCTION |
| Add output | MMH Kelowna | PROD-RECEIVING |

Config: `FLOWS.MO_INGREDIENT_LOCATION = 'PRODUCTION'`
         `FLOWS.MO_OUTPUT_LOCATION = 'PROD-RECEIVING'`

## Ingredient Lot Resolution (3-tier fallback)

For batch-tracked ingredients, F4 resolves the lot number using:

**Tier 1 — Katana batch_transactions:**
Extract lot from `batch_transactions` in the recipe row (via `extractKatanaBatchNumber_`).
Falls back to `batch_stock` embedded object, then `fetchKatanaBatchStock(batchId)`,
then variant-level `batch_stocks?variant_id=X&include_deleted=true`.

**Tier 2 — WASP single-lot fallback:**
If Katana can't resolve the lot, query WASP for all lots at PRODUCTION.
If exactly 1 lot exists → use it.

**Tier 3 — WASP qty-match fallback (deployed @412):**
If multiple WASP lots exist, filter to lots with sufficient quantity for the
deduction. If exactly 1 has enough → use it. If multiple have enough → skip.

Example: FGJ-IS-1 needs 1785 PCS. WASP has TTFC091A (1760), TTFC092A (1760),
TTFC093 (1785). Only TTFC093 has enough → picked automatically.

## Deduplication

1. **Execution guard**: `mo_done_guard_{moId}` (120s TTL, PropertiesService)
2. **Cache dedup**: `mo_done_{moId}` (300s / 5-min TTL, CacheService)
3. **Snapshot check**: `mo_snapshot_{moRef}` (PropertiesService, permanent until revert)
4. **Activity log**: `isMOAlreadyCompleted()` — only blocks on status "Complete" (not Partial/Failed)

## Snapshot System

After processing, F4 saves a snapshot to `mo_snapshot_{moRef}` containing:
- All ingredients (SKU, qty, lot, expiry, location)
- Output product (SKU, qty, lot, expiry)
- Stage (INTERMEDIATE, FINISHED)

The snapshot is used by:
- Dedup: prevents re-processing
- Revert: knows exactly what to reverse in WASP
- id_ref mapping: `mo_id_ref_{moId}` maps numeric ID → human-readable ref

## Revert Handling

`handleManufacturingOrderUpdated` handles ANY non-DONE status as a potential revert:
- Checks for F4 snapshot (uses `waitForMOSnapshotForRevert_` with adaptive wait)
- If snapshot found → acquires revert guard → reverses WASP changes
- Ingredients: added back to PRODUCTION (restored)
- Output: removed from PROD-RECEIVING (removed)
- Snapshot deleted after successful revert
- ALL cache keys cleared to allow immediate re-completion

**Revert triggers on:** NOT_STARTED, IN_PROGRESS, BLOCKED, or any non-DONE status.
All are treated the same — the snapshot presence is what matters.

**Revert window:** 14 days (`REVERT_WINDOW_DAYS`). Older reverts are blocked.

## Critical: How the Revert Cycle Works (Do Not Break)

The complete→revert→re-complete→re-revert cycle depends on these pieces
working together. If any future change breaks the cycle, check this list.

### On completion (`handleManufacturingOrderDone`):
1. Execution guard acquired: `mo_done_guard_{moId}` (120s TTL)
2. Dedup key set: `mo_done_{moId}` (300s TTL)
3. Ingredients removed from WASP, output added
4. Snapshot saved: `mo_snapshot_{moRef}` (ScriptProperties, permanent)
5. ID mapping saved: `mo_id_ref_{moId}` (ScriptProperties, permanent)

### On revert (`reverseMOSnapshot`):
1. Ingredients added back to WASP, output removed
2. Snapshot deleted: `mo_snapshot_{moRef}`
3. ID mapping deleted: `mo_id_ref_{moId}`
4. Consumed map cleared: `mo_consumed_{moRef}`
5. **ALL cache keys cleared** (this is what enables the re-completion cycle):
   - `mo_done_{moId}` — dedup key
   - `mo_staging_{moId}` — staging key
   - `mo_done_guard_{moId}` — execution guard
   - `mo_revert_guard_{moId}` — revert guard

If ANY of these cache keys are not cleared, re-completion will be blocked.
The symptom is: "MO completion already in progress" or "MO already processed"
in the webhook queue, and no new Activity entry appears.

### On status confirmation (`F4_CONFIRM_STATUS_REVERT` hotfix):
When a non-DONE webhook arrives for a completed MO, the code re-fetches the
MO from Katana API with delays `[0, 2000, 4000, 8000]` to confirm the status.
If Katana API still returns DONE after all retries (14 seconds), the code
**proceeds with the revert anyway** — the webhook is trusted over the slow API.
This is logged as `MO_REVERT_API_SLOW_PROCEEDING`.

**Do not change this back to blocking.** Katana's API propagation is slower
than its webhooks. Blocking caused reverts to silently fail (the original bug).

## Smart Retry (Consumed Map)

`mo_consumed_{moRef}` tracks which ingredients were already deducted from WASP.
On retry (re-processing a failed/partial MO), already-consumed items are skipped
to prevent double-deduction.

## Hotfix Flags (00_Config.js)

| Flag | Default | Purpose |
|---|---|---|
| `F4_CONFIRM_STATUS_REVERT` | true | Re-fetch MO status before acting on revert |

## Confirmed Working (2026-03-19)

- MO completion ✅ (MO-7460, MO-7462, MO-7463, MO-7464)
- Lot fallback (single lot) ✅ (MO-7460 — FGJ-IS-1 picked TTFC091A)
- Lot fallback (qty match) ✅ (multiple lots, only 1 with enough qty)
- Revert (first cycle) ✅ (MO-7448, MO-7462, MO-7463 — all ingredients restored)
- Revert (full cycle: complete → revert → re-complete → re-revert) ✅ (MO-7464)
- Revert triggers on IN_PROGRESS and NOT_STARTED ✅
- WIP status change without snapshot → no action (correct) ✅ (MO-7464)
- Non-batch ingredients ✅ (PL-GREEN-1, B-PURPLE-4, L-M, NI-B221210, FBA-FRAGILE)

## Known Limitations

### ~~Rapid complete → revert → complete → revert timing issue~~ FIXED @423
Fixed 2026-03-19. `reverseMOSnapshot` now clears execution guards
(`mo_done_guard_`, `mo_revert_guard_`) alongside dedup keys. Also trusts
webhook status over slow Katana API confirmation. Full cycle works immediately
with no waiting period.

### Partial MO re-processing
If an MO completes with some ingredients Skipped (e.g., lot not found), the
snapshot is saved. Re-completing the same MO requires clearing the snapshot first
via `fullResetMO{ID}` in `99_TestSetup.js`.

### FGJ-IS-1 lot resolution
For labelling MOs where the input (FGJ-IS-1) uses the same batch as the output
(LTG-1), the lot may not be resolved from Katana's API. The WASP qty-match
fallback handles this when the correct lot is the only one with sufficient quantity.

## Issues History

### FIXED — WASP lot fallback for multi-lot items (2026-03-19 @412)
Added qty-based matching when multiple WASP lots exist. Filters to lots with
sufficient quantity for the deduction. If exactly 1 matches → use it.

### FIXED — Snapshot not saving (from previous session @402)
Moved `saveMOSnapshot` before `logFlowDetail` and Slack notification.
Wrapped Slack in try-catch so it can never crash the execution before snapshot save.

### FIXED — Revert not firing (from previous session @402)
Fixed `mo_id_ref_{moId}` mapping and added fallback scan in `getMOSnapshotRaw_`.

## Key Files

| File | Role |
|------|------|
| `debug/script/02_Handlers.js` | handleManufacturingOrderDone (~line 3625), revert (~line 4811), reverseMOSnapshot (~line 5009) |
| `debug/script/00_Config.js` | FLOWS.MO_INGREDIENT_LOCATION, FLOWS.MO_OUTPUT_LOCATION, HOTFIX_FLAGS |
| `debug/script/05_WaspAPI.js` | waspRemoveInventoryWithLot, waspAddInventoryWithLot, waspLookupAllLots_ |
| `debug/script/08_Logging.js` | logActivity, runWithScriptWriteLock_ |
| `debug/script/99_TestSetup.js` | fullResetMO, checkMO, listAllMOSnapshots diagnostics |

## Deployment History

| Version | Date | Changes |
|---------|------|---------|
| @402 | 2026-03-18 | Snapshot save fix, revert fix, dedup fix |
| @412 | 2026-03-19 | WASP lot fallback with qty matching |
| @416 | 2026-03-19 | Script lock revert (fixed missing Activity entries) |
| @423 | 2026-03-19 | Fix revert cycle: clear guards on revert, trust webhook over slow API, revert window 5d→14d |
