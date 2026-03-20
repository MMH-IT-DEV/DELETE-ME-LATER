# F1 — PO Receiving Flow: Setup, Rules & History

## What F1 Does

F1 processes Purchase Order receives from Katana. When a PO is marked as received
(fully or partially), Katana fires a webhook. F1 adds the received items to WASP
at the correct site/location. When a receive is reverted, F1 removes the items.

## Webhook Events

| Katana Event | Handler | Purpose |
|---|---|---|
| `purchase_order.created` | `handlePurchaseOrderCreated` | Logs creation, captures open-row snapshot |
| `purchase_order.received` | `handlePurchaseOrderReceived` | Main receive — adds items to WASP |
| `purchase_order.partially_received` | `handlePurchaseOrderReceived` | Same handler for partial receives |
| `purchase_order_row.received` | Delegates to above | Per-row receive |
| `purchase_order.updated` | `handlePurchaseOrderUpdated` | Revert detection + fallback receive |

## Location Mapping (KATANA_LOCATION_TO_WASP)

All sites go to **PRODUCTION** (updated 2026-03-19, deployed @410):

| PO "Ship to" | WASP Site | WASP Location |
|---|---|---|
| MMH Kelowna | MMH Kelowna | **PRODUCTION** |
| MMH Mayfair | MMH Mayfair | **PRODUCTION** |
| Storage Warehouse | Storage Warehouse | **SW-STORAGE** |

## Deduplication (4 layers)

1. **Execution guard**: `po_state_guard_{poId}` (120s TTL, PropertiesService)
   Prevents concurrent processing of same PO
2. **Cache dedup**: `po_received_{poNumber}` (600s TTL for full, 8s for partial)
   Blocks repeat webhooks within 10 minutes
3. **Activity log scan**: `isPOAlreadyReceived()` checks Activity sheet
   Blocks if non-reverted, non-Failed F1 entry exists for this PO
4. **Partial signature**: MD5 hash of open-row quantities prevents duplicate partials

## Partial Receive Handling

F1 supports multiple partial receives on the same PO:
- Each partial receive is stored as a separate **receipt block**
- Blocks are labeled PO-XXX, PO-XXX/2, PO-XXX/3, etc.
- Each block can be reverted independently
- "Revert all" reverts ALL blocks (both partial receives)

State is stored in PropertiesService:
- `f1_recv_blocks_{poRef}` — receipt blocks (grouped by receive event)
- `f1_recv_{poRef}` — flat receive state (legacy)
- `f1_open_rows_{poRef}` — open-row snapshot
- `f1_recv_partial_{poRef}` — partial receive flag
- `f1_count_{poRef}` — receive counter for /2, /3 labels

## Revert Handling

`purchase_order.updated` detects reverts:
- **Full revert (Revert All)**: PO status = NOT_RECEIVED → removes all stored items
  from ALL receipt blocks in WASP. Each block logs a separate revert entry.
- **Partial revert (single Revert button)**: Compares stored batch IDs against
  current Katana rows. Only reverts the specific receipt block that was un-done.
- **Non-batch partial revert**: Compares received_quantity deltas (requires HOTFIX_FLAGS)
- After revert: clears dedup cache so fresh receives can go through

## Race Condition Prevention

When Katana fires both `purchase_order.partially_received` AND `purchase_order.updated`
simultaneously, the updated handler now **skips entirely** if a recent receive was just
processed (`hasRecentPOReceive_` cache flag, 20s TTL). This prevents the updated handler
from running a stale open-row diff and adding/reverting phantom items.

Reverts (NOT_RECEIVED) are NOT affected by this skip — they always process normally.

## Purchase UOM Conversion

```javascript
var puomRate = parseFloat(row.purchase_uom_conversion_rate) || 1;
quantity = rawQty * puomRate;
```
Items ordered in bulk units (pallets, dozens) are converted to stocking units.

## Cost Update

After receiving, F1 updates each item's WASP cost from Katana's purchase_price
via `waspUpdateItemCost()`.

## Item Creation

If a PO item doesn't exist in WASP, F1 will fail with "item does not exist."
The item must be created first via the sync sheet's Katana tab:
1. Change WASP Status from "NEW" → "Push"
2. Run "Push to WASP" from the menu
3. Re-receive the PO

The push handles category mismatches automatically (retries without category if
WASP rejects the Katana category name).

## Hotfix Flags (00_Config.js)

| Flag | Default | Purpose |
|---|---|---|
| `F1_PARTIAL_NON_BATCH_DELTA` | true | Use delta qty for non-batch partial receives |
| `F1_CONFIRM_NON_BATCH_REVERT` | true | Re-fetch to confirm non-batch revert |
| `F1_CONFIRM_FULL_REVERT` | true | Adaptive re-fetch to confirm NOT_RECEIVED |

## Confirmed Working (2026-03-19)

- Full receive ✅ (PO-602, PO-575)
- Full revert ✅ (PO-602, PO-575)
- Retry after Failed ✅ (PO-575 — isPOAlreadyReceived allows Failed entries)
- MMH Mayfair → PRODUCTION ✅ (PO-575)
- MMH Kelowna → PRODUCTION ✅ (deployed @410)
- Partial receive (2 batches) ✅ (PO-605: B-QUAD+B-PINK-1, then B-PURPLE-1+B-YELLOW-2)
- Partial revert (single batch) ✅ (PO-605/2 reverted independently)
- Revert All (all batches) ✅ (PO-605: both blocks reverted correctly, exactly 4 items)
- Item creation via Push ✅ (NI-LADDER10 created, then received at MMH Mayfair)

## Issues History

### FIXED — Failed PO permanently blocks re-receive (2026-03-19 @408)

`isPOAlreadyReceived()` blocked re-processing when the most recent F1 entry
was "Failed". Fixed by reading the Status column (E) and allowing retry when
status = "Failed".

### FIXED — Race condition: phantom items added then reverted (2026-03-19 @409)

When a partial receive happens, Katana fires BOTH `purchase_order.partially_received`
AND `purchase_order.updated` simultaneously. The receive handler processes correctly,
but the updated handler also runs an open-row diff that picks up stale/wrong rows
from the Katana API — adding extra items, then auto-correcting by reverting them.

Fix: when `hasRecentPOReceive_()` is true and status is RECEIVED/PARTIALLY_RECEIVED,
`handlePurchaseOrderUpdated` now skips entirely (returns immediately). The receive
handler already did the work. Reverts (NOT_RECEIVED) still process normally.

### FIXED — MMH Kelowna location RECEIVING-DOCK → PRODUCTION (2026-03-19 @410)

Changed `KATANA_LOCATION_TO_WASP['MMH Kelowna'].location` from RECEIVING-DOCK
to PRODUCTION. All F1 receives now land at PRODUCTION regardless of site.

### FIXED — Category mismatch blocks item creation (2026-03-19)

When pushing a new item from the Katana tab, WASP rejected the Katana category
(e.g., "SUPPLIES" doesn't exist in WASP). Fixed by retrying the create without
`CategoryDescription` if WASP returns error -57006.

### NOT SUPPORTED — Historical PO reverts

POs received before the F1 system was tracking them have no stored state. Reverting
these POs produces no Activity entry because there's nothing to compare against.
Only POs received AFTER the system was deployed can be reverted.

## Key Files

| File | Role |
|------|------|
| `debug/script/02_Handlers.js` | All F1 handlers (receive lines 19-664, revert lines 677-1276) |
| `debug/script/00_Config.js` | KATANA_LOCATION_TO_WASP, HOTFIX_FLAGS |
| `debug/script/08_Logging.js` | logActivity, flow labels/colors |
| `debug/script/05_WaspAPI.js` | waspAddInventoryWithLot, waspRemoveInventoryWithLot |
| `sync/script/05_PushEngine.js` | pushKatanaSyncRows_ (item creation from Katana tab) |
| `sync/script/04_SyncEngine.js` | SYNC_LOCATION_MAP, waspCreateItem |
| `sync/script/04b_RawTabs.js` | getKatanaSyncTarget_ (location routing for push) |

## Deployment History

| Version | Date | Change |
|---------|------|--------|
| @408 | 2026-03-19 | Allow retry after Failed PO receive |
| @409 | 2026-03-19 | Skip open-row diff when receive handler just processed |
| @410 | 2026-03-19 | MMH Kelowna RECEIVING-DOCK → PRODUCTION |
