# F2 — WASP Adjustments Flow: Setup, Rules & History

## What F2 Does

F2 keeps Katana inventory in sync with manual adjustments made in WASP.

When a user adds or removes stock in WASP (via UI or API), WASP fires an HTTP callout
to the debug webhook. F2 receives it, validates it, waits for a 10-second batch window
to collect any related items, then creates one Katana Stock Adjustment (SA) per item.
The SA is logged to the Activity sheet under "F2 Adjustments".

---

## Architecture

```
WASP UI / WASP API
        │
        ▼
WASP Callout (HTTP POST) ─────────────────────────────┐
        │                                              │
        ▼                                              │
debug webhook doPost()                                 │
        │                                              │
        ▼                                              │
handleWaspQuantityAdded() / handleWaspQuantityRemoved()│
        │                                              │
        ├─ Suppression checks (see Rules below)        │
        │   ├─ F5 echo? (ShipStation notes) → skip     │
        │   ├─ Internal note? (Sheet push etc.) → skip │
        │   └─ enginPreMark_ cache hit? → skip         │
        │                                              │
        ├─ addToBatchQueue() ──────────────────────────┘
        │   └─ 10-second sliding window (BATCH_WINDOW_MS)
        │
        ▼
processBatchIfReady()
        │
        ├─ Acquires LockService.getScriptLock() (10s wait)
        ├─ processBatchAdd() / processBatchRemove()
        │   └─ processSiteBatchAdd() / processSiteBatchRemove()
        │       ├─ getKatanaVariantBySku()
        │       ├─ getWaspLotInfo() (lot lookup if needed)
        │       ├─ createKatanaBatchStockAdjustment()
        │       └─ logActivity('F2', ...) → Activity sheet
        │
        └─ Releases lock
```

---

## Two Callouts Configured in WASP

Both are under WASP UI → Settings → Callouts.

| Callout | Event | Action field | Status |
|---------|-------|--------------|--------|
| Katana Sync - Item Added | Inventory add | `quantity_added` | Enabled/Disabled |
| Katana Sync - Item Removed | Inventory remove | `quantity_removed` | Enabled/Disabled |

**Both post to the same debug webhook URL.**

**Payload template (same structure for both, action field differs):**
```json
{
  "action": "quantity_added",
  "ItemNumber": "{trans.AssetTag}",
  "Quantity": "{trans.AssetTransQuantity}",
  "SiteName": "{trans.SiteName}",
  "LocationCode": "{trans.LocationCode}",
  "Lot": "{trans.Lot}",
  "DateCode": "{trans.DateCode}",
  "Notes": "{trans.Notes}",
  "Assignee": "{trans.Assignee}"
}
```

**Static headers always sent:**
- `Wasp-Callout-Timestamp`
- `Wasp-Callout-Name`
- `Wasp-Callout-Signature` (HMAC using Private Key — not currently verified by debug script)
- `Content-Type: application/json`

---

## Critical: {trans.Notes} Does NOT Work for API Transactions

WASP only substitutes `{trans.Notes}` for transactions entered via the UI.
For API-triggered transactions (e.g. sync sheet push), `{trans.Notes}` remains
as the literal string `{trans.Notes}` in the callout payload.

**Consequence:** Note-based suppression (`isInternalWaspAdjustmentNote_`) can NEVER
suppress API-triggered callouts. It only works as a last-resort fallback.

---

## Suppression System — Preventing Double Adjustments

When the sync sheet (or any internal code) pushes to WASP via API, WASP fires a
callout back. Without suppression, F2 would create a Katana SA for the same change
that was already made intentionally — double-counting.

### Suppression Method 1 — enginPreMark_ (primary)

The sync script calls `enginPreMark_` BEFORE every WASP API write. This POSTs
`{ action: "engin_mark", sku, location, op }` to the debug webhook URL. The debug
script sets a CacheService key (120-second TTL) for that SKU+location+operation.

When the callout arrives, `wasRecentlySyncedToWasp()` finds the cache key → F2 skips.

**Required setup:**
- `GAS_WEBHOOK_URL` in the sync script's Script Properties must equal the EXACT same
  URL that the WASP callouts are configured to POST to.
- Current correct URL: `https://script.google.com/macros/s/AKfycbw4Z2YtIz7hIkS-48fi6f5DswvzafmgqyDuxmAgHuofzKYf3KiLNtiwbiBfwVuaBBsYTg/exec`
- This matches debug script deployment @405.

**The enginPreMark_ is baked into all WASP write functions in sync/03_WaspAPI.js:**
```
waspAddInventory()           → enginPreMark_(sku, location, 'add')
waspAddInventoryWithLot()    → enginPreMark_(sku, location, 'add')
waspRemoveInventory()        → enginPreMark_(sku, location, 'remove')
waspRemoveInventoryWithLot() → enginPreMark_(sku, location, 'remove')
waspAdjustInventory()        → enginPreMark_(sku, location, 'add')
```

### Suppression Method 2 — Note matching (fallback only)

`isInternalWaspAdjustmentNote_()` checks if the callout's Notes field starts with
known internal prefixes ('Sheet push ', 'Batch push ', 'ShipStation shipment ', etc.).
This only works for UI-triggered WASP transactions where the user manually entered
one of those note strings. For API transactions, Notes is always `{trans.Notes}`.
**Do not rely on this as a primary suppression mechanism.**

### Suppression Method 3 — F5 echo detection

`isF5ManagedWaspAdjustmentEcho_()` checks for F5-specific note prefixes
('ShipStation shipment ', 'Voided label', 'SO Cancelled: '). This prevents F5's
WASP deductions from also triggering F2 Katana SAs.

---

## F2 Skip Locations

`F2_SKIP_LOCATIONS = [LOCATIONS.SHOPIFY]`

The SHOPIFY location is managed exclusively by F5. F2 used to blanket-skip all
callouts from SHOPIFY. As of the current version, F2 only suppresses SHOPIFY callouts
when they clearly came from F5/internal sync (F5 echo detection + enginPreMark_ cache).
Manual WASP adjustments at SHOPIFY DO create Katana SAs.

---

## Cost Behaviour (Known Issue)

When F2 creates a Katana SA for an add, it sources `cost_per_unit` from:
1. Katana variant `purchase_price` / `default_purchase_price` / `cost`
2. Falls back to `getKatanaAverageCost_()` (Katana average cost API call)

**WASP item cost is NOT used.** WASP callout payloads contain no cost field.
A `getWaspItemCost_()` function exists in the debug script but is not yet wired into
the F2 batch processing path.

**Observed discrepancy:** WASP shows 0.07/EA for 2OZSEAL, Katana SA shows 0.06600 CAD
(Katana average cost). These differ because Katana tracks weighted average cost, not
WASP's catalogue cost. This is a known open issue — fix requires calling
`getWaspItemCost_()` before SA creation. (See issues/open-f2-lock-timeout-slack-alert.md)

For remove SAs: `cost_per_unit` is intentionally omitted — Katana rejects it on
negative adjustments and uses average cost automatically.

---

## Batch Window

`BATCH_WINDOW_MS = 10000` (10 seconds)

Multiple callouts arriving within 10 seconds of each other are grouped into one batch.
Each item still gets its own Katana SA (one SA per item, not one SA for all items).
The batch window prevents rapid sequential WASP movements from creating duplicate SAs.

---

## Issues History

### FIXED — Double logging from sheet adjustments (2026-03-18)

**Root cause (found 2026-03-18):**
- `GAS_WEBHOOK_URL` in the sync script pointed to a stale/wrong URL that was not
  the same as the active debug webhook URL.
- The enginPreMark_ HTTP call was going to a different script → different CacheService
  namespace → cache key not found when the callout arrived → suppression failed.
- `{trans.Notes}` is never substituted by WASP for API transactions, so note-based
  suppression also never worked for sheet adjustments.

**Fix applied 2026-03-18 — confirmed working:**
- `GAS_WEBHOOK_URL` in sync script Script Properties updated from stale URL
  (`AKfycbwVfK...ll_DQ`) to the correct deployment @405 URL:
  `https://script.google.com/macros/s/AKfycbw4Z2YtIz7hIkS-48fi6f5DswvzafmgqyDuxmAgHuofzKYf3KiLNtiwbiBfwVuaBBsYTg/exec`
- Verified: sheet adjustment made after fix → NOT logged in Activity sheet ✅
- Updated via temporary `fixGasWebhookUrl()` function in sync/test.js (removed after use).

**Ongoing rule — must maintain:**
- `GAS_WEBHOOK_URL` must always match the URL the WASP callouts are configured to POST to.
- When the debug script is redeployed (new version number), BOTH the WASP callout URL
  AND GAS_WEBHOOK_URL must be updated to the new deployment URL simultaneously.

### OPEN — F2 cost not using WASP item cost

`getWaspItemCost_()` exists but is not called in the F2 batch processing path.
Katana uses its own average cost instead of WASP's catalogue cost.

### OBSERVING — Script lock timeout causing Slack alerts

Fix applied 2026-03-18 — observing to confirm alerts stop.
- `processBatchIfReady` now releases the script lock before any Katana API calls.
- `acquireExecutionGuard_` / `releaseExecutionGuard_` now catch lock timeouts gracefully.
(See issues/open-f2-lock-timeout-slack-alert.md)

---

## Deployment Rule (Critical)

The debug script has multiple deployments. WASP callouts and enginPreMark_ must
ALWAYS target the SAME deployment URL.

Current active deployment: **@405**
URL: `https://script.google.com/macros/s/AKfycbw4Z2YtIz7hIkS-48fi6f5DswvzafmgqyDuxmAgHuofzKYf3KiLNtiwbiBfwVuaBBsYTg/exec`

`clasp push` only updates @HEAD code — it does NOT update numbered deployments.
To deploy code changes to the active URL, run `clasp deploy` to create a new version,
then update both:
1. WASP callout URL (both add and remove callouts) → new deployment URL
2. `GAS_WEBHOOK_URL` in sync script Script Properties → same new deployment URL

---

## Key Files

| File | Role |
|------|------|
| `debug/01_Router.js` | `doPost` entry point, routes `quantity_added/removed` to F2 |
| `debug/03_WaspCallouts.js` | F2 handlers, batch queue, suppression logic, SA creation |
| `debug/08_Logging.js` | `logActivity`, `runWithScriptWriteLock_` |
| `sync/03_WaspAPI.js` | `enginPreMark_`, all WASP write functions |
| `sync/05_PushEngine.js` | Sheet push logic that triggers WASP writes |
