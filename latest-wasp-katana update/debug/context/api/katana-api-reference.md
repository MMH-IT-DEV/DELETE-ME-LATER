# Katana MRP API — Working Reference

**Last updated**: 2026-02-26
**Source**: Katana developer docs + observed production behavior in this project (Feb 2026).

---

## Base URL & Authentication

```
Base URL:  https://api.katanamrp.com/v1
Auth:      Bearer token in Authorization header
```

All requests require:
```
Authorization: Bearer {KATANA_API_KEY}
Accept: application/json          (GET requests)
Content-Type: application/json    (POST requests)
```

The API key is stored in GAS ScriptProperties as `KATANA_API_KEY` and accessed via `CONFIG.KATANA_API_KEY`.

---

## Response Envelope

**Single-resource endpoints** (GET by ID) wrap the object in `{ data: { ... } }`:
```javascript
// GET /sales_orders/12345  → { data: { id, order_no, status, sales_order_rows, ... } }
var so = soData.data ? soData.data : soData;  // always unwrap like this
```

**Collection endpoints** (GET with filters/pagination) return `{ data: [ ... ] }`:
```javascript
var rows = result.data || result || [];
```

**Some endpoints return the object directly** (no wrapper). Variants returned by ID do NOT have a `data` wrapper — `variant.sku` works directly. Always write safe unwrap:
```javascript
var v = result.data ? result.data : result;
```

---

## Pagination

| Param | Notes |
|-------|-------|
| `limit` | Items per page (default 50, max 250) |
| `page` | Page number (1-indexed) |
| `sort` | e.g. `-updated_at` (descending) |
| `per_page` | Also accepted on some endpoints (see audit code) |

Response headers include `X-Pagination` with `total_records` and `total_pages`.

---

## Endpoints Used in This Project

### Products

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/products/{id}` | Returns product with `type` field ('product', 'material', 'intermediate') |

Key response fields: `id`, `name`, `type`, `category`.

Used in F1/F4 to determine if a variant is a material (→ PRODUCTION) or product (→ default location).

---

### Variants

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/variants/{id}` | Returns variant (SKU, product type). No `data` wrapper. |
| GET | `/variants?sku={sku}` | Search by SKU — returns `{ data: [ ... ] }` array |

Key response fields: `id`, `sku`, `product_id`, `material_id`, `batch_tracking` (boolean).

The `product.type` may be embedded on the variant if the product data is included in the response. If not, call `GET /products/{product_id}` as fallback.

**GAS note**: Cached in `CacheService` in `07_F3_AutoPoll.gs` (1-hour TTL) to avoid repeated variant lookups during bulk operations.

---

### Locations

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/locations` | All warehouse locations |
| GET | `/locations?name={name}` | Filter by name |

Key response fields: `id`, `name`.

Known Katana locations in this project: `MMH Kelowna`, `Storage Warehouse`, `Amazon USA` (FBA, skip), `Shopify` (virtual channel, skip).

**GAS note**: `fetchKatanaLocation()` is defined in `07_F3_AutoPoll.gs` with CacheService (60s TTL) since it's called repeatedly inside loop.

---

### Inventory

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/inventory` | Per-location stock for all variants |
| GET | `/inventory?limit=250&page={n}` | Paginated |

Key response fields: `variant_id`, `location_id`, `quantity_in_stock`, `quantity_committed`, `quantity_expected`.

Used by the sync engine (`10_InventorySync.gs`) to compare Katana stock against WASP.

---

### Batch Stocks

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/batch_stocks/{id}` | Single batch record by ID |
| GET | `/batch_stocks/{id}?include_deleted=true` | Include depleted batches (often still 404) |
| GET | `/batch_stocks?variant_id={id}` | All batches for a variant |
| GET | `/batch_stocks?batch_id={id}&include_deleted=true` | Query by batch_id with deleted |

Key response fields: `id` (or `batch_id`), `batch_number` (also try `batch_nr`, `nr`), `expiration_date` (also try `expiry_date`, `best_before_date`), `in_stock`.

**CRITICAL QUIRK — Katana permanently deletes batch records when stock reaches zero.** `GET /batch_stocks/{id}` returns 404 for a depleted batch. `include_deleted=true` does not recover them. All three fallback tiers in `fetchKatanaBatchStock()` will return null for a fully-consumed batch.

The field for expiry is `expiration_date` (NOT `expiry_date` or `best_before_date` — those are the wrong names for this endpoint). Always check all three:
```javascript
var expiry = batch.expiration_date || batch.expiry_date || batch.best_before_date || '';
```

---

### Purchase Orders

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/purchase_orders/{id}` | Single PO with header data |
| GET | `/purchase_orders?per_page=50&sort=-updated_at` | Paginated list |

Key response fields: `id`, `order_no`, `status`, `location_id` (Ship To location), `updated_at`.

PO statuses relevant to this project: `PARTIALLY_RECEIVED`, `RECEIVED`, `DONE`, `COMPLETED`.

`location_id` is the Katana warehouse location the PO is shipped to. Resolve via `GET /locations/{id}` to get the name, then map through `KATANA_LOCATION_TO_WASP` in `00_Config.gs`.

---

### Purchase Order Rows

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/purchase_order_rows?purchase_order_id={id}&include=batch_transactions` | All rows for a PO, with batch data |

Key response fields: `id`, `variant_id`, `quantity`, `received_quantity`, `status`, `batch_transactions`.

`batch_transactions` structure (when `include=batch_transactions` is appended):
```javascript
[{
  batch_id: 3099449,
  batch_number: "UFC410B",    // may be empty — resolve via batch_stocks
  quantity: 100,
  expiry_date: "2029-02-28",  // may be empty
  best_before_date: "...",    // alternate field name
  batch_stock: { batch_number: "...", nr: "..." }  // sometimes embedded
}]
```

If `batch_transactions.length > 1`, the row is split across multiple batches — loop each entry and add separately to WASP.

---

### Sales Orders

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/sales_orders/{id}` | Single SO — response has `data` wrapper |
| GET | `/sales_order_rows?sales_order_id={id}` | Line items for an SO |
| GET | `/sales_orders?per_page=50&sort=-updated_at` | Paginated list |

Key response fields on SO: `id`, `order_no`, `status`, `sales_order_rows` (may be embedded), `addresses` (for customer name), `updated_at`.

Key response fields on SO rows: `id`, `variant_id`, `quantity`, `delivered_quantity`.

Use `delivered_quantity` over `quantity` in the SO Delivered handler — it reflects what was actually shipped:
```javascript
var qty = (item.delivered_quantity != null) ? item.delivered_quantity : (item.quantity || 0);
```

SO statuses seen in practice: `NOT_SHIPPED`, `PARTIALLY_DELIVERED`, `FULFILLED`, `DELIVERED`, `CANCELLED`, `VOIDED`.

---

### Manufacturing Orders

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/manufacturing_orders/{id}` | Single MO with header data |
| GET | `/manufacturing_orders?limit=250&page={n}` | Paginated list |
| GET | `/manufacturing_order_recipe_rows?manufacturing_order_id={id}&include=batch_transactions` | Ingredients with batch allocation |

Key MO header fields: `id`, `order_no`, `status`, `variant_id` (the output product), `quantity`, `actual_quantity`, `batch_number`, `expiry_date`, `completed_at`, `done_at`, `updated_at`.

Prefer `actual_quantity` over `quantity` for the finished goods output quantity.

For MO expiry, check `mo.expiry_date` first. If empty and a lot exists, default to +3 years from `mo.completed_at || mo.done_at || mo.updated_at`.

MO statuses seen in practice: `IN_PROGRESS`, `DONE`, `COMPLETED`, `RESOURCE_COMPLETED`.

**Ingredient (recipe row) fields**: `id`, `variant_id`, `quantity`, `total_consumed_quantity`, `consumed_quantity`, `actual_quantity`, `batch_transactions`.

Prefer `total_consumed_quantity` over `consumed_quantity` or `quantity` for the consumed amount:
```javascript
var ingQty = ing.total_consumed_quantity || ing.consumed_quantity || ing.actual_quantity || ing.quantity || 0;
```

Ingredient `batch_transactions` are only populated AFTER the user assigns the batch in the "Done" popup in Katana UI. The webhook fires after this popup is submitted, so there is a brief race condition — the F4 handler uses `Utilities.sleep(15000)` to wait 15 seconds before fetching ingredients.

Ingredient `batch_transactions` contain `batch_id` and `quantity` but often NO `batch_number`. Resolve via `GET /batch_stocks/{batch_id}`. If the batch was fully consumed, it will be deleted and return 404 — use WASP lot lookup as fallback.

---

### Stock Adjustments

| Method | URL | Notes |
|--------|-----|-------|
| POST | `/stock_adjustments` | Create a stock adjustment |

Request body:
```javascript
{
  stock_adjustment_number: "WASP - Imported",
  stock_adjustment_date: new Date().toISOString(),
  location_id: 12345,          // Katana location ID
  reason: "WASP Batch Import",
  additional_info: "notes",
  stock_adjustment_rows: [
    {
      variant_id: 67890,
      quantity: 5,              // positive = add, negative = remove
      batch_transactions: [     // only for batch-tracked items
        { batch_id: 111, quantity: 5 }
      ]
    }
  ]
}
```

For batch-tracked items, you must include `batch_transactions` with the `batch_id`. Resolve the batch_id by calling `GET /batch_stocks?variant_id={id}` and matching on `batch_number` (case-insensitive).

---

### Stock Transfers

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/stock_transfers` | List all STs |
| GET | `/stock_transfers?per_page=50&sort=-updated_at` | Paginated, sorted by recent |

Key response fields: `id`, `stock_transfer_number`, `status`, `source_location_id`, `target_location_id`, `stock_transfer_rows`, `created_at`, `updated_at`.

Stock transfer row fields: `variant_id`, `quantity`, `batch_transactions`.

ST statuses seen in practice: `IN_PROGRESS`, `COMPLETED`, `DONE`, `RECEIVED`, `PARTIAL`, `IN_TRANSIT`.

**There is no stock transfer webhook.** F3 uses polling (`07_F3_AutoPoll.gs`) every 5 minutes via a time-based trigger. Syncs on statuses: `completed`, `done`, `received`, `partial`, `in_transit`.

---

### Webhooks (Management)

| Method | URL | Notes |
|--------|-----|-------|
| GET | `/webhooks` | List all registered webhooks |
| PATCH | `/webhook/{id}` | Update a webhook |
| DELETE | `/webhook/{id}` | Delete a webhook |

Webhooks are registered in the Katana UI or via API. Our webhook URL is the GAS web app deployment URL.

---

## Webhook Events

### Payload Structure

All Katana webhook payloads have the same top-level shape:
```javascript
{
  action: "purchase_order.received",   // event name (string)
  object: {
    id: 12345,          // Katana internal ID of the resource
    status: "RECEIVED"  // sometimes included on the object
    // other fields are NOT reliably populated here — always fetch via API
  }
}
```

The `object` in the payload is a lightweight reference, not the full resource. Always fetch the full record via API after receiving a webhook (e.g., `GET /purchase_orders/{id}`).

**Retry header**: `X-Katana-Retry-Num` — present on retry attempts. Values: `1`, `2`, or `3`. Not present on the initial delivery attempt.

### Complete Event List

#### Purchase Orders

| Event | Fires When | Action in This Project |
|-------|-----------|------------------------|
| `purchase_order.created` | PO is created | Log only (F1 PO Created handler) |
| `purchase_order.updated` | Any PO field changes | Not handled |
| `purchase_order.deleted` | PO deleted | Not handled |
| `purchase_order.partially_received` | PO status → `PARTIALLY_RECEIVED` | **Not yet handled** (Issue 3 — planned) |
| `purchase_order.received` | PO status → `RECEIVED` (fully received) | F1 handler: adds items to WASP |

#### Purchase Order Rows (line-item level)

| Event | Fires When | Action in This Project |
|-------|-----------|------------------------|
| `purchase_order_row.created` | PO row added | Not handled |
| `purchase_order_row.updated` | PO row modified | Not handled |
| `purchase_order_row.deleted` | PO row removed | Not handled |
| `purchase_order_row.received` | Individual row received | Not handled |

#### Sales Orders

| Event | Fires When | Action in This Project |
|-------|-----------|------------------------|
| `sales_order.created` | New SO created | F5: creates WASP pick order |
| `sales_order.updated` | Any SO field changes (~80% of all traffic) | Only queued if status = CANCELLED/VOIDED, then routes to cancel handler |
| `sales_order.deleted` | SO deleted | **Removed** — Katana API returns 404 for deleted SOs; no action possible |
| `sales_order.packed` | Status → PACKED | Not handled |
| `sales_order.delivered` | Status → DELIVERED | F5/F3: removes items from WASP (if not already picked) |
| `sales_order.cancelled` | SO cancelled | **Unreliable** — may not fire. Caught via `sales_order.updated` status check instead |
| `sales_order.availability_updated` | Stock availability or expected date changes | Not handled |

#### Manufacturing Orders

| Event | Fires When | Action in This Project |
|-------|-----------|------------------------|
| `manufacturing_order.done` | MO marked Done | F4: removes ingredients from PRODUCTION, adds output to PROD-RECEIVING |
| `manufacturing_order.updated` | Any MO field changes | Not handled |

No `manufacturing_order.completed` event exists. The correct event name is `manufacturing_order.done`.

#### Stock Transfers

No webhook events exist for stock transfers. Use polling.

---

## Webhook Retry Behavior

Katana expects an HTTP 2xx response within **10 seconds**. If not received, it retries:

| Attempt | Delay after previous |
|---------|---------------------|
| Initial | Immediate |
| Retry 1 | 30 seconds (`X-Katana-Retry-Num: 1`) |
| Retry 2 | 2 minutes (`X-Katana-Retry-Num: 2`) |
| Retry 3 | 15 minutes (`X-Katana-Retry-Num: 3`) |

After 4 total attempts (initial + 3 retries), Katana stops. Maximum delay from initial to final retry: ~17.5 minutes.

---

## Known Webhook Quirks

### Duplicate Fires (Critical)

**Katana fires each event multiple times.** This is not a retry issue — it is a consistent behavior pattern observed across all event types:

| Event | Observed duplicate pattern |
|-------|---------------------------|
| `manufacturing_order.done` | 2-3 fires within ~2 min, plus 1 delayed refire at ~17 min |
| `sales_order.delivered` | 3-4 fires for the same event |
| `purchase_order.received` | 4 fires over 18 minutes (observed: PO-547, WK-722 through WK-725) |
| `purchase_order.created` | 3 fires (initial + retries at 30s, 2 min) |

The ~17 minute delayed refire is distinct from the retry mechanism — it appears to be a separate Katana background job re-confirming the event.

**Dedup strategy in this project** (two layers):

1. **CacheService** (short-window): Blocks the rapid 2-3 duplicates within 2 minutes.
   - MO done: `mo_done_{moId}` — 300s TTL
   - SO delivered: `so_delivered_{soId}` — 300s TTL
   - SO created: `so_created_{soId}` — 300s TTL
   - PO received: `po_received_{poNumber}` — 600s TTL

2. **Activity log idempotency** (permanent): Blocks the ~17 min delayed refire after cache expires.
   - `isMOAlreadyCompleted(moRef)` — scans Activity sheet for any existing F4 entry with that MO ref
   - `isPOAlreadyReceived(poNumber)` — scans Activity sheet for any existing F1 entry with that PO ref

CacheService alone is insufficient for PO/MO processing — the 17-minute refire consistently arrives after the 300-600s cache window expires.

### sales_order.updated is ~80% of Traffic

Almost every SO status change (e.g., NOT_SHIPPED → PARTIALLY_DELIVERED) fires `sales_order.updated`. Most are useless for this integration.

Filter before queuing (done in `doPost()`):
```javascript
// Only queue sales_order.updated if status is CANCELLED or VOIDED
if (action === 'sales_order.updated') {
  var status = (payload.object && payload.object.status || '').toUpperCase();
  if (status !== 'CANCELLED' && status !== 'VOIDED') return;  // skip
}
```

### sales_order.cancelled Does Not Reliably Fire

`sales_order.cancelled` is a registered webhook event in Katana, but in practice it does not fire. The fallback is to listen to `sales_order.updated` and check `payload.object.status` for `'CANCELLED'` or `'VOIDED'` (always uppercase from Katana).

### Partial vs Full PO Receive

Two separate events distinguish partial and full receipt:
- `purchase_order.partially_received` → status = `PARTIALLY_RECEIVED` (some rows received)
- `purchase_order.received` → status = `RECEIVED` (all rows received)

There is no `is_partial` flag in the payload body. The event name is the only indicator. After receiving either event, fetch the PO rows to see per-row received quantities.

The current F1 handler only handles `purchase_order.received`. `purchase_order.partially_received` is not yet implemented (Issue 3).

### manufacturing_order.done Fires After Batch Assignment

The `manufacturing_order.done` webhook fires **after** the user assigns the batch number in the Katana "Done" popup. By the time the webhook arrives, `batch_transactions` ARE populated on the recipe rows. However, there is a short propagation delay — the F4 handler sleeps 15 seconds before fetching ingredients (`Utilities.sleep(15000)`).

### Depleted Batches Return 404

When all stock of a batch is consumed, Katana permanently deletes the batch record. `GET /batch_stocks/{id}` returns 404. There is no undelete. The `include_deleted=true` parameter does not recover these records. Fall back to WASP lot lookup (`waspLookupItemLotAndDate`) to get the lot number and date code from WASP inventory.

### Webhook Test Button

The Katana "Send a test" button shows "Test unsuccessful" even when the webhook is correctly configured. The test payload has no valid `action` field — the handler returns `{status: 'ignored'}`, which Katana interprets as failure. Real events work correctly.

---

## Webhook Payload — Full Example

```javascript
// purchase_order.received
{
  "action": "purchase_order.received",
  "object": {
    "id": 12345
  }
}

// manufacturing_order.done
{
  "action": "manufacturing_order.done",
  "object": {
    "id": 15260040
  }
}

// sales_order.updated (cancel — only status field reliably present on object)
{
  "action": "sales_order.updated",
  "object": {
    "id": 67890,
    "status": "CANCELLED"
  }
}
```

Always treat the `object` as a reference only. Fetch the full resource by ID.

---

## GAS V8 Compatibility Notes

All code in this project runs in Google Apps Script (V8 runtime). GAS V8 has stricter compatibility requirements than Node.js.

| Rule | Correct | Wrong (silently fails) |
|------|---------|----------------------|
| Variable declarations | `var x = 1` | `const x = 1` or `let x = 1` |
| Loops | `for (var i = 0; i < arr.length; i++)` | `for (var item of arr)` |
| String templates | `'hello ' + name` | `` `hello ${name}` `` |
| Arrow functions | `function(x) { return x; }` | `(x) => x` |

**Silent failure mode**: If any file contains `const`, `let`, arrow functions, or `for-of` loops, the entire file fails to load. Every function in that file becomes `undefined`. No error is thrown. Other files that reference functions from the broken file also fail silently.

**Deployment**: Changes to `.gs` files require pushing ALL files together and creating a NEW deployment version. Saving in the editor is not enough to update the live web app URL.

---

## Katana Web URLs (for Activity Log Links)

```javascript
var KATANA_WEB_PATHS = {
  so: '/salesorder/',
  mo: '/manufacturingorders/',
  po: '/purchases/',
  st: '/stocktransfers/'
};
// Usage: 'https://factory.katanamrp.com' + path + id
// e.g.:  https://factory.katanamrp.com/manufacturingorders/15260040
```

---

## Quick Reference: Field Name Variations

Katana is inconsistent about field naming across endpoints. Always check multiple possibilities:

| Concept | Field names to try |
|---------|-------------------|
| Batch number | `batch_number`, `batch_nr`, `nr`, `number` |
| Expiry date | `expiration_date`, `expiry_date`, `best_before_date` |
| Batch ID | `batch_id`, `id` |
| Updated timestamp | `updated_at`, `updatedAt` |
| ST number | `stock_transfer_number` |
| MO consumed qty | `total_consumed_quantity`, `consumed_quantity`, `actual_quantity`, `quantity` |
| Batch transactions | `batch_transactions`, `batchTransactions` |

---

## Error Handling Pattern

```javascript
function katanaApiCall(endpoint) {
  var url = CONFIG.KATANA_BASE_URL + '/' + endpoint;
  var options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.KATANA_API_KEY,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true   // prevents GAS from throwing on 4xx/5xx
  };
  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  if (code === 200) return JSON.parse(response.getContentText());
  return null;  // caller must handle null
}
```

`muteHttpExceptions: true` is required — without it, GAS throws an exception on any non-2xx response, which cannot be caught cleanly. Always check return value for null before accessing fields.

Stock adjustment errors are in `result.response` (the raw response body), not `result.error`. The `.error` field is only set on network-level exceptions (catch block):
```javascript
// Wrong — .error is almost always undefined for API rejections
var errMsg = result.error;

// Correct
var errMsg = result.error || parseKatanaError(result.response) || 'Katana SA rejected';
```
