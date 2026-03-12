# WASP InventoryCloud Public API — Reference

> **Sources:** Official docs at `https://help.waspinventorycloud.com/Help/API` (Feb 2026) +
> empirical knowledge from production code in `src/05_WaspAPI.gs`, `src/03_WaspCallouts.gs`,
> `src/10_InventorySync.gs`, `src/11_SyncHelpers.gs`, `src/13_Adjustments.gs`.
>
> Where docs and code diverge, **code is authoritative** — it reflects what the live API
> actually accepts and returns.

---

## Table of Contents

1. [Base URL and Authentication](#1-base-url-and-authentication)
2. [General Request Rules](#2-general-request-rules)
3. [Success and Error Detection](#3-success-and-error-detection)
4. [Transaction Endpoints](#4-transaction-endpoints)
   - [Add](#41-add-inventory)
   - [Remove](#42-remove-inventory)
   - [Adjust](#43-adjust-inventory)
   - [Move](#44-move-inventory)
   - [Check-In / Check-Out](#45-check-in--check-out)
5. [Audit / Reconcile Endpoints](#5-audit--reconcile-endpoints)
6. [Inventory Search Endpoints](#6-inventory-search-endpoints)
   - [inventorysearch (simple)](#61-inventorysearch-simple)
   - [advancedinventorysearch (paged bulk)](#62-advancedinventorysearch-paged-bulk)
7. [Item Info Search Endpoints](#7-item-info-search-endpoints)
8. [Pick Pack Ship Order Endpoints](#8-pick-pack-ship-order-endpoints)
9. [Adjust Reason Code Endpoints](#9-adjust-reason-code-endpoints)
10. [Transaction History Endpoints](#10-transaction-history-endpoints)
11. [Callout (Webhook) System](#11-callout-webhook-system)
    - [quantity_added](#111-quantity_added-payload)
    - [quantity_removed](#112-quantity_removed-payload)
    - [Unresolved template variable issue](#113-unresolved-template-variable-issue)
12. [Lot-Level Targeting — Patterns and Rules](#12-lot-level-targeting--patterns-and-rules)
13. [Field Name Variants Reference](#13-field-name-variants-reference)
14. [No-Go List — Things WASP Cannot Do](#14-no-go-list--things-wasp-cannot-do)
15. [Rate Limiting and Batch Limits](#15-rate-limiting-and-batch-limits)

---

## 1. Base URL and Authentication

```
Base URL:  https://<tenant>.waspinventorycloud.com
Example:   https://mymagichealer.waspinventorycloud.com
```

Every request requires a Bearer token in the `Authorization` header.

```
Authorization: Bearer <WASP_TOKEN>
Content-Type:  application/json
```

The token is a long-lived API key configured in WASP's admin UI and stored as a script property. It does not expire on a schedule but can be revoked from the UI.

---

## 2. General Request Rules

| Rule | Detail |
|------|--------|
| All transaction endpoints are `POST` | Even fetches (search, history) use `POST` |
| Transaction body is always a JSON **array** | Wrap a single record in `[{ ... }]` — sending a plain object returns a 400 or silent error |
| Search/history body is a JSON **object** | `inventorysearch`, `advancedinventorysearch`, history endpoints take a plain `{ }` object |
| Batch limit | Up to **500 records** per transaction array |
| Quantities are always in stocking units | The API ignores alternate unit conversions |
| All mutations are **relative**, never absolute | There is no "set to X" endpoint |

---

## 3. Success and Error Detection

WASP does not reliably use HTTP status codes to signal failures. A failed transaction can return HTTP 200 with an error body.

**Correct success check (both conditions must be true):**

```javascript
var success = (response.getResponseCode() === 200)
           && !body.includes('"HasError":true');
```

**Error response structure:**
```json
{
  "HasError": true,
  "Message": "Item not found: SKU-999",
  "Data": null
}
```

**Success response structure:**
```json
{
  "HasError": false,
  "Message": null,
  "Data": [ ... ]
}
```

Never rely on HTTP 200 alone. Always check `HasError`.

---

## 4. Transaction Endpoints

### 4.1 Add Inventory

```
POST /public-api/transactions/item/add
```

Adds quantity to an item. For lot-tracked items, `Lot` and `DateCode` are required.

**Payload — without lot tracking:**
```json
[{
  "ItemNumber":   "SKU-001",
  "Quantity":     10,
  "SiteName":     "MMH Kelowna",
  "LocationCode": "RECEIVING-DOCK",
  "Notes":        "PO Received: PO-542"
}]
```

**Payload — with lot tracking:**
```json
[{
  "ItemNumber":   "SKU-001",
  "Quantity":     50,
  "SiteName":     "MMH Kelowna",
  "LocationCode": "RECEIVING-DOCK",
  "Lot":          "LOT-2024-A",
  "DateCode":     "2026-06-30",
  "Notes":        "PO Received: PO-542"
}]
```

**Field reference:**

| Field | Type | Required | Notes |
|-------|------|----------|-------|
| `ItemNumber` | string | Yes | The item's SKU/part number as configured in WASP |
| `Quantity` | number | Yes | Always positive. Use remove endpoint to subtract |
| `SiteName` | string | Yes | Exact site name as configured in WASP |
| `LocationCode` | string | Yes | Exact location code within that site |
| `Lot` | string | Conditional | Required when item has lot tracking enabled in WASP |
| `DateCode` | string | Conditional | Required when item has lot tracking. Format: `YYYY-MM-DD`. This is the **expiry date**, not a manufacturing code |
| `Notes` | string | No | Free text, visible in WASP transaction history |

**Permission required:** Allow Add

---

### 4.2 Remove Inventory

```
POST /public-api/transactions/item/remove
```

Removes quantity from an item. For lot-tracked items, WASP needs `Lot` (and optionally `DateCode`) to identify which lot record to decrement. Without these, WASP may refuse or remove from the wrong lot.

**Payload — without lot tracking:**
```json
[{
  "ItemNumber":   "SKU-001",
  "Quantity":     5,
  "SiteName":     "MMH Kelowna",
  "LocationCode": "SHOPIFY",
  "Notes":        "Correction"
}]
```

**Payload — with lot tracking:**
```json
[{
  "ItemNumber":   "SKU-001",
  "Quantity":     5,
  "SiteName":     "MMH Kelowna",
  "LocationCode": "SHOPIFY",
  "Lot":          "LOT-2024-A",
  "DateCode":     "2026-06-30",
  "Notes":        "Correction"
}]
```

**Field reference:** Same fields as Add. `Quantity` is always positive (it represents the magnitude of the removal).

**Key pattern from codebase:** When zeroing lot-tracked items during a full sync, use `remove` (not `adjust`). The `adjust` endpoint may not correctly target a specific lot record.

**Permission required:** Allow Remove

---

### 4.3 Adjust Inventory

```
POST /public-api/transactions/item/adjust
```

Applies a signed delta to quantity. Positive increases, negative decreases. This is the best endpoint for syncing discrepancies when you know the exact difference needed.

**Payload:**
```json
[{
  "ItemNumber":   "SKU-001",
  "Quantity":     -3,
  "SiteName":     "MMH Kelowna",
  "LocationCode": "SHOPIFY",
  "Notes":        "Katana sync 2026-02-26"
}]
```

**Field reference:** Same fields as Add. `Quantity` is signed (+/-).

**Docs say** "An adjust type and reason must be provided" but the production codebase never sends `AdjustType` or `ReasonCode` and calls succeed. These may be optional in the public API tier, or may default to a system value. If WASP starts rejecting adjust calls, try fetching valid reason codes from the reason endpoints (section 9) and adding a `ReasonCode` field.

**Do not use `adjust` to target a specific lot.** Use `remove` then `add` for lot-level corrections. The `adjust` endpoint works reliably for non-lot-tracked items.

**Permission required:** Allow Adjust

---

### 4.4 Move Inventory

```
POST /public-api/transactions/item/move
```

Moves quantity from one location to another within or across sites. Not currently used in the codebase but documented by WASP.

**Payload (inferred from WASP docs pattern):**
```json
[{
  "ItemNumber":      "SKU-001",
  "Quantity":        10,
  "SiteName":        "MMH Kelowna",
  "LocationCode":    "RECEIVING-DOCK",
  "ToSiteName":      "MMH Kelowna",
  "ToLocationCode":  "PRODUCTION",
  "Notes":           "Transfer to production"
}]
```

**Permission required:** Allow Move

---

### 4.5 Check-In / Check-Out

```
POST /public-api/transactions/item/check-in
POST /public-api/transactions/item/check-out
```

Used for asset-style tracking where items are loaned to customers or vendors. Requires either a `CustomerCode` or `VendorCode` (not both). Not used in this codebase.

**Permission required:** Allow Check In / Allow Check Out

---

## 5. Audit / Reconcile Endpoints

Used for physical count workflows. The audit-count records discrepancies; the reconcile endpoint resolves them.

| Endpoint | Method | URL |
|----------|--------|-----|
| Audit Count | POST | `/public-api/transactions/public-api/audit-count` |
| Audit Count v2 (container support) | POST | `/public-api/transactions/audit-count-v2` |
| Reconcile | POST | `/public-api/transactions/public-api/reconcile` |
| Reconcile v2 (container support) | POST | `/public-api/transactions/reconcile-v2` |

**Permission required:** Allow Audit

Reconcile is the closest thing WASP has to an "absolute set" operation — it can accept or reject discrepancies found during an audit count. However it requires a prior audit-count to exist in WASP before reconcile can be called.

---

## 6. Inventory Search Endpoints

### 6.1 inventorysearch (simple)

```
POST /public-api/ic/item/inventorysearch
```

Returns inventory records for a specific item, optionally filtered by site and/or location. Analogous to the Edit Item screen in the WASP UI. Returns all matching records (no pagination needed for a single item).

**Request:**
```json
{
  "ItemNumber":   "SKU-001",
  "SiteName":     "MMH Kelowna",
  "LocationCode": "SHOPIFY"
}
```

All three fields are optional filters. Omit `LocationCode` to get all locations for that item and site.

**Response:**
```json
{
  "HasError": false,
  "Data": [
    {
      "ItemNumber":    "SKU-001",
      "SiteName":      "MMH Kelowna",
      "LocationCode":  "SHOPIFY",
      "Location":      "SHOPIFY",
      "Quantity":      42,
      "Lot":           "LOT-2024-A",
      "DateCode":      "2026-06-30"
    }
  ]
}
```

**Known response fields (observed in production):**

| Field | Notes |
|-------|-------|
| `ItemNumber` | SKU |
| `SiteName` | Site name |
| `LocationCode` | Location code (also appears as `Location`) |
| `Quantity` | Current quantity on hand |
| `Lot` | Lot number (also appears as `LotNumber`, `BatchNumber`) |
| `DateCode` | Expiry date (also appears as `ExpiryDate`, `ExpDate`) |

**Best use case:** Looking up the current lot and expiry for a known SKU at a known location before a targeted remove. Also returns `Quantity` for QOH checks.

**Permission required:** Allow Item; enforces Role Site permission

---

### 6.2 advancedinventorysearch (paged bulk)

```
POST /public-api/ic/item/advancedinventorysearch
```

Paged query similar to the main View Item grid in the WASP UI. Designed for bulk inventory export.

**Request — search by pattern:**
```json
{
  "SearchPattern": "SKU-001",
  "PageSize":      100,
  "PageNumber":    1
}
```

**Request — dump all inventory (no filter):**
```json
[{}]
```

Passing an empty array-wrapped object returns all inventory records across all items and sites. This is the most efficient way to get a full inventory snapshot. Paginate until a page returns fewer results than `PageSize`.

**CRITICAL:** `SearchPattern` is a **text search**, not an exact ItemNumber filter. It does substring matching. You must filter results client-side by `rec.ItemNumber === targetSku` to isolate a specific item.

**Response:**
```json
{
  "HasError": false,
  "Data": [
    {
      "ItemNumber":    "SKU-001",
      "SiteName":      "MMH Kelowna",
      "LocationCode":  "PRODUCTION",
      "QuantityOnHand": 15,
      "Quantity":       15,
      "Lot":            "LOT-2024-B",
      "DateCode":       "2026-09-01"
    }
  ]
}
```

**Quantity field:** Use `QuantityOnHand` first, fall back to `Quantity`.

**Pagination pattern:**
```javascript
for (var page = 1; page <= 20; page++) {
  var result = waspApiCall(url, { SearchPattern: sku, PageSize: 100, PageNumber: page });
  var rows = JSON.parse(result.response).Data || [];
  if (rows.length === 0) break;
  // process rows...
  if (rows.length < 100) break; // last page
}
```

**Permission required:** Allow Item; enforces Role Site permission

---

## 7. Item Info Search Endpoints

Used to look up item master data (not inventory quantities).

| Endpoint | Method | URL | Notes |
|----------|--------|-----|-------|
| Info Search | POST | `/public-api/ic/item/infosearch` | Matches by ItemNumber, AltItemNumber, or Description text |
| Advanced Info Search | POST | `/public-api/ic/item/advancedinfosearch` | Paged version, returns richer item data |

**infosearch request:**
```json
{
  "ItemNumber":      "SKU-001",
  "AltItemNumber":   "",
  "ItemDescription": ""
}
```

Returns all active items matching any of the provided fields. Useful for resolving SKU → WASP item ID before transactions.

---

## 8. Pick Pack Ship Order Endpoints

| Endpoint | Method | URL |
|----------|--------|-----|
| Create pick order | POST | `/public-api/ic/pickpackshiporder/create` |
| Get orders by number | POST | `/public-api/ic/pickpackshiporder/getordersbynumber` |

**Get by number request:**
```json
{
  "OrderNumbers": ["PICK-2024-001"]
}
```

**Response fields used in codebase:**
```json
{
  "Data": [{
    "PickOrderLines": [
      {
        "ItemNumber":          "SKU-001",
        "Quantity":            10,
        "OutstandingQuantity": 0
      }
    ]
  }]
}
```

An order is fully picked when all `OutstandingQuantity` values are `0`.

---

## 9. Adjust Reason Code Endpoints

WASP requires reason codes for adjustments (though the public API tier may not enforce this in practice). Use these endpoints to fetch valid codes before submitting adjustments if you need to include a reason.

| Endpoint | Method | URL | Returns |
|----------|--------|-----|---------|
| Adjust-down reasons | GET | `/public-api/transactions/item/adjustDown/reasons` | All active down reasons |
| Adjust-up reasons | GET | `/public-api/transactions/item/adjustUp/reasons` | All active up reasons |
| Adjust-count reasons | GET | `/public-api/transactions/item/adjustCount/reasons` | All active count reasons |

**Permission required:** Allow Adjust Reason Code or Allow Adjust

Response structure not confirmed from docs; likely an array of `{ Code, Description }` objects.

---

## 10. Transaction History Endpoints

WASP provides several ways to query past transactions. None of these are currently called in the codebase — all patterns below come from docs alone.

| Endpoint | Method | URL | Notes |
|----------|--------|-----|-------|
| History (public API) | POST | `/public-api/transactions/grid-query/history-public-api` | Current recommended endpoint |
| History notes | POST | `/public-api/transactions/grid-query/history-notes-public-api` | Returns attachments/notes given a list of transaction IDs |
| History v2 | POST | `/public-api/transactions/grid-query/transaction-history-v2` | Newer version; omits internal ID fields; supports page/filter/sort |
| History v1 (deprecated) | POST | `/public-api/transactions/grid-query/transaction-history` | Deprecated |
| Stream CSV | POST | `/public-api/transactions/streamgridrequestcsv` | Streams ALL history as CSV; no pagination |
| Stream Archive CSV | POST | `/public-api/transactions/streamgridrequestarchivecsv` | Streams archived history as CSV |

**History v2 request (inferred from docs):**
```json
{
  "PageNumber": 1,
  "PageSize":   100,
  "Filter": {
    "ItemNumber": "SKU-001"
  },
  "Sort": {}
}
```

The exact filter field names are **not published** in the help site HTML. The filter structure likely mirrors WASP's internal grid query model. To find usable filters, either:
- Test the endpoint with known values and observe what fields WASP accepts
- Check if the WASP UI's network requests show the filter payload when searching transaction history

**Filtering by Notes/reference:** Unknown whether WASP supports filtering history by the `Notes` field. Treat this as unconfirmed until tested.

**History notes endpoint request:**
```json
{
  "TransactionHistoryIds": [12345, 12346]
}
```

Returns binary attachment data and user notes for specific transactions.

---

## 11. Callout (Webhook) System

WASP's callout system sends an HTTP POST to a configured URL when specific events occur in WASP. Callouts are configured in the WASP UI (not via API). The event type is typically encoded in the request URL path or a header that the receiving server uses to route the event.

There is **no reversal event** — WASP does not fire a callout when a transaction is undone (there is no undo feature anyway).

### 11.1 `quantity_added` Payload

Fired when inventory is added to WASP (via UI, scanner, or API).

```json
{
  "AssetTag":          "SKU-001",
  "Quantity":          "10",
  "LocationCode":      "RECEIVING-DOCK",
  "SiteName":          "MMH Kelowna",
  "AssetDescription":  "User notes entered at time of transaction",
  "Notes":             "User notes entered at time of transaction",
  "Lot":               "LOT-2024-A",
  "DateCode":          "2026-06-30",
  "Assignee":          "john.smith",
  "UserName":          "john.smith",
  "User":              "john.smith",
  "ModifiedBy":        "john.smith",
  "PerformedBy":       "john.smith"
}
```

**Field notes:**

| Field | Notes |
|-------|-------|
| `AssetTag` | This is the **item's SKU/ItemNumber** — not a barcode or physical asset tag. Confusing name. |
| `Quantity` | Arrives as a **string**, not a number. Must parse: `parseFloat(payload.Quantity)` |
| `LocationCode` | The WASP location code where inventory was added |
| `SiteName` | The WASP site name |
| `AssetDescription` / `Notes` | Both may carry the user-entered notes. Try `AssetDescription` first, fall back to `Notes` |
| `Lot` | Lot number — may be an unresolved template (see section 11.3) |
| `DateCode` | Expiry date — may be an unresolved template (see section 11.3) |
| `Assignee` / `UserName` / `User` / `ModifiedBy` / `PerformedBy` | WASP sends the user under one of these fields depending on configuration. Try all as fallbacks |

---

### 11.2 `quantity_removed` Payload

Fired when inventory is removed from WASP (manual removal, pick order fulfillment, etc.).

```json
{
  "AssetTag":         "SKU-001",
  "Quantity":         "5",
  "LocationCode":     "SHOPIFY",
  "SiteName":         "MMH Kelowna",
  "AssetDescription": "User notes",
  "Notes":            "User notes",
  "PickOrderNumber":  "PICK-2024-001",
  "OrderNumber":      "PICK-2024-001",
  "Lot":              "LOT-2024-A",
  "DateCode":         "2026-06-30",
  "Assignee":         "john.smith"
}
```

**Additional fields vs. quantity_added:**

| Field | Notes |
|-------|-------|
| `PickOrderNumber` | Present when the removal was triggered by pick order fulfillment. Absent for manual removals |
| `OrderNumber` | Same value as `PickOrderNumber` — WASP sends it under both names |

**Routing logic:** If `PickOrderNumber` (or `OrderNumber`) is present, treat the event as a pick order completion, not a general manual removal. These require different downstream handling (marking SO delivered in Katana, triggering ShipStation, etc.).

---

### 11.3 Unresolved Template Variable Issue

WASP uses a template engine for callout payloads. When configured to include `{trans.Lot}` or `{trans.DateCode}` in the payload, these are normally substituted at send time. However if the item has no lot tracking configured in WASP, or the template mapping is misconfigured, WASP sends the **literal unresolved template string** instead of an empty value.

**Example of broken payload:**
```json
{
  "Lot":      "{trans.Lot}",
  "DateCode": "{trans.DateCode}"
}
```

**Detection and fix — always sanitize before using:**
```javascript
var lot    = payload.Lot    || payload.LotNumber || payload.BatchNumber || '';
var expiry = payload.DateCode || payload.ExpiryDate || payload.ExpDate  || '';

// Discard unresolved WASP template variables
if (lot    && lot.indexOf('{')    >= 0) lot    = '';
if (expiry && expiry.indexOf('{') >= 0) expiry = '';
```

Check for the `{` character as a sentinel. Any lot or expiry value containing `{` is a failed template substitution and should be treated as absent.

---

## 12. Lot-Level Targeting — Patterns and Rules

### When does a lot need to be specified?

If an item is configured in WASP with lot tracking enabled, all transactions against that item must include the `Lot` field. WASP will reject or misbehave if you omit it.

### How to look up the current lot before a remove

Use `inventorysearch` to find the active lot at a specific location:

```javascript
function getWaspLotInfo(siteName, itemNumber, locationCode) {
  var payload = {
    ItemNumber:   itemNumber,
    SiteName:     siteName,
    LocationCode: locationCode  // optional
  };
  var result = waspApiCall('/public-api/ic/item/inventorysearch', payload);
  var data = JSON.parse(result.response).Data || [];

  for (var i = 0; i < data.length; i++) {
    var rec = data[i];
    var lot    = rec.Lot    || rec.LotNumber || rec.BatchNumber || '';
    var expiry = rec.DateCode || rec.ExpiryDate || rec.ExpDate  || '';
    if (lot) return { lot: lot, expiry: expiry, quantity: rec.Quantity || 0 };
  }
  return { lot: '', expiry: '' };
}
```

### Correcting a lot-tracked item's quantity (no "adjust to X")

WASP has no absolute-set endpoint. To correct a lot-tracked item to a target quantity:

1. Look up current lot and quantity via `inventorysearch`
2. Calculate delta = target - current
3. If delta > 0: call `/add` with `Lot` + `DateCode`
4. If delta < 0: call `/remove` with `Lot` + `DateCode`, quantity = `Math.abs(delta)`

Do not use `/adjust` for lot-tracked items — it may not correctly decrement the target lot record.

### What DateCode actually means

Despite the name, `DateCode` in WASP is the **expiry date** (best-before / expiration date), not a manufacturing date code. Format as `YYYY-MM-DD`. Confirmed by the field being populated from Katana's `expiry_date` / `best_before_date` fields and by the codebase defaulting it to 2 years from now when not provided:

```javascript
var future = new Date();
future.setFullYear(future.getFullYear() + 2);
dateCode = future.toISOString().slice(0, 10); // "YYYY-MM-DD"
```

---

## 13. Field Name Variants Reference

WASP uses inconsistent field names across different endpoints and payload directions (inbound vs. response). Always try the primary name first, then fall back through alternates.

### Lot number

| Context | Field names to try (in order) |
|---------|-------------------------------|
| Transaction payload (you send) | `Lot` |
| inventorysearch response | `Lot`, `LotNumber`, `BatchNumber` |
| advancedinventorysearch response | `Lot`, `LotNumber`, `BatchNumber` |
| Callout payload (WASP sends) | `Lot`, `LotNumber`, `BatchNumber` |

```javascript
var lot = rec.Lot || rec.LotNumber || rec.BatchNumber || '';
```

### Expiry date (DateCode)

| Context | Field names to try (in order) |
|---------|-------------------------------|
| Transaction payload (you send) | `DateCode` |
| inventorysearch response | `DateCode`, `ExpiryDate`, `ExpDate` |
| Callout payload (WASP sends) | `DateCode`, `ExpiryDate`, `ExpDate` |

```javascript
var expiry = rec.DateCode || rec.ExpiryDate || rec.ExpDate || '';
```

### Item number (SKU)

| Context | Field names to try (in order) |
|---------|-------------------------------|
| Transaction payload (you send) | `ItemNumber` |
| Search response | `ItemNumber`, `itemNumber` |
| Callout payload (WASP sends) | `AssetTag` (this is the SKU, not a physical tag) |

### Location

| Context | Field names to try (in order) |
|---------|-------------------------------|
| Transaction payload (you send) | `LocationCode` |
| Search response | `LocationCode`, `Location` |
| Callout payload (WASP sends) | `LocationCode` |

### Quantity on hand

| Context | Field names to try (in order) |
|---------|-------------------------------|
| advancedinventorysearch response | `QuantityOnHand`, `Quantity` |
| inventorysearch response | `Quantity` |
| Callout payload (WASP sends) | `Quantity` (string — must `parseFloat`) |

### User identity

| Context | Field names to try (in order) |
|---------|-------------------------------|
| Callout payload (WASP sends) | `Assignee`, `UserName`, `User`, `ModifiedBy`, `PerformedBy` |

```javascript
var user = payload.Assignee || payload.UserName || payload.User
        || payload.ModifiedBy || payload.PerformedBy || '';
```

### Notes / description

| Context | Field names to try (in order) |
|---------|-------------------------------|
| Callout payload (WASP sends) | `AssetDescription`, `Notes` |

```javascript
var notes = payload.AssetDescription || payload.Notes || '';
```

### Pick order number

| Context | Field names to try (in order) |
|---------|-------------------------------|
| quantity_removed callout payload | `PickOrderNumber`, `OrderNumber` |

---

## 14. No-Go List — Things WASP Cannot Do

| Capability | Status | Workaround |
|------------|--------|-----------|
| Set inventory to an absolute quantity ("adjust to X") | Not available | Calculate delta vs. current QOH; use `/add` or `/remove` |
| Reverse or void a previous transaction | Not available | Post a counter-transaction of equal and opposite quantity |
| Filter transaction history by Notes field | Unknown/unconfirmed | Fetch pages and filter client-side |
| Adjust a specific lot via `/adjust` endpoint | Unreliable | Use `/remove` + `/add` for lot-level corrections |
| Receive callout events for API-triggered transactions | Unknown | WASP callouts may or may not fire for programmatic adds/removes — not confirmed |
| Query history by reference/notes | Not documented | No known endpoint; would require CSV stream and client-side parse |

---

## 15. Rate Limiting and Batch Limits

| Setting | Value | Notes |
|---------|-------|-------|
| Max records per transaction call | 500 | Enforced by WASP |
| Observed safe delay between calls | 300 ms | Used in production sync (`SYNC_CONFIG.RATE_LIMIT_MS`) |
| Max pagination safety limit | 20 pages | Self-imposed in `waspLookupItemLots` to prevent runaway loops |
| Batch window for callout grouping | 4 seconds | Multiple callouts within 4s are grouped into one Katana SA |
| Sync dedup cache TTL | 120 seconds | After syncing an item, ignore WASP callouts for that SKU+location for 2 minutes (loop prevention) |

There is no published rate limit from WASP. The 300 ms delay is empirically safe. Do not drop below ~100 ms for bulk operations to avoid transient failures.

---

*Last updated: 2026-02-26. Regenerate if WASP API changes or new endpoints are discovered.*
