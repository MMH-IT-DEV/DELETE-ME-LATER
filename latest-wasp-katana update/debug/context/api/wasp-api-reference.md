# WASP InventoryCloud Public API — Reference

> **Sources:**
> - Official API docs: `https://help.waspinventorycloud.com/Help/API` (Feb 2026)
> - Official Callout docs: `https://help.waspinventorycloud.com/index.htm#t=Callouts_-_List.htm` (Mar 2026)
>   - Direct page (no JS): `https://help.waspinventorycloud.com/Callouts_-_List.htm`
>   - Sub-pages (403 from outside tenant — must be logged in): `Callouts_-_Add_Edit_-_Set_Triggers.htm`, `Callouts_-_Add_Edit_-_Configure_Callout.htm`, `Callouts_-_Variables.htm`, `Callouts_-_Security.htm`
> - Empirical knowledge from production code: `src/05_WaspAPI.gs`, `src/03_WaspCallouts.gs`, `src/10_InventorySync.gs`, `src/11_SyncHelpers.gs`, `src/13_Adjustments.gs`
> - Screenshots of live WASP callout config (Mar 2026): `context/wasp-remove-callout-*.png`
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

WASP's callout system sends an outbound HTTP request to a configured URL when specific
events occur. Callouts are configured in the WASP UI under **Settings → Callouts**
(not via API). Each callout has a name, status (Enabled/Disabled), trigger rules,
HTTP config, and optional security.

**Callout docs URL (must be logged into WASP tenant to access sub-pages):**
`https://help.waspinventorycloud.com/index.htm#t=Callouts_-_List.htm`

There is **no reversal event** — WASP does not fire a callout when a transaction is
undone (there is no undo feature anyway).

---

### 11.1 Callout Configuration

Each callout is configured across three tabs:

**Set Triggers tab:**
Defines which events fire the callout. Event types include: `move`, `check out`,
`check in`, `created`, `add`, `remove`, and others. Triggers can be scoped to:
any item, a specific item, a specific location, a specific customer, or a specific
purchase order.

**Configure Callout tab:**
- **Method**: GET / POST / PUT / DELETE
- **URL**: supports embedded template variables anywhere in the string
- **Headers**: custom key-value pairs. Three static headers are **always** added
  automatically by WASP (see section 11.3)
- **Values / Request Body**: available for POST/PUT. Supports JSON, XML, or Form Data.
  Template variables can be used anywhere in the body. Validation button available.

**Security tab:**
- **Private Key**: a secret string used to generate the `Wasp-Callout-Signature` header.
  Can be regenerated at any time. Rotating it invalidates all previous signatures.
- The receiver should use this key to verify the HMAC signature on inbound callouts.
  The debug script currently does NOT verify signatures (ignores the header).

---

### 11.2 Static Headers Always Sent

Every callout request includes these three headers automatically — they cannot be removed:

| Header | Value |
|--------|-------|
| `Wasp-Callout-Name` | The callout's configured name (e.g. `Katana Sync - Item Removed`) |
| `Wasp-Callout-Timestamp` | ISO 8601 timestamp of when the callout fired (e.g. `2026-02-24T21:33:06.2411895Z`) |
| `Wasp-Callout-Signature` | HMAC signature generated using the configured Private Key |

Plus any custom headers configured in the Headers tab (e.g. `Content-Type: application/json`).

---

### 11.3 Template Variables

WASP uses a `{trans.FieldName}` template engine to inject transaction data into the
callout URL, headers, and body. Variables are substituted at send time.

**Known variables (confirmed from live config screenshots — Mar 2026):**

| Variable | Description |
|----------|-------------|
| `{trans.AssetTag}` | Item SKU / ItemNumber (confusingly named "AssetTag" in WASP) |
| `{trans.AssetTransQuantity}` | Transaction quantity (as string — must `parseFloat`) |
| `{trans.SiteName}` | WASP site name (e.g. `MMH Kelowna`) |
| `{trans.LocationCode}` | WASP location code (e.g. `SHOPIFY`) |
| `{trans.Lot}` | Lot number — may be unresolved if item is not lot-tracked (see 11.5) |
| `{trans.DateCode}` | Expiry date — may be unresolved if item is not lot-tracked (see 11.5) |
| `{trans.Notes}` | Notes entered at time of transaction. **See 11.6 for critical API note.** |
| `{trans.Assignee}` | User who performed the transaction |
| `{trans.AssetTransDate}` | Date of the transaction |
| `{trans.TransTypeDescription}` | Transaction type description |
| `{trans.AssetDescription}` | Item description |
| `{trans.SerialNumber}` | Serial number (if applicable) |
| `{trans.TransactionDueDate}` | Due date (for check-out transactions) |

Variables not listed here may also exist — the full list requires access to the
WASP Variables sub-page (403 from outside the tenant).

---

### 11.4 Live Callout Configuration — "Katana Sync - Item Removed"

This is the production callout configured for the `quantity_removed` event.
Screenshots saved in `context/wasp-remove-callout-*.png`.

- **Name**: Katana Sync - Item Removed
- **Status**: Disabled (as of 2026-03-18 — enabling for testing)
- **Method**: POST
- **URL**: debug webhook (`https://script.google.com/macros/s/AKfycbw4Z2Ytiz7hlkS-48fi6f5Dswvzafmgqy.../exec`)
- **Trigger**: item removed transaction

**Request Body (JSON template):**
```json
{
  "action": "quantity_removed",
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

**Headers (custom):**
- `Content-Type: application/json`
- Plus the 3 static headers (Wasp-Callout-Name, Wasp-Callout-Timestamp, Wasp-Callout-Signature)

---

### 11.5 `quantity_removed` Payload (as received by the webhook)

Fired when inventory is removed from WASP (manual removal, pick order fulfillment, API call, etc.).

```json
{
  "action":       "quantity_removed",
  "ItemNumber":   "SKU-001",
  "Quantity":     "5",
  "SiteName":     "MMH Kelowna",
  "LocationCode": "SHOPIFY",
  "Lot":          "LOT-2024-A",
  "DateCode":     "2026-06-30",
  "Notes":        "User notes or API-provided notes",
  "Assignee":     "john.smith"
}
```

**Field notes:**

| Field | Notes |
|-------|-------|
| `ItemNumber` | Maps from `{trans.AssetTag}` — this IS the SKU |
| `Quantity` | String — must `parseFloat`. Maps from `{trans.AssetTransQuantity}` |
| `Notes` | Maps from `{trans.Notes}`. See 11.6 about API-triggered transactions |
| `Lot` / `DateCode` | May be unresolved template strings if item not lot-tracked — see 11.7 |
| `Assignee` | WASP also sends user under `UserName`, `User`, `ModifiedBy`, `PerformedBy` in some configs |
| `PickOrderNumber` / `OrderNumber` | Present when removal was triggered by pick order fulfillment. If present → route to pick handler, not F2 |

**Note on `action` field:** The `"action": "quantity_removed"` field is part of our
custom payload template, not something WASP adds. It must be explicitly included in
the callout body template for routing to work in `doPost`.

---

### 11.6 Critical: Does `{trans.Notes}` Work for API-Triggered Transactions?

**Unconfirmed as of 2026-03-18.** This is the key question for F2 echo suppression.

When the sync sheet calls `waspRemoveInventory(..., 'Sheet push qty decrease 2026-03-18', ...)`:
- WASP receives the API call with that notes value
- WASP fires the callout
- **Does `{trans.Notes}` substitute with `'Sheet push qty decrease 2026-03-18'`?**

If YES → `isInternalWaspAdjustmentNote_()` catches it → F2 suppresses the echo ✅
If NO (empty or different) → suppression fails → F2 creates duplicate Katana SA ❌

The `enginPreMark_` cache mechanism (120s TTL) is the backup suppression layer that
doesn't depend on notes at all. It works as long as `GAS_WEBHOOK_URL` is set in the
sync script's properties.

**To confirm:** enable the callout, make a sheet-based adjustment, check the Activity
log and the "Recent Calls" section of the WASP callout config.

---

### 11.7 Unresolved Template Variable Issue

WASP substitutes `{trans.Lot}` and `{trans.DateCode}` at send time. If the item has
no lot tracking configured in WASP, WASP may send the **literal unresolved template
string** instead of an empty value.

**Example of broken payload:**
```json
{
  "Lot":      "{trans.Lot}",
  "DateCode": "{trans.DateCode}"
}
```

**Detection — always sanitize before using:**
```javascript
var lot    = payload.Lot    || payload.LotNumber || payload.BatchNumber || '';
var expiry = payload.DateCode || payload.ExpiryDate || payload.ExpDate  || '';

if (lot    && lot.indexOf('{')    >= 0) lot    = '';
if (expiry && expiry.indexOf('{') >= 0) expiry = '';
```

---

### 11.8 Recent Calls Monitoring

The WASP callout UI shows a **Recent Calls** log with:
- Notice Date, Status, Status Date, Retry count
- Statuses: Created, Success, Pending Retry, Failed
- Expandable per-attempt detail with HTTP status codes

This is useful for confirming whether callouts are firing and what responses the
webhook returned. Failed callouts can be retried manually from the UI.

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

*Last updated: 2026-03-18. Section 11 expanded with callout configuration details, template variables, static headers, signature mechanism, and live production callout config from screenshots.*
