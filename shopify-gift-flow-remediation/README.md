# Shopify Gift Flow Remediation

## Goal

Stop gift items from being added:

- by the wrong flow
- before fraud or hold review completes
- from overlapping flow paths
- when qualifying quantity rules are not actually met

## Confirmed Defects

1. `Add gift when 2+ 4oz products ordered` uses `shop.name != "Shop App"`.
   This is the wrong field and makes the exclusion ineffective.

2. All gift flows run on `Order created`.
   This is earlier than the separate risk-analysis flow.

3. `Shop App - Gift Trigger - 4oz Jars` uses total order quantity instead of qualifying 4oz quantity.

4. No gift flow has an explicit fraud, hold, or review guard.

## Target State

There must be exactly one path that adds the gift for a given order.

That path must run only after:

- channel/source is correctly identified
- qualifying SKU quantity is correctly evaluated
- fraud review has completed
- the order is not high-risk and not on hold

## Fix Sequence

### P0: Remove overlap immediately

Apply one of these options first:

- safest temporary action: turn off `Add gift when 2+ 4oz products ordered` until the corrected flow is ready
- or patch it immediately by replacing `shop.name != "Shop App"` with `order.channelInformation.app.title != "Shop App"`

Do not leave the current condition in place.

### P0: Give one workflow ownership of gift insertion

Preferred design:

- use source-specific flows only to determine eligibility
- route final gift insertion through one gated flow or one single owned action path

Minimum rule:

- no two active flows should be able to add the same gift variant to the same order

### P0: Move gift execution behind review clearance

Gift insertion must not happen on raw `Order created` if fraud status is still pending.

Preferred design:

- trigger gift logic from a post-review event such as `Order risk analyzed`
- then require `risk is not high`
- then require order is not on hold

If Shopify Flow field availability prevents a clean single-trigger implementation, add an intermediate tag/state and execute gift insertion only after the review workflow clears the order.

### P1: Fix qualifying quantity logic

For `Shop App - Gift Trigger - 4oz Jars`:

- remove the total-order quantity check
- count only qualifying 4oz variant IDs
- require qualifying quantity `>= 2`

Use the same counting approach already proven in the general 4oz flow's code step.

### P1: Verify gift pricing behavior

For both Shop App gift flows:

- confirm whether `Universal Flare Care Gift - 1oz` is priced at `$0.00`
- if not, enable `Add for free`

This is required to avoid turning a gift workflow defect into a billing defect.

### P1: Add explicit write-once protection

Add a guard before gift insertion:

- gift SKU not already present on the order

If available, also add a marker after successful insertion:

- tag or metafield such as `gift_added`

Tags alone are not sufficient as the only protection against races, but they help with observability and repeat suppression.

## Execution Checklist

1. Export screenshots of the current versions of the 3 gift flows.
2. Disable or patch `Add gift when 2+ 4oz products ordered`.
3. Update the Shop App 4oz quantity logic.
4. Move gift execution to post-risk/post-hold clearance.
5. Verify `Add for free` behavior against the actual variant price.
6. Add an explicit "gift not already present" guard.
7. Save screenshots of the final state of each edited flow.

## Validation Plan

## What We Can Prove After Fixes

Yes, we can verify whether the new behavior works as intended with controlled test orders and Flow run logs.

What we can prove:

- the correct flow fires for each channel
- the wrong flow does not fire
- qualifying quantity logic behaves correctly
- high-risk and held orders do not receive a gift
- one qualifying order gets one gift only

What we cannot fully prove from one test pass:

- that no production-only race condition will ever happen
- that no undiscovered hidden flow also edits gifts

Those require a short observation window in production after rollout.

## Required Test Cases

Run these after the fixes are live in a test-safe environment if possible.

1. Non-Shop-App order with 2 qualifying 4oz items
   Expected: gift added once

2. Shop App order with 2 qualifying 4oz items
   Expected: gift added once

3. Shop App order with 1 qualifying 4oz item and 1 unrelated item
   Expected: no gift

4. Shop App bundle order that should receive a gift
   Expected: gift added once and free

5. High-risk order meeting gift criteria
   Expected: no gift

6. Order placed on hold meeting gift criteria
   Expected: no gift while hold remains

7. Duplicate-trigger regression check
   Expected: no order ends with 2 copies of the same gift variant from workflow activity

## Evidence To Collect For Each Test

- order ID
- channel/source
- line items before and after
- risk outcome
- hold status
- relevant Shopify Flow run history
- final tags or markers

## Acceptance Criteria

- no active gift flow uses `shop.name` to detect Shop App
- no gift is added before review clearance
- no gift flow uses total order quantity when qualifying SKU quantity is required
- each eligible order receives at most one gift
- ineligible, high-risk, and held orders receive no gift
- gift line item is free if the business rule says it is a gift

## Rollout Recommendation

Roll out in two stages:

1. ship the P0 overlap and sequencing fixes
2. run the validation matrix and monitor recent Flow runs for 3 to 5 business days

If any unexpected path appears in logs, keep the general flow disabled until ownership is fully consolidated.
