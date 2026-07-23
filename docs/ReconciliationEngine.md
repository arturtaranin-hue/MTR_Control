# Reconciliation Engine

`clsReconciliationService` combines the DV Plan, Movement Summary, and GPN Angara arrivals into a final `Collection` of `clsMaterialState` objects. The service works from in-memory arrays so callers can load workbook ranges once and avoid per-cell reconciliation loops.

## Inputs

All three inputs use `clsImportMapping` column definitions by default:

1. SAP code
2. Material code
3. Material name
4. Quantity
5. Arrival date

Call `Initialize` with custom mappings when a workbook layout differs between the DV Plan, Movement Summary, and GPN Angara files.

## Matching order

Arrivals are indexed once and then matched to every planned material by:

1. SAP code
2. Material code
3. Normalized material name

Movement Summary and GPN Angara quantities are retained separately as `MovementQty` and `AngaraQty`, then combined into `ArrivedQty`.

## Status rules

| Status | Rule |
| --- | --- |
| `DELIVERED` | Arrived quantity equals planned quantity. |
| `PARTIAL` | Arrived quantity is greater than zero but less than planned quantity. |
| `OVERDELIVERED` | Arrived quantity is greater than planned quantity. |
| `NOT_FOUND` | No matching arrival was found for a positive planned quantity. |
| `MANUAL_REVIEW` | The record needs user validation, including duplicate arrivals, unknown materials, or invalid planned quantities. |
| `PLANNED` | Reserved for callers that need to mark future planned-only material before arrival checks run. |

## Calculated fields

For each `clsMaterialState`:

- `PlannedQty` comes from the DV Plan.
- `ArrivedQty` is the combined Movement Summary plus GPN Angara quantity.
- `RemainingQty` is `PlannedQty - ArrivedQty` and is clamped to zero for overdelivery.
- `DeliveryPercent` is `ArrivedQty / PlannedQty` when the planned quantity is positive.

## Detection rules

- **Duplicate arrivals**: identical SAP code, material code, normalized name, quantity, and arrival date appearing more than once in the combined arrival feed.
- **Negative balance**: calculated remaining quantity is below zero before overdelivery clamping.
- **Over delivery**: arrived quantity exceeds planned quantity.
- **Unknown material**: an arrival bucket does not match any DV Plan row and is added as a manual-review state.
- **Manual review**: invalid planned quantity, duplicate arrivals, and unknown materials are flagged with `ManualReview = True` and a `ReviewReason`.

## Reporting

`clsReportBuilder` converts the final state collection into:

- a summary dictionary with counts and totals;
- a two-dimensional row array suitable for writing to a report worksheet.
