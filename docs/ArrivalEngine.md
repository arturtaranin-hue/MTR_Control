# Arrival Engine

`clsArrivalEngine` calculates production arrival results from already-imported Movement Summary and GPN Angara arrays. Callers must read worksheets once into arrays, then pass those arrays to `CalculateArrivals(plannedRows, movementRows)`.

## Calculated fields

For each planned material, `clsArrivalResult` stores:

- `PlannedQty`
- `ArrivedQty`
- `RemainingQty`
- `CompletionPercent`
- `FirstArrivalDate`
- `LastArrivalDate`
- `DeliveryCount`

## Statuses

The `ArrivalStatus` enum contains the production statuses `ARRIVED`, `PARTIALLY_ARRIVED`, `NOT_ARRIVED`, `OVER_DELIVERED`, and `MANUAL_REVIEW`.

## Business rules

- Multiple deliveries are accumulated into one dictionary bucket per SAP code, material code, or normalized material name.
- Partial deliveries keep positive `RemainingQty` and become `PARTIALLY_ARRIVED`.
- Duplicate movement records are de-duplicated by SAP code, material code, normalized name, quantity, and movement date before totals are calculated.
- Over deliveries set `RemainingQty` to zero and become `OVER_DELIVERED`.
- Missing planned materials or non-positive planned quantities become `MANUAL_REVIEW`.
- Unexpected arrivals that do not match a planned material are appended as `MANUAL_REVIEW` results with `UnexpectedArrival = True`.

## Performance notes

The engine builds `Scripting.Dictionary` indexes for SAP code, material code, normalized material name, and all movement buckets. It processes arrays only and does not loop through worksheet cells.
