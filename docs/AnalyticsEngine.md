# Analytics Engine

`clsAnalyticsEngine` converts a collection of `clsArrivalResult` records into `clsExecutionStatistics`.

## Metrics

The statistics object exposes:

- `TotalPlanned`
- `TotalArrived`
- `TotalRemaining`
- `CompletionPercent`
- `DailyArrivals`
- `OpenDeliveries`
- `LateDeliveries`
- `SupplierStatistics`

## Calculation rules

- Totals are summed from arrival results.
- Completion percentage is `TotalArrived / TotalPlanned`, or zero when nothing is planned.
- Daily arrivals are grouped by `LastArrivalDate` using `yyyy-mm-dd` keys.
- Open deliveries include results with remaining quantity or open arrival statuses.
- Late deliveries are open deliveries where `RequiredDate` is before the report date.
- Supplier statistics are dictionary buckets keyed by supplier, with `UNKNOWN` used when a supplier is blank.

## Performance notes

Analytics reads the results collection once, maintains dictionary aggregates, and avoids worksheet access entirely.
