# Reporting and Excel export

`clsReportService` turns reconciliation or matching results into a `FinalReport` collection grouped by material. Each report row is a dictionary containing material, planned quantity, arrived quantity, remaining quantity, delivery percentage, status, last arrival date, confidence, match rule, and manual-review metadata.

## Report statuses

- **Completed**: arrived quantity is greater than or equal to planned quantity.
- **Partial**: arrived quantity is greater than zero but less than planned quantity.
- **Missing**: no arrivals are associated with the material.
- **Manual Review**: upstream data flags a review reason or the result indicates a manual-review status.

## Dashboard metrics

`BuildDashboard` summarizes the final report into total planned, total arrived, total remaining, completion percentage, number of matched materials, and number requiring manual review. Low-confidence rows below 80% are counted as requiring review.

## Excel export

`clsExcelExporter.ExportFinalReport` writes two worksheets to the target workbook:

1. `Report` contains the detailed final report, AutoFilter, frozen first row, AutoFit columns, a totals row, percentage formatting, and conditional status colors.
2. `Dashboard` contains high-level totals and completion metrics.

Conditional formatting uses green for `Completed`, yellow for `Partial`, and red for `Missing`. The exporter logs start and completion messages with an `[Export]` prefix in the Immediate window.
