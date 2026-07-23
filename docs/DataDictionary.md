# Data Dictionary

This document defines the initial domain vocabulary for MTR_Control. Field names are placeholders and may be refined as workbook templates are finalized.

## Material

| Field | Type | Description |
| --- | --- | --- |
| MaterialId | String | Unique material identifier. |
| MaterialCode | String | Business-facing material code. |
| Description | String | Human-readable material description. |
| UnitOfMeasure | String | Unit used for quantities. |

## Movement

| Field | Type | Description |
| --- | --- | --- |
| MovementId | String | Unique movement identifier. |
| MaterialId | String | Related material identifier. |
| MovementDate | Date | Date of inventory movement. |
| Quantity | Double | Movement quantity. |
| SourceLocation | String | Origin location. |
| DestinationLocation | String | Destination location. |

## Arrival

| Field | Type | Description |
| --- | --- | --- |
| ArrivalId | String | Unique arrival identifier. |
| MaterialId | String | Related material identifier. |
| ArrivalDate | Date | Date material arrived. |
| Quantity | Double | Arrival quantity. |
| ReferenceNumber | String | External reference number. |
