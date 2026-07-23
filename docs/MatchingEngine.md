# Matching Engine

Sprint 2 adds a normalize-and-match pipeline for reconciling planned materials from **DV Angara** with actual arrivals from **Movement Summary** and **GPN Angara**.

## Components

- `clsImportMapping` configures source array column positions for SAP code, material code, material name, quantity, and arrival date.
- `clsMatchingRules` configures fuzzy threshold, measurement-unit aliases, and abbreviation expansion dictionaries.
- `clsNormalizeService` prepares material names for stable comparison.
- `clsMatchingService` builds dictionary indexes over arrival arrays and matches each planned row by priority.
- `clsMatchResult` stores the matching outcome for a planned material.

## Normalization

`clsNormalizeService.NormalizeText` applies the following steps:

1. Convert to string, trim leading/trailing spaces, and uppercase.
2. Remove punctuation by replacing punctuation characters with spaces.
3. Collapse repeated whitespace into a single space.
4. Normalize measurement units from `clsMatchingRules.UnitDictionary`.
5. Expand known abbreviations from `clsMatchingRules.AbbreviationDictionary`.

The dictionaries are configurable via `AddUnitAlias` and `AddAbbreviation`.

## Matching priority

`clsMatchingService.MatchMaterials(plannedRows, arrivalRows)` expects worksheet data that has already been read into two-dimensional arrays. It never scans worksheets directly.

Arrivals are indexed once in dictionaries, then each planned row is matched in this order:

1. **SAP Code** — exact lookup, confidence `1.00`.
2. **Material Code** — exact lookup, confidence `1.00`.
3. **Normalized Name** — exact lookup on normalized names, confidence `0.98`.
4. **Fuzzy Name** — Levenshtein similarity over normalized names, accepted when score is at least the configured threshold (`0.95` by default).

## Quantities and partial deliveries

Multiple arrival rows with the same key are accumulated into one bucket. The result stores:

- `Material`
- `PlannedQty`
- `ArrivedQty`
- `RemainingQty`
- `ArrivalDate` (latest arrival date for the bucket)
- `Confidence`
- `MatchKey`
- `MatchRule`

Result status is set as follows:

- `MatchResultMatched`: arrived quantity is greater than or equal to planned quantity.
- `MatchResultPartiallyMatched`: arrived quantity is positive but below planned quantity.
- `MatchResultNotMatched`: no priority produced an accepted match.
- `MatchResultNeedManualReview`: reserved for downstream workflows that require user adjudication.

## Performance notes

- Import worksheet ranges into arrays before calling the service.
- Arrival data is indexed in `Scripting.Dictionary` instances for SAP code, material code, and normalized name.
- Matching does not use nested worksheet loops.
- Fuzzy matching iterates over unique normalized arrival names only after exact keys fail.
