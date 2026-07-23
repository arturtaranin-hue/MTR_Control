from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def test_arrival_and_analytics_classes_exist():
    for path in [
        "src/Services/clsArrivalEngine.cls",
        "src/Services/clsAnalyticsEngine.cls",
        "src/Models/clsArrivalResult.cls",
        "src/Models/clsExecutionStatistics.cls",
    ]:
        assert (ROOT / path).exists(), path


def test_arrival_result_statuses_and_fields():
    source = read("src/Models/clsArrivalResult.cls")
    for expected in [
        "ARRIVED",
        "PARTIALLY_ARRIVED",
        "NOT_ARRIVED",
        "OVER_DELIVERED",
        "MANUAL_REVIEW",
        "PlannedQty",
        "ArrivedQty",
        "RemainingQty",
        "CompletionPercent",
        "FirstArrivalDate",
        "LastArrivalDate",
        "DeliveryCount",
    ]:
        assert expected in source


def test_arrival_engine_business_rules_and_performance():
    source = read("src/Services/clsArrivalEngine.cls")
    assert "BuildMovementIndexes" in source
    assert 'indexes("sap")' in source
    assert 'indexes("material")' in source
    assert 'indexes("name")' in source
    assert "seenRecords.Exists(recordKey)" in source
    assert 'bucket("quantity") = CDbl(bucket("quantity")) + qty' in source
    assert 'bucket("deliveryCount") = CLng(bucket("deliveryCount")) + 1' in source
    assert "EarliestDate" in source
    assert "LatestDate" in source
    assert "OVER_DELIVERED" in source
    assert "PARTIALLY_ARRIVED" in source
    assert "UnexpectedArrival = True" in source
    assert "For rowIndex = LBound(movementRows, 1) To UBound(movementRows, 1)" in source
    assert ".Cells(" not in source


def test_analytics_engine_metrics():
    source = read("src/Services/clsAnalyticsEngine.cls")
    for expected in [
        "TotalPlanned",
        "TotalArrived",
        "TotalRemaining",
        "CompletionPercent",
        "DailyArrivals",
        "OpenDeliveries",
        "LateDeliveries",
        "SupplierStatistics",
        "For Each result In arrivalResults",
    ]:
        assert expected in source
    assert ".Cells(" not in source


def test_arrival_and_analytics_documentation_exists():
    arrival_doc = read("docs/ArrivalEngine.md")
    analytics_doc = read("docs/AnalyticsEngine.md")
    assert "Business rules" in arrival_doc
    assert "Performance notes" in arrival_doc
    assert "Daily arrivals" in analytics_doc
    assert "Supplier statistics" in analytics_doc
