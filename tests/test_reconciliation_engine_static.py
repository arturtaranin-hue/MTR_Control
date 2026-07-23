from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def test_reconciliation_classes_exist():
    for path in [
        "src/clsReconciliationService.cls",
        "src/clsMaterialState.cls",
        "src/clsMaterialStatus.cls",
        "src/clsReportBuilder.cls",
    ]:
        assert (ROOT / path).exists(), path


def test_material_statuses_and_state_fields():
    status = read("src/clsMaterialStatus.cls")
    for expected in ["PLANNED", "PARTIAL", "DELIVERED", "OVERDELIVERED", "NOT_FOUND", "MANUAL_REVIEW"]:
        assert expected in status

    state = read("src/clsMaterialState.cls")
    for expected in [
        "PlannedQty", "ArrivedQty", "RemainingQty", "DeliveryPercent",
        "DuplicateArrival", "NegativeBalance", "OverDelivery", "UnknownMaterial", "ManualReview",
        "MovementQty", "AngaraQty",
    ]:
        assert expected in state


def test_reconciliation_service_combines_sources_and_detects_exceptions():
    source = read("src/clsReconciliationService.cls")
    for expected in [
        "Reconcile(ByVal dvPlanRows", "movementSummaryRows", "gpnAngaraRows",
        "BuildArrivalIndexes", "AddArrivalRows indexes, movementRows", "AddArrivalRows indexes, angaraRows",
        'sourceName = "GPN Angara"', 'bucket("quantity") = CDbl(bucket("quantity")) + qty',
        'indexes("records").Exists(recordKey)', "DuplicateArrival", "NegativeBalance", "OverDelivery",
        "AddUnknownMaterials", "UnknownMaterial = True", "MarkManualReview", "Set Reconcile = states",
    ]:
        assert expected in source
    assert ".Cells(" not in source


def test_report_builder_and_documentation():
    report = read("src/clsReportBuilder.cls")
    for expected in ["BuildSummary", "BuildRows", "TotalPlannedQty", "TotalArrivedQty", "DeliveryPercent"]:
        assert expected in report

    doc = read("docs/ReconciliationEngine.md")
    for expected in ["DV Plan", "Movement Summary", "GPN Angara", "Duplicate arrivals", "Unknown material"]:
        assert expected in doc
