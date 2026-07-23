from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def test_reporting_classes_exist():
    for path in [
        "src/Services/clsReportService.cls",
        "src/Services/clsExcelExporter.cls",
    ]:
        assert (ROOT / path).exists(), path


def test_report_service_builds_grouped_final_report_and_dashboard():
    source = read("src/Services/clsReportService.cls")
    for expected in [
        "BuildFinalReport",
        "grouped.CompareMode = vbTextCompare",
        "FinalizeRow grouped(key)",
        'row("Material")',
        'row("PlannedQty")',
        'row("ArrivedQty")',
        'row("RemainingQty")',
        'row("DeliveryPercent")',
        'row("Status")',
        'row("LastArrivalDate")',
        'row("Confidence")',
        'row("MatchRule")',
        "BuildDashboard",
        'dashboard("TotalPlanned")',
        'dashboard("ManualReviewCount")',
    ]:
        assert expected in source


def test_excel_exporter_formats_report_and_dashboard():
    source = read("src/Services/clsExcelExporter.cls")
    for expected in [
        'REPORT_SHEET As String = "Report"',
        'DASHBOARD_SHEET As String = "Dashboard"',
        "ExportFinalReport",
        ".AutoFilter",
        ".Columns.AutoFit",
        "FreezePanes = True",
        "FormatConditions.Add",
        '"Completed", RGB(198, 239, 206)',
        '"Partial", RGB(255, 235, 156)',
        '"Missing", RGB(255, 199, 206)',
        "AddTotalsRow",
        '"Total planned"',
        '"Number of matched materials"',
        '"Number requiring manual review"',
        '" [Export] "',
    ]:
        assert expected in source


def test_reporting_documentation_exists():
    doc = read("docs/Reporting.md")
    assert "FinalReport" in doc
    assert "Dashboard" in doc
    assert "Conditional formatting" in doc
