from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def test_mvp_module_exists_and_runs_complete_pipeline():
    source = read("src/modMvp.bas")
    for expected in [
        "Public Sub Run()",
        "PickSourceFile(\"Select DV Angara file\")",
        "PickSourceFile(\"Select Movement Summary file\")",
        "PickSourceFile(\"Select GPN Angara file\")",
        "LoadDV(dvPath)",
        "LoadMovement(movementPath)",
        "LoadAngara(angaraPath)",
        "NormalizeRows(dvRows)",
        "MatchMaterials(normalizedDvRows, normalizedMovementRows)",
        "CalculateArrivals(normalizedDvRows, normalizedMovementRows)",
        "Reconcile(normalizedDvRows, normalizedMovementRows, normalizedAngaraRows)",
        "BuildFinalReport(states)",
        "ExportFinalReport finalReport, targetWorkbook",
        "SaveReport(targetWorkbook)",
    ]:
        assert expected in source


def test_mvp_dashboard_buttons_and_error_handling():
    source = read("src/modMvp.bas")
    for expected in [
        'APP_WORKBOOK_NAME As String = "MTR_Control.xlsm"',
        'DASHBOARD_SHEET As String = "Dashboard"',
        'EnsureButton dashboard, "Import Files", "ImportFiles"',
        'EnsureButton dashboard, "Generate Report", "GenerateReport"',
        'EnsureButton dashboard, "Open Log", "OpenLog"',
        "Application.FileDialog(msoFileDialogFilePicker)",
        ".Filters.Add \"Excel files\"",
        "On Error GoTo Fail",
        "LogMvpStep \"Error\"",
        "RestoreApplicationState",
    ]:
        assert expected in source


def test_binary_workbook_is_generated_not_tracked():
    assert not (ROOT / "MTR_Control.xlsm").exists()
    assert (ROOT / "BUILD.md").exists()
    assert (ROOT / "scripts" / "Build-MTRControl.ps1").exists()
    assert (ROOT / "releases" / ".gitkeep").exists()

    build_doc = read("BUILD.md")
    assert "generated release artifact" in build_doc
    assert "scripts\\Build-MTRControl.ps1" in build_doc
    assert "Trust access to the VBA project object model" in build_doc

    script = read("scripts/Build-MTRControl.ps1")
    assert "VBComponents.Import" in script
    assert "SetupDashboard" in script
    assert "MTR_Control.xlsm" in script
