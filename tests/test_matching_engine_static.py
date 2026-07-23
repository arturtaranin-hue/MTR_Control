from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def test_required_classes_exist():
    for path in [
        "src/Services/clsNormalizeService.cls",
        "src/Services/clsMatchingService.cls",
        "src/Config/clsImportMapping.cls",
        "src/Config/clsMatchingRules.cls",
        "src/Models/clsMatchResult.cls",
    ]:
        assert (ROOT / path).exists(), path


def test_normalize_service_covers_required_steps():
    source = read("src/Services/clsNormalizeService.cls")
    assert "UCase$" in source
    assert "Trim$" in source
    assert "RemovePunctuation" in source
    assert "CollapseSpaces" in source
    assert "UnitDictionary" in source
    assert "AbbreviationDictionary" in source


def test_matching_service_priority_and_accumulation():
    source = read("src/Services/clsMatchingService.cls")
    sap = source.index('"SAP Code"')
    material = source.index('"Material Code"')
    normalized = source.index('"Normalized Name"')
    fuzzy = source.index('"Fuzzy Name"')
    assert sap < material < normalized < fuzzy
    assert 'indexes("sap")' in source
    assert 'indexes("material")' in source
    assert 'indexes("name")' in source
    assert 'bucket("quantity") = CDbl(bucket("quantity")) +' in source
    assert "For rowIndex = LBound(arrivalRows, 1) To UBound(arrivalRows, 1)" in source


def test_match_result_statuses_and_fields():
    source = read("src/Models/clsMatchResult.cls")
    for expected in [
        "MatchResultMatched",
        "MatchResultPartiallyMatched",
        "MatchResultNotMatched",
        "MatchResultNeedManualReview",
        "Material",
        "PlannedQty",
        "ArrivedQty",
        "RemainingQty",
        "ArrivalDate",
        "Confidence",
    ]:
        assert expected in source


def test_documentation_exists():
    doc = read("docs/MatchingEngine.md")
    assert "Matching priority" in doc
    assert "Quantities and partial deliveries" in doc
    assert "Performance notes" in doc
