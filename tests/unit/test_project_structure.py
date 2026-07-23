from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


def test_layered_project_structure_exists() -> None:
    expected_paths = [
        ROOT / "src" / "mtr_control" / "domain",
        ROOT / "src" / "mtr_control" / "application",
        ROOT / "src" / "mtr_control" / "infrastructure",
        ROOT / "src" / "mtr_control" / "interfaces",
        ROOT / "src" / "mtr_control" / "shared",
        ROOT / "docs" / "architecture.md",
    ]

    missing_paths = [path for path in expected_paths if not path.exists()]

    assert missing_paths == []
