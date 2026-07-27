from datetime import date
from pathlib import Path

from pm_analyzer.engine import classify_material, normalize_asset, parse_date, parse_duration_hours
from pm_analyzer.xlsx import read_workbook, write_report


def test_normalize_asset_unifies_spacing_and_dashes() -> None:
    assert normalize_asset("  ht‑24-sth-1274 ") == "HT-24-STH-1274"


def test_preventive_material_classification_avoids_false_oil_matches() -> None:
    assert classify_material("ENGINE OIL 15W-40 CK-4") == "Engine Oil"
    assert classify_material("CAT – OIL FILTER 1R-1808") == "Oil Filter"
    assert classify_material("FG - FUEL/WATER SEPARATOR FS19532") == "Water Separator"
    assert classify_material("CORTECO-OIL SEAL 19030143") is None
    assert classify_material("ATF OIL") is None


def test_source_formats_are_parsed() -> None:
    assert parse_duration_hours("1:30:00") == 1.5
    assert parse_duration_hours("2 days 1:00:00") == 49
    assert parse_date("46174") == date(2026, 6, 1)


def test_generated_report_can_be_read_back(tmp_path: Path) -> None:
    report = tmp_path / "report.xlsx"
    write_report(report, [("تحليل", ["المعدة", "الحالة"], [["HT-1", "Due"]])])

    workbook = read_workbook(report)

    assert workbook["تحليل"] == [{"المعدة": "HT-1", "الحالة": "Due"}]
