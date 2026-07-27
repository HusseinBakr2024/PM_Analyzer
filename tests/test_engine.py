from datetime import date
from pathlib import Path

import pytest

from pm_analyzer.engine import (
    analyze,
    classify_material,
    classify_orders,
    normalize_asset,
    parse_date,
    parse_duration_hours,
)
from pm_analyzer.material_map import PMMaterialMap
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


def test_negative_idle_equivalent_is_rejected() -> None:
    with pytest.raises(ValueError, match="policy"):
        analyze(
            Path("maintenance.xlsx"),
            Path("materials.xlsx"),
            [Path("gps.xlsx")],
            interval_km=10_000,
            due_soon_percent=80,
            idle_hour_equivalent_km=-1,
        )


def test_material_map_covers_extended_pm_knowledge() -> None:
    material_map = PMMaterialMap.load()

    def category(description: str) -> str | None:
        match = material_map.match(description)
        return match[0] if match else None

    assert category("ENGINE LUBRICANT SAE 15W40") == "Engine Oil"
    assert category("LUBE FILTER ABC") == "Oil Filter"
    assert category("AIR CLEANER ELEMENT") == "Air Filter"
    assert category("ANTIFREEZE COOLANT") == "Coolant"
    assert material_map.match("GEAR OIL SAE 90") is None


def test_order_classification_scores_the_complete_material_set() -> None:
    material_map = PMMaterialMap.load()
    orders = {
        "1": {"asset_id": "HT-1"},
        "2": {"asset_id": "HT-2"},
        "3": {"asset_id": "HT-3"},
    }
    materials = [
        {"order_id": "1", "category": "Oil Filter", "pm_weight": 3, "matched_keyword": "OIL FILTER", "material": "OIL FILTER", "posting_date": date(2026, 1, 1)},
        {"order_id": "1", "category": "Fuel Filter", "pm_weight": 2, "matched_keyword": "FUEL FILTER", "material": "FUEL FILTER", "posting_date": date(2026, 1, 1)},
        {"order_id": "2", "category": "Air Filter", "pm_weight": 2, "matched_keyword": "AIR FILTER", "material": "AIR FILTER", "posting_date": date(2026, 1, 2)},
        {"order_id": "3", "category": "", "pm_weight": 0, "matched_keyword": "", "material": "WHEEL BEARING", "posting_date": date(2026, 1, 3)},
    ]

    result = {row["order_id"]: row for row in classify_orders(orders, materials, material_map)}

    assert result["1"]["classification"] == "PM"
    assert result["1"]["pm_score"] == 5
    assert result["2"]["classification"] == "Uncertain"
    assert result["3"]["classification"] == "Corrective"


def test_generated_report_can_be_read_back(tmp_path: Path) -> None:
    report = tmp_path / "report.xlsx"
    write_report(report, [("تحليل", ["المعدة", "الحالة"], [["HT-1", "Due"]])])

    workbook = read_workbook(report)

    assert workbook["تحليل"] == [{"المعدة": "HT-1", "الحالة": "Due"}]
