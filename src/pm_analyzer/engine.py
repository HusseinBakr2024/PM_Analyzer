"""Preventive-maintenance analysis engine."""

from __future__ import annotations

import math
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from pm_analyzer.config import DUE_SOON_PERCENT, IDLE_HOUR_EQUIVALENT_KM, PM_INTERVAL_KM
from pm_analyzer.xlsx import read_workbook, write_report

_TRUCK_PATTERN = re.compile(r"HT-\d{2}-(?:M|C|STH)-\d+", re.IGNORECASE)
_ORDER_PATTERN = re.compile(r"^(.*?)\s*\((\d+)\)\s*$")
_DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_MISSING = {"", "-----", "N/A", "NONE", "NULL"}
_CATEGORIES = (
    ("Engine Oil", ("ENGINE OIL", "MOTOR OIL")),
    ("Oil Filter", ("OIL FILTER", "FILTER OIL")),
    ("Fuel Filter", ("FUEL FILTER", "DIESEL FILTER")),
    ("Air Filter", ("AIR FILTER",)),
    ("Water Separator", ("WATER SEPARATOR", "FUEL SEPARATOR", "FILTER-SEPARATOR")),
)


@dataclass(slots=True)
class AnalysisResult:
    """Report-ready analysis result and audit detail."""

    analysis: list[dict[str, Any]] = field(default_factory=list)
    maintenance: list[dict[str, Any]] = field(default_factory=list)
    gps: list[dict[str, Any]] = field(default_factory=list)
    materials: list[dict[str, Any]] = field(default_factory=list)
    quality: list[dict[str, Any]] = field(default_factory=list)


def analyze(
    maintenance_path: Path,
    materials_path: Path,
    gps_paths: list[Path],
    *,
    interval_km: int = PM_INTERVAL_KM,
    due_soon_percent: int = DUE_SOON_PERCENT,
) -> AnalysisResult:
    """Analyze preventive maintenance using SAP as the authoritative asset population."""
    if not 1 <= len(gps_paths) <= 7:
        raise ValueError("Select between 1 and 7 GPS files")
    if interval_km <= 0 or not 0 <= due_soon_percent <= 100:
        raise ValueError("Maintenance policy values are invalid")

    result = AnalysisResult()
    orders = _read_maintenance(maintenance_path, result)
    material_rows = _read_materials(materials_path, orders, result)
    gps_rows = _read_gps(gps_paths, result)
    result.maintenance = list(orders.values())
    result.materials = material_rows
    result.gps = gps_rows
    result.analysis = _calculate(orders, material_rows, gps_rows, result, interval_km, due_soon_percent)
    return result


def _first_sheet(path: Path, preferred: str) -> list[dict[str, str]]:
    workbook = read_workbook(path)
    if preferred in workbook:
        return workbook[preferred]
    if not workbook:
        raise ValueError(f"No worksheets found in {path.name}")
    return next(iter(workbook.values()))


def _read_maintenance(path: Path, result: AnalysisResult) -> dict[str, dict[str, Any]]:
    rows = _first_sheet(path, "SAPUI5 Export")
    required = {"Execution Object", "Main Work Center", "Order Type"}
    if rows and not required.issubset(rows[0]):
        raise ValueError(f"Maintenance columns not recognized in {path.name}")
    orders: dict[str, dict[str, Any]] = {}
    for index, row in enumerate(rows, 2):
        match = _ORDER_PATTERN.match(row.get("Execution Object", ""))
        if not match:
            result.quality.append(_issue("أمر بدون معدة أو تنسيق غير صالح", path.name, index, row.get("Execution Object", "")))
            continue
        asset = normalize_asset(match.group(1))
        order = match.group(2)
        if not asset:
            result.quality.append(_issue("أمر بدون معدة", path.name, index, order))
            continue
        order_type = row.get("Order Type", "").strip()
        if "YA02" not in order_type.upper() and "PROACTIVE" not in order_type.upper():
            continue
        orders[order] = {
            "asset_id": asset,
            "order_id": order,
            "work_center": row.get("Main Work Center", "").strip(),
            "order_type": order_type,
            "created_at": parse_date(row.get("Created On/At (UTC)", "")),
            "reference_at": parse_date(row.get("Reference Date/Time (UTC)", "")),
            "actual_cost": parse_number(row.get("Actual Cost", "")),
            "subphase": row.get("Subphase", "").strip(),
        }
    return orders


def _read_materials(
    path: Path, orders: dict[str, dict[str, Any]], result: AnalysisResult
) -> list[dict[str, Any]]:
    rows = _first_sheet(path, "SAPUI5 Export")
    required = {"Order", "Material", "Posting Date"}
    if rows and not required.issubset(rows[0]):
        raise ValueError(f"Material columns not recognized in {path.name}")
    output = []
    for index, row in enumerate(rows, 2):
        order = row.get("Order", "").strip()
        if not order:
            result.quality.append(_issue("مادة بدون رقم أمر", path.name, index, row.get("Material", "")))
            continue
        if order not in orders:
            continue
        description = clean_text(row.get("Material", ""))
        category = classify_material(description)
        if category is None:
            continue
        posting_date = parse_date(row.get("Posting Date", ""))
        if posting_date is None:
            result.quality.append(_issue("تاريخ صرف غير صالح", path.name, index, order))
        output.append(
            {
                "asset_id": orders[order]["asset_id"],
                "order_id": order,
                "category": category,
                "material_group": row.get("Material Group", "").strip(),
                "material": description,
                "quantity": parse_number(row.get("Quantity", "")),
                "value": parse_number(row.get("Value in Local Currency", "")),
                "posting_date": posting_date,
                "storage_location": row.get("Storage Location", "").strip(),
                "valuation_type": row.get("Valuation Type", "").strip(),
            }
        )
    return output


def _read_gps(paths: list[Path], result: AnalysisResult) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    seen: set[tuple[object, ...]] = set()
    for path in paths:
        workbook = read_workbook(path)
        sheet_name = "Trucks Summary-MAN" if "Trucks Summary-MAN" in workbook else "Trucks Summary"
        rows = workbook.get(sheet_name)
        if not rows:
            result.quality.append(_issue("ملف GPS لم يتم التعرف على شيت الملخص", path.name, 0, ""))
            continue
        header = rows[0]
        if "Grouping" not in header:
            result.quality.append(_issue("ملف GPS لم يتم التعرف على أعمدته", path.name, 1, "Grouping"))
            continue
        distance_column = _find_column(header, "Distance Travelled")
        engine_column = _find_column(header, "Engine Run hours")
        idle_column = _find_column(header, "Idling Hours")
        if not all((distance_column, engine_column, idle_column)):
            result.quality.append(_issue("ملف GPS لم يتم التعرف على أعمدته", path.name, 1, "distance/engine/idle"))
            continue
        assert distance_column is not None
        assert engine_column is not None
        assert idle_column is not None
        current_date: date | None = None
        brand = _brand(path.name)
        for index, row in enumerate(rows, 2):
            grouping = clean_text(row.get("Grouping", ""))
            if _DATE_PATTERN.fullmatch(grouping):
                current_date = date.fromisoformat(grouping)
                continue
            if not _TRUCK_PATTERN.fullmatch(normalize_asset(grouping)):
                continue
            if current_date is None:
                result.quality.append(_issue("صف GPS بلا تاريخ سابق", path.name, index, grouping))
                continue
            asset = normalize_asset(grouping)
            record: dict[str, Any] = {
                "asset_id": asset,
                "gps_date": current_date,
                "distance_km": parse_number(row.get(distance_column, "")),
                "engine_hours": parse_duration_hours(row.get(engine_column, "")),
                "idle_hours": parse_duration_hours(row.get(idle_column, "")),
                "brand": brand,
                "source_file": path.name,
            }
            key = (asset, current_date, record["distance_km"], record["engine_hours"], record["idle_hours"])
            if key in seen:
                result.quality.append(_issue("سجل GPS مكرر", path.name, index, f"{asset} {current_date}"))
                continue
            seen.add(key)
            if record["distance_km"] is not None and record["distance_km"] < 0:
                result.quality.append(_issue("مسافة GPS سالبة", path.name, index, f"{asset} {current_date}"))
                record["distance_km"] = None
            output.append(record)
    return output


def _calculate(
    orders: dict[str, dict[str, Any]],
    materials: list[dict[str, Any]],
    gps: list[dict[str, Any]],
    result: AnalysisResult,
    interval: int,
    threshold: int,
) -> list[dict[str, Any]]:
    engine_oil_by_asset: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for material in materials:
        if material["category"] == "Engine Oil" and material["posting_date"] is not None:
            engine_oil_by_asset[str(material["asset_id"])].append(material)
    orders_by_asset: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for order in orders.values():
        orders_by_asset[str(order["asset_id"])].append(order)
    gps_by_asset: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in gps:
        gps_by_asset[str(row["asset_id"])].append(row)

    output = []
    for asset in sorted(orders_by_asset):
        oil_rows = engine_oil_by_asset.get(asset, [])
        latest_oil = max(oil_rows, key=lambda item: item["posting_date"]) if oil_rows else None
        if latest_oil is None:
            latest_order = max(orders_by_asset[asset], key=lambda item: item["reference_at"] or date.min)
            output.append(_analysis_row(asset, latest_order, None, [], interval, threshold, "No Previous PM"))
            continue
        latest_order = orders[str(latest_oil["order_id"])]
        pm_date = latest_oil["posting_date"]
        available = sorted((row for row in gps_by_asset.get(asset, []) if row["gps_date"] >= pm_date), key=lambda row: row["gps_date"])
        status = None if available else "No GPS Data"
        record = _analysis_row(asset, latest_order, pm_date, available, interval, threshold, status)
        all_asset_gps = gps_by_asset.get(asset, [])
        if available and all_asset_gps and min(row["gps_date"] for row in all_asset_gps) > pm_date:
            record["note"] = "تغطية GPS تبدأ بعد تاريخ آخر صيانة"
        output.append(record)

    sap_assets = set(orders_by_asset)
    gps_assets = set(gps_by_asset)
    for asset in sorted(sap_assets - gps_assets):
        result.quality.append(_issue("معدة SAP غير موجودة في GPS", "SAP/GPS", 0, asset))
    for asset in sorted(gps_assets - sap_assets):
        result.quality.append(_issue("معدة GPS غير موجودة في SAP", "SAP/GPS", 0, asset))
    return output


def _analysis_row(
    asset: str,
    order: dict[str, Any],
    pm_date: date | None,
    gps: list[dict[str, Any]],
    interval: int,
    threshold: int,
    forced_status: str | None,
) -> dict[str, Any]:
    distance = sum(float(row["distance_km"] or 0) for row in gps)
    engine = sum(float(row["engine_hours"] or 0) for row in gps)
    idle = sum(float(row["idle_hours"] or 0) for row in gps)
    idle_km = idle * IDLE_HOUR_EQUIVALENT_KM
    equivalent = distance + idle_km
    usage = equivalent / interval * 100
    if forced_status:
        status = forced_status
    elif usage >= 100:
        status = "Due"
    elif usage >= threshold:
        status = "Due Soon"
    else:
        status = "OK"
    return {
        "asset_id": asset,
        "order_id": order["order_id"],
        "pm_date": pm_date,
        "work_center": order["work_center"],
        "brand": gps[0]["brand"] if gps else _brand(asset),
        "gps_first": gps[0]["gps_date"] if gps else None,
        "gps_last": gps[-1]["gps_date"] if gps else None,
        "gps_days": len({row["gps_date"] for row in gps}),
        "distance_km": round(distance, 2),
        "engine_hours": round(engine, 2),
        "idle_hours": round(idle, 2),
        "idle_equivalent_km": round(idle_km, 2),
        "equivalent_km": round(equivalent, 2),
        "interval_km": interval,
        "remaining_km": round(interval - equivalent, 2),
        "usage_percent": round(usage, 2),
        "status": status,
        "note": "لا توجد بيانات GPS بعد الصيانة" if status == "No GPS Data" else "لم يتم العثور على صرف زيت محرك" if status == "No Previous PM" else "",
    }


def export_report(result: AnalysisResult, path: Path) -> None:
    """Export the ten-sheet Arabic preventive-maintenance report."""
    columns = [
        ("asset_id", "كود المعدة"), ("order_id", "رقم آخر أمر صيانة"),
        ("pm_date", "تاريخ آخر صيانة"), ("work_center", "مركز العمل"), ("brand", "الشركة"),
        ("gps_first", "أول تاريخ GPS محسوب"), ("gps_last", "آخر تاريخ GPS محسوب"),
        ("gps_days", "عدد أيام GPS"), ("distance_km", "إجمالي الكيلومترات"),
        ("engine_hours", "إجمالي ساعات التشغيل"), ("idle_hours", "ساعات التشغيل الساكن"),
        ("idle_equivalent_km", "الكيلومترات المكافئة للساكن"),
        ("equivalent_km", "إجمالي الكيلومترات المكافئة"), ("interval_km", "فترة الصيانة"),
        ("remaining_km", "المتبقي حتى الصيانة"), ("usage_percent", "نسبة استهلاك الدورة %"),
        ("status", "حالة الصيانة"), ("note", "ملاحظة"),
    ]
    headers = [label for _, label in columns]
    table = [[row.get(key) for key, _ in columns] for row in result.analysis]
    counts: dict[str, int] = defaultdict(int)
    for row in result.analysis:
        counts[str(row["status"])] += 1
    dashboard = [
        ("إجمالي المعدات", len(result.analysis)), ("معدات OK", counts["OK"]),
        ("معدات Due Soon", counts["Due Soon"]), ("معدات Due", counts["Due"]),
        ("معدات بدون GPS", counts["No GPS Data"]), ("معدات بدون صيانة سابقة", counts["No Previous PM"]),
        ("إجمالي الكيلومترات", round(sum(float(row["distance_km"]) for row in result.analysis), 2)),
        ("إجمالي ساعات Idle", round(sum(float(row["idle_hours"]) for row in result.analysis), 2)),
        ("إجمالي الكيلومترات المكافئة", round(sum(float(row["equivalent_km"]) for row in result.analysis), 2)),
    ]
    def filtered(status: str) -> list[list[object]]:
        return [
            row
            for row, source in zip(table, result.analysis, strict=True)
            if source["status"] == status
        ]
    maintenance_headers = ["كود المعدة", "رقم الأمر", "مركز العمل", "نوع الأمر", "تاريخ الإنشاء", "التاريخ المرجعي", "التكلفة", "المرحلة"]
    maintenance_by_order = {str(row["order_id"]): row for row in result.maintenance}
    latest_maintenance = [
        maintenance_by_order[str(row["order_id"])] for row in result.analysis
    ]
    maintenance_rows = [[r.get(key) for key in ("asset_id", "order_id", "work_center", "order_type", "created_at", "reference_at", "actual_cost", "subphase")] for r in latest_maintenance]
    gps_headers = ["كود المعدة", "التاريخ", "المسافة كم", "ساعات التشغيل", "ساعات الساكن", "الشركة", "الملف"]
    gps_rows = [[r.get(key) for key in ("asset_id", "gps_date", "distance_km", "engine_hours", "idle_hours", "brand", "source_file")] for r in result.gps]
    material_headers = ["كود المعدة", "رقم الأمر", "التصنيف", "مجموعة المادة", "المادة", "الكمية", "القيمة", "تاريخ الصرف", "المخزن", "نوع التقييم"]
    material_rows = [[r.get(key) for key in ("asset_id", "order_id", "category", "material_group", "material", "quantity", "value", "posting_date", "storage_location", "valuation_type")] for r in result.materials]
    quality_rows = [[r["issue"], r["source"], r["row"], r["value"]] for r in result.quality]
    write_report(path, [
        ("لوحة التحكم", ["المؤشر", "القيمة"], dashboard),
        ("تحليل استحقاق الصيانة", headers, table),
        ("المعدات المستحقة", headers, filtered("Due")),
        ("صيانة قريبة", headers, filtered("Due Soon")),
        ("بدون بيانات GPS", headers, filtered("No GPS Data")),
        ("بدون صيانة سابقة", headers, filtered("No Previous PM")),
        ("آخر صيانة لكل معدة", maintenance_headers, maintenance_rows),
        ("تفاصيل GPS المدمجة", gps_headers, gps_rows),
        ("تفاصيل مواد الصيانة الوقائية", material_headers, material_rows),
        ("جودة البيانات", ["المشكلة", "المصدر", "الصف", "القيمة"], quality_rows),
    ])


def normalize_asset(value: str) -> str:
    """Normalize asset codes without guessing missing identifiers."""
    text = clean_text(value).upper()
    text = "".join("-" if unicodedata.category(character) == "Pd" else character for character in text)
    return re.sub(r"\s+", "", text)


def clean_text(value: str) -> str:
    return " ".join(unicodedata.normalize("NFKC", str(value)).replace("\u200b", "").split())


def classify_material(description: str) -> str | None:
    upper = clean_text(description).upper().replace("–", "-").replace("‑", "-")
    if "OIL SEAL" in upper or "COIL" in upper or "HYDRAULIC OIL" in upper or "GEAR OIL" in upper or "ATF OIL" in upper:
        return None
    for category, keywords in _CATEGORIES:
        if any(keyword in upper for keyword in keywords):
            return category
    return None


def parse_number(value: object) -> float | None:
    text = clean_text(str(value)).upper().replace(",", "").replace(" KMPL", "").replace(" KM", "")
    if text in _MISSING:
        return None
    try:
        number = float(text)
        return number if math.isfinite(number) else None
    except ValueError:
        return None


def parse_duration_hours(value: object) -> float | None:
    text = clean_text(str(value))
    if text.upper() in _MISSING:
        return None
    match = re.fullmatch(r"(?:(\d+) days? )?(\d+):(\d{2}):(\d{2})", text)
    if not match:
        return parse_number(text)
    days, hours, minutes, seconds = match.groups()
    return (int(days or 0) * 24) + int(hours) + int(minutes) / 60 + int(seconds) / 3600


def parse_date(value: object) -> date | None:
    text = clean_text(str(value))
    if text.upper() in _MISSING:
        return None
    try:
        if re.fullmatch(r"\d+(?:\.\d+)?", text):
            return (datetime(1899, 12, 30) + timedelta(days=float(text))).date()
        return date.fromisoformat(text[:10])
    except (ValueError, OverflowError):
        return None


def _find_column(row: dict[str, str], fragment: str) -> str | None:
    return next((name for name in row if fragment.casefold() in name.casefold()), None)


def _brand(value: str) -> str:
    upper = value.upper()
    if "MAN" in upper or "-M-" in upper:
        return "MAN"
    if "MERC" in upper or "-C-" in upper:
        return "Mercedes"
    if "SINO" in upper or "-STH-" in upper:
        return "SINO"
    return "غير محدد"


def _issue(issue: str, source: str, row: int, value: str) -> dict[str, Any]:
    return {"issue": issue, "source": source, "row": row, "value": value}
