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

from pm_analyzer.material_map import PMMaterialMap
from pm_analyzer.xlsx import read_workbook, write_report

_TRUCK_PATTERN = re.compile(r"HT-\d{2}-(?:M|C|STH)-\d+", re.IGNORECASE)
_ORDER_PATTERN = re.compile(r"^(.*?)\s*\((\d+)\)\s*$")
_DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_MISSING = {"", "-----", "N/A", "NONE", "NULL"}


@dataclass(slots=True)
class AnalysisResult:
    """Report-ready analysis result and audit detail."""

    analysis: list[dict[str, Any]] = field(default_factory=list)
    maintenance: list[dict[str, Any]] = field(default_factory=list)
    gps: list[dict[str, Any]] = field(default_factory=list)
    materials: list[dict[str, Any]] = field(default_factory=list)
    quality: list[dict[str, Any]] = field(default_factory=list)
    order_classifications: list[dict[str, Any]] = field(default_factory=list)


def analyze(
    maintenance_path: Path,
    materials_path: Path,
    gps_paths: list[Path],
    *,
    interval_km: int,
    due_soon_percent: int,
    idle_hour_equivalent_km: float,
    material_map: PMMaterialMap | None = None,
) -> AnalysisResult:
    """Analyze preventive maintenance using SAP as the authoritative asset population."""
    if not 1 <= len(gps_paths) <= 7:
        raise ValueError("Select between 1 and 7 GPS files")
    if interval_km <= 0 or not 0 <= due_soon_percent <= 100 or idle_hour_equivalent_km < 0:
        raise ValueError("Maintenance policy values are invalid")

    result = AnalysisResult()
    knowledge = material_map or PMMaterialMap.load()
    orders = _read_maintenance(maintenance_path, result)
    material_rows = _read_materials(materials_path, orders, result, knowledge)
    classifications = classify_orders(orders, material_rows, knowledge)
    gps_rows = _read_gps(gps_paths, result)
    result.maintenance = list(orders.values())
    result.materials = material_rows
    result.gps = gps_rows
    result.order_classifications = classifications
    result.analysis = _calculate(
        orders,
        gps_rows,
        result,
        interval_km,
        due_soon_percent,
        idle_hour_equivalent_km,
        classifications,
    )
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
    required = {"Execution Object", "Main Work Center"}
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
        orders[order] = {
            "asset_id": asset,
            "order_id": order,
            "work_center": row.get("Main Work Center", "").strip(),
            "order_type": row.get("Order Type", "").strip(),
            "created_at": parse_date(row.get("Created On/At (UTC)", "")),
            "reference_at": parse_date(row.get("Reference Date/Time (UTC)", "")),
            "actual_cost": parse_number(row.get("Actual Cost", "")),
            "subphase": row.get("Subphase", "").strip(),
        }
    return orders


def _read_materials(
    path: Path,
    orders: dict[str, dict[str, Any]],
    result: AnalysisResult,
    material_map: PMMaterialMap,
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
        match = material_map.match(description)
        posting_date = parse_date(row.get("Posting Date", ""))
        if posting_date is None:
            result.quality.append(_issue("تاريخ صرف غير صالح", path.name, index, order))
        output.append(
            {
                "asset_id": orders[order]["asset_id"],
                "order_id": order,
                "category": match[0] if match else "",
                "pm_weight": match[1] if match else 0,
                "matched_keyword": match[2] if match else "",
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


def classify_orders(
    orders: dict[str, dict[str, Any]],
    materials: list[dict[str, Any]],
    material_map: PMMaterialMap,
) -> list[dict[str, Any]]:
    """Classify each order once, based exclusively on its complete material set."""
    materials_by_order: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for material in materials:
        materials_by_order[str(material["order_id"])].append(material)
    classifications = []
    for order_id, order in orders.items():
        rows = materials_by_order.get(order_id, [])
        category_weights: dict[str, int] = {}
        evidence: list[str] = []
        for row in rows:
            category = str(row["category"])
            if not category:
                continue
            category_weights[category] = max(category_weights.get(category, 0), int(row["pm_weight"]))
            evidence.append(f"{category}: {row['matched_keyword']}")
        score = sum(category_weights.values())
        descriptions = [str(row["material"]) for row in rows]
        if score >= material_map.pm_score_threshold:
            classification = "PM"
        elif score > 0:
            classification = "Uncertain"
        elif material_map.has_breakdown_indicator(descriptions):
            classification = "Breakdown"
        elif rows:
            classification = "Corrective"
        else:
            classification = "Unclassified"
        posting_dates = [
            row["posting_date"]
            for row in rows
            if row["category"] and row["posting_date"] is not None
        ]
        classifications.append(
            {
                "asset_id": order["asset_id"],
                "order_id": order_id,
                "classification": classification,
                "pm_score": score,
                "pm_threshold": material_map.pm_score_threshold,
                "evidence": " | ".join(sorted(set(evidence))),
                "material_count": len(rows),
                "posting_date": max(posting_dates) if posting_dates else None,
                "map_version": material_map.version,
            }
        )
    return classifications


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
    gps: list[dict[str, Any]],
    result: AnalysisResult,
    interval: int,
    threshold: int,
    idle_equivalent: float,
    classifications: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    pm_by_asset: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for classification in classifications:
        if classification["classification"] == "PM" and classification["posting_date"] is not None:
            pm_by_asset[str(classification["asset_id"])].append(classification)
    orders_by_asset: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for order in orders.values():
        orders_by_asset[str(order["asset_id"])].append(order)
    gps_by_asset: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in gps:
        gps_by_asset[str(row["asset_id"])].append(row)

    output = []
    for asset in sorted(orders_by_asset):
        pm_orders = pm_by_asset.get(asset, [])
        latest_pm = max(pm_orders, key=lambda item: item["posting_date"]) if pm_orders else None
        if latest_pm is None:
            latest_order = max(orders_by_asset[asset], key=lambda item: item["reference_at"] or date.min)
            output.append(
                _analysis_row(
                    asset,
                    latest_order,
                    None,
                    [],
                    interval,
                    threshold,
                    idle_equivalent,
                    "No Previous PM",
                )
            )
            continue
        latest_order = orders[str(latest_pm["order_id"])]
        latest_order["pm_score"] = latest_pm["pm_score"]
        latest_order["pm_evidence"] = latest_pm["evidence"]
        pm_date = latest_pm["posting_date"]
        available = sorted((row for row in gps_by_asset.get(asset, []) if row["gps_date"] >= pm_date), key=lambda row: row["gps_date"])
        status = None if available else "No GPS Data"
        record = _analysis_row(
            asset,
            latest_order,
            pm_date,
            available,
            interval,
            threshold,
            idle_equivalent,
            status,
        )
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
    idle_equivalent: float,
    forced_status: str | None,
) -> dict[str, Any]:
    distance = sum(float(row["distance_km"] or 0) for row in gps)
    engine = sum(float(row["engine_hours"] or 0) for row in gps)
    idle = sum(float(row["idle_hours"] or 0) for row in gps)
    idle_km = idle * idle_equivalent
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
        "idle_hour_factor": idle_equivalent,
        "remaining_km": round(interval - equivalent, 2),
        "usage_percent": round(usage, 2),
        "pm_score": order.get("pm_score"),
        "pm_evidence": order.get("pm_evidence", ""),
        "status": status,
        "note": "لا توجد بيانات GPS بعد الصيانة" if status == "No GPS Data" else "لم يتم العثور على أمر مصنف صيانة وقائية من المواد" if status == "No Previous PM" else "",
    }


def export_report(result: AnalysisResult, path: Path) -> None:
    """Export the eleven-sheet Arabic preventive-maintenance report."""
    columns = [
        ("asset_id", "كود المعدة"), ("order_id", "رقم آخر أمر صيانة"),
        ("pm_date", "تاريخ آخر صيانة"), ("work_center", "مركز العمل"), ("brand", "الشركة"),
        ("gps_first", "أول تاريخ GPS محسوب"), ("gps_last", "آخر تاريخ GPS محسوب"),
        ("gps_days", "عدد أيام GPS"), ("distance_km", "إجمالي الكيلومترات"),
        ("engine_hours", "إجمالي ساعات التشغيل"), ("idle_hours", "ساعات التشغيل الساكن"),
        ("idle_hour_factor", "معامل الساعة الساكنة كم/ساعة"),
        ("idle_equivalent_km", "الكيلومترات المكافئة للساكن"),
        ("equivalent_km", "إجمالي الكيلومترات المكافئة"), ("interval_km", "فترة الصيانة"),
        ("remaining_km", "المتبقي حتى الصيانة"), ("usage_percent", "نسبة استهلاك الدورة %"),
        ("pm_score", "درجة ثقة الصيانة الوقائية"),
        ("pm_evidence", "دليل التصنيف من المواد"),
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
    maintenance_headers = ["كود المعدة", "رقم الأمر", "مركز العمل", "نوع الأمر (للعرض فقط)", "تاريخ الإنشاء", "التاريخ المرجعي", "التكلفة", "المرحلة", "درجة PM", "دليل المواد"]
    maintenance_by_order = {str(row["order_id"]): row for row in result.maintenance}
    latest_maintenance = [
        maintenance_by_order[str(row["order_id"])] for row in result.analysis
    ]
    maintenance_rows = [[r.get(key) for key in ("asset_id", "order_id", "work_center", "order_type", "created_at", "reference_at", "actual_cost", "subphase", "pm_score", "pm_evidence")] for r in latest_maintenance]
    gps_headers = ["كود المعدة", "التاريخ", "المسافة كم", "ساعات التشغيل", "ساعات الساكن", "الشركة", "الملف"]
    gps_rows = [[r.get(key) for key in ("asset_id", "gps_date", "distance_km", "engine_hours", "idle_hours", "brand", "source_file")] for r in result.gps]
    material_headers = ["كود المعدة", "رقم الأمر", "التصنيف", "الوزن", "الكلمة المطابقة", "مجموعة المادة (للعرض فقط)", "المادة", "الكمية", "القيمة", "تاريخ الصرف", "المخزن", "نوع التقييم"]
    pm_materials = [row for row in result.materials if row["category"]]
    material_rows = [[r.get(key) for key in ("asset_id", "order_id", "category", "pm_weight", "matched_keyword", "material_group", "material", "quantity", "value", "posting_date", "storage_location", "valuation_type")] for r in pm_materials]
    classification_headers = ["كود المعدة", "رقم الأمر", "التصنيف", "درجة PM", "حد PM", "دليل المواد", "عدد المواد", "آخر تاريخ صرف", "نسخة الخريطة"]
    classification_rows = [[r.get(key) for key in ("asset_id", "order_id", "classification", "pm_score", "pm_threshold", "evidence", "material_count", "posting_date", "map_version")] for r in result.order_classifications]
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
        ("تصنيف أوامر الصيانة", classification_headers, classification_rows),
        ("جودة البيانات", ["المشكلة", "المصدر", "الصف", "القيمة"], quality_rows),
    ])


def normalize_asset(value: str) -> str:
    """Normalize asset codes without guessing missing identifiers."""
    text = clean_text(value).upper()
    text = "".join("-" if unicodedata.category(character) == "Pd" else character for character in text)
    return re.sub(r"\s+", "", text)


def clean_text(value: str) -> str:
    return " ".join(unicodedata.normalize("NFKC", str(value)).replace("\u200b", "").split())


def classify_material(description: str, material_map: PMMaterialMap | None = None) -> str | None:
    match = (material_map or PMMaterialMap.load()).match(description)
    return match[0] if match else None


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
