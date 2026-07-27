"""Dependency-free XLSX reading and report writing.

The implementation intentionally supports the tabular XLSX subset used by the SAP/GPS
exports. Source workbooks are opened read-only and never modified.
"""

from __future__ import annotations

import html
import re
from collections.abc import Iterable, Sequence
from datetime import date, datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def read_workbook(path: Path) -> dict[str, list[dict[str, str]]]:
    """Return every worksheet as a list of header-keyed string rows."""
    with ZipFile(path) as archive:
        strings = _shared_strings(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {item.attrib["Id"]: item.attrib["Target"] for item in relationships}
        result: dict[str, list[dict[str, str]]] = {}
        sheets = workbook.find(f"{{{_MAIN}}}sheets")
        if sheets is None:
            return result
        for sheet in sheets:
            target = targets[sheet.attrib[f"{{{_REL}}}id"]]
            member = target if target.startswith("xl/") else f"xl/{target.lstrip('/')}"
            matrix = _read_sheet(archive, member, strings)
            if not matrix:
                result[sheet.attrib["name"]] = []
                continue
            headers = matrix[0]
            rows = []
            for values in matrix[1:]:
                rows.append({header: values[index] if index < len(values) else "" for index, header in enumerate(headers) if header})
            result[sheet.attrib["name"]] = rows
        return result


def _shared_strings(archive: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in item.iter(f"{{{_MAIN}}}t")) for item in root]


def _read_sheet(archive: ZipFile, member: str, strings: list[str]) -> list[list[str]]:
    root = ET.fromstring(archive.read(member))
    matrix: list[list[str]] = []
    for row in root.findall(f".//{{{_MAIN}}}sheetData/{{{_MAIN}}}row"):
        values: dict[int, str] = {}
        for cell in row.findall(f"{{{_MAIN}}}c"):
            reference = cell.attrib["r"]
            match = re.match(r"[A-Z]+", reference)
            if match is None:
                continue
            column = _column_number(match.group())
            value_node = cell.find(f"{{{_MAIN}}}v")
            inline = cell.find(f"{{{_MAIN}}}is")
            cell_type = cell.attrib.get("t")
            if cell_type == "s" and value_node is not None:
                value = strings[int(value_node.text or "0")]
            elif cell_type == "inlineStr" and inline is not None:
                value = "".join(node.text or "" for node in inline.iter(f"{{{_MAIN}}}t"))
            else:
                value = value_node.text if value_node is not None and value_node.text else ""
            values[column] = value
        if values:
            matrix.append([values.get(index, "") for index in range(max(values) + 1)])
    return matrix


def _column_number(letters: str) -> int:
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - 64
    return number - 1


def _column_letters(number: int) -> str:
    result = ""
    number += 1
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def write_report(path: Path, sheets: Sequence[tuple[str, Sequence[str], Iterable[Sequence[object]]]]) -> None:
    """Write a styled, right-to-left XLSX report."""
    path.parent.mkdir(parents=True, exist_ok=True)
    sheet_list = [(name[:31], list(headers), list(rows)) for name, headers, rows in sheets]
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types(len(sheet_list)))
        archive.writestr("_rels/.rels", _root_relationships())
        archive.writestr("xl/workbook.xml", _workbook(sheet_list))
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_relationships(len(sheet_list)))
        archive.writestr("xl/styles.xml", _styles())
        archive.writestr("docProps/core.xml", _core_properties())
        archive.writestr("docProps/app.xml", _app_properties(sheet_list))
        for index, (name, headers, rows) in enumerate(sheet_list, 1):
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                _sheet_xml(name, headers, rows, index),
            )


def _xml_cell(reference: str, value: object, style: int) -> str:
    if value is None:
        return f'<c r="{reference}" s="{style}"/>'
    if isinstance(value, bool):
        return f'<c r="{reference}" s="{style}" t="b"><v>{int(value)}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{reference}" s="{style}"><v>{value}</v></c>'
    if isinstance(value, (date, datetime)):
        value = value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    escaped = html.escape(str(value))
    return f'<c r="{reference}" s="{style}" t="inlineStr"><is><t>{escaped}</t></is></c>'


def _sheet_xml(
    name: str, headers: list[str], rows: list[Sequence[object]], sheet_index: int
) -> str:
    all_rows: list[Sequence[object]] = [headers, *rows]
    width = max((len(row) for row in all_rows), default=1)
    row_xml = []
    for row_number, row in enumerate(all_rows, 1):
        cells = []
        for column_number, value in enumerate(row):
            style = _cell_style(name, row_number, column_number, value)
            cells.append(_xml_cell(f"{_column_letters(column_number)}{row_number}", value, style))
        height = ' ht="30" customHeight="1"' if row_number == 1 else ''
        row_xml.append(f'<row r="{row_number}"{height}>{"".join(cells)}</row>')
    last = f"{_column_letters(width - 1)}{max(len(all_rows), 1)}"
    widths = _column_widths(all_rows, width)
    tab_colors = ("FF1F4E78", "FF4472C4", "FFC00000", "FFFFC000", "FF7F8C8D")
    tab_color = tab_colors[(sheet_index - 1) % len(tab_colors)]
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{_MAIN}"><sheetPr><tabColor rgb="{tab_color}"/></sheetPr>'
        '<sheetViews><sheetView rightToLeft="1" showGridLines="0" workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews><cols>'
        + "".join(
            f'<col min="{i}" max="{i}" width="{column_width}" customWidth="1"/>'
            for i, column_width in enumerate(widths, 1)
        )
        + f'</cols><sheetData>{"".join(row_xml)}</sheetData>'
        f'<autoFilter ref="A1:{last}"/><printOptions horizontalCentered="1"/>'
        '<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>'
        '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
        '</worksheet>'
    )


def _column_widths(rows: list[Sequence[object]], width: int) -> list[int]:
    result = []
    for column in range(width):
        longest = max(
            (len(str(row[column])) for row in rows[:250] if column < len(row)),
            default=10,
        )
        result.append(min(max(longest + 3, 12), 38))
    return result


def _cell_style(name: str, row_number: int, column: int, value: object) -> int:
    if row_number == 1:
        return 1
    if str(value) in {"OK", "سليم"}:
        return 2
    if str(value) in {"Due Soon", "صيانة قريبة"}:
        return 3
    if str(value) in {"Due", "مستحق"}:
        return 4
    if str(value) in {"No GPS Data", "No Previous PM"}:
        return 5
    if name == "لوحة التحكم":
        return 8 if column == 0 else 9
    if isinstance(value, (int, float)):
        return 6 if row_number % 2 == 0 else 7
    return 10 if row_number % 2 == 0 else 0


def _content_types(count: int) -> str:
    sheets = "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, count + 1))
    return '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' + sheets + '</Types>'


def _root_relationships() -> str:
    return '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'


def _workbook(sheets: list[tuple[str, list[str], list[Sequence[object]]]]) -> str:
    items = "".join(f'<sheet name="{html.escape(name)}" sheetId="{i}" r:id="rId{i}"/>' for i, (name, _, _) in enumerate(sheets, 1))
    return f'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{_MAIN}" xmlns:r="{_REL}"><sheets>{items}</sheets></workbook>'


def _workbook_relationships(count: int) -> str:
    items = "".join(f'<Relationship Id="rId{i}" Type="{_REL}/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, count + 1))
    return f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{items}<Relationship Id="rId{count + 1}" Type="{_REL}/styles" Target="styles.xml"/></Relationships>'


def _styles() -> str:
    colors = ("FF1F4E78", "FFC6EFCE", "FFFFEB9C", "FFFFC7CE", "FFD9E1F2", "FFF2F6FA", "FFD9EAF7", "FF5B9BD5")
    fills = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' + "".join(f'<fill><patternFill patternType="solid"><fgColor rgb="{color}"/><bgColor indexed="64"/></patternFill></fill>' for color in colors)
    border = '<border><left style="thin"><color rgb="FFD9E2F3"/></left><right style="thin"><color rgb="FFD9E2F3"/></right><top style="thin"><color rgb="FFD9E2F3"/></top><bottom style="thin"><color rgb="FFD9E2F3"/></bottom></border>'
    base = '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>'
    header = '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    statuses = "".join(f'<xf numFmtId="0" fontId="2" fillId="{fill}" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>' for fill in range(3, 7))
    numeric = '<xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0"/><xf numFmtId="4" fontId="0" fillId="7" borderId="1" xfId="0"/>'
    dashboard = '<xf numFmtId="0" fontId="2" fillId="8" borderId="1" xfId="0"/><xf numFmtId="4" fontId="3" fillId="9" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>'
    alternate = '<xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>'
    xfs = base + header + statuses + numeric + dashboard + alternate
    return f'<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{_MAIN}"><fonts count="4"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="14"/><name val="Arial"/></font></fonts><fills count="10">{fills}</fills><borders count="2"><border/>{border}</borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="11">{xfs}</cellXfs></styleSheet>'


def _core_properties() -> str:
    return '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>PM Analyzer</dc:creator><dc:title>Preventive Maintenance Analysis</dc:title></cp:coreProperties>'


def _app_properties(sheets: list[tuple[str, list[str], list[Sequence[object]]]]) -> str:
    names = "".join(f'<vt:lpstr>{html.escape(name)}</vt:lpstr>' for name, _, _ in sheets)
    return f'<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>PM Analyzer</Application><TitlesOfParts><vt:vector size="{len(sheets)}" baseType="lpstr">{names}</vt:vector></TitlesOfParts></Properties>'
