from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .options import ChartData, SeriesData

_CHART_TYPES = ["barChart", "lineChart", "areaChart", "pieChart", "scatterChart"]
_CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")


def get_chart_data(workspace: Path, chart_index: int) -> ChartData:
    if chart_index < 1:
        raise ValueError("chart_index must be >= 1")

    chart_path = workspace / "word" / "charts" / f"chart{chart_index}.xml"
    if not chart_path.exists():
        raise FileNotFoundError(f"chart{chart_index}.xml not found")

    chart_xml = chart_path.read_text(encoding="utf-8")
    chart_title, cat_title, val_title = _extract_titles(chart_xml)
    chart_type = _find_chart_type(chart_xml)

    try:
        workbook_path = _resolve_workbook_for_chart(workspace, chart_index, chart_xml)
        categories, series = _read_workbook_data(workbook_path, chart_type)
    except (FileNotFoundError, ValueError, zipfile.BadZipFile, KeyError, ET.ParseError, UnicodeDecodeError):
        categories, series = _read_chart_cache(chart_xml, chart_type)

    return ChartData(
        categories=categories,
        series=series,
        chart_title=chart_title,
        category_axis_title=cat_title,
        value_axis_title=val_title,
    )


def _find_chart_type(content: str) -> str:
    for t in _CHART_TYPES:
        if f"<c:{t}" in content or f"<{t}" in content:
            return t
    return ""


def _extract_titles(chart_xml: str) -> tuple[str, str, str]:
    try:
        root = ET.fromstring(chart_xml)
    except ET.ParseError:
        return "", "", ""

    ns = {
        "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }

    def _text(path: str) -> str:
        node = root.find(path, ns)
        if node is None:
            return ""
        texts = [t.text or "" for t in node.findall(".//a:t", ns)]
        return "".join(texts).strip()

    chart_title = _text(".//c:chart/c:title")
    cat_title = _text(".//c:catAx/c:title")
    val_title = _text(".//c:valAx/c:title")
    return chart_title, cat_title, val_title


def _read_workbook_data(xlsx_path: Path, chart_type: str) -> tuple[list[str], list[SeriesData]]:
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        entries = {name: zf.read(name) for name in zf.namelist()}

    worksheet_path = _resolve_worksheet_path(entries)
    if not worksheet_path:
        raise ValueError("no worksheet found in embedded workbook")

    shared_strings = _read_shared_strings(entries.get("xl/sharedStrings.xml"))
    sheet = entries[worksheet_path].decode("utf-8")
    rows = _read_sheet_rows(sheet, shared_strings)

    if not rows:
        return [], []

    header = rows.get(1, {})
    max_col = max(header.keys(), default=1)
    categories = [str(rows.get(r, {}).get(1, "")) for r in range(2, max(rows.keys(), default=1) + 1)]

    if chart_type == "scatterChart":
        series = _read_scatter_series(rows, header, max_col)
        return categories, series

    series = []
    for col in range(2, max_col + 1):
        name = str(header.get(col, "")).strip()
        if not name:
            continue
        values = [_to_float(rows.get(r, {}).get(col, 0)) for r in range(2, max(rows.keys()) + 1)]
        series.append(SeriesData(name=name, values=values))
    return categories, series


def _read_scatter_series(
    rows: dict[int, dict[int, object]], header: dict[int, object], max_col: int
) -> list[SeriesData]:
    series = []
    col = 2
    series_idx = 1
    max_row = max(rows.keys(), default=1)
    while col + 1 <= max_col:
        name = str(header.get(col + 1, "")).strip()
        if not name:
            name = f"Series {series_idx}"
        x_values = [_to_float(rows.get(r, {}).get(col, 0)) for r in range(2, max_row + 1)]
        y_values = [_to_float(rows.get(r, {}).get(col + 1, 0)) for r in range(2, max_row + 1)]
        series.append(SeriesData(name=name, values=y_values, x_values=x_values))
        col += 2
        series_idx += 1
    return series


def _read_shared_strings(raw: bytes | None) -> list[str]:
    if not raw:
        return []
    try:
        root = ET.fromstring(raw.decode("utf-8"))
    except ET.ParseError:
        return []
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    strings = []
    for si in root.findall(".//s:si", ns):
        texts = [t.text or "" for t in si.findall(".//s:t", ns)]
        strings.append("".join(texts))
    return strings


def _read_sheet_rows(sheet_xml: str, shared_strings: list[str]) -> dict[int, dict[int, object]]:
    rows: dict[int, dict[int, object]] = {}
    try:
        root = ET.fromstring(sheet_xml)
    except ET.ParseError:
        return rows
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    for row in root.findall(".//s:row", ns):
        r_idx = int(row.get("r", "0"))
        if r_idx == 0:
            continue
        row_vals: dict[int, object] = {}
        for cell in row.findall("s:c", ns):
            ref = cell.get("r", "")
            col_idx = _col_from_ref(ref)
            if col_idx == 0:
                continue
            row_vals[col_idx] = _cell_value(cell, shared_strings, ns)
        rows[r_idx] = row_vals
    return rows


def _cell_value(cell: ET.Element, shared_strings: list[str], ns: dict[str, str]) -> object:
    cell_type = cell.get("t", "")
    if cell_type == "inlineStr":
        t = cell.find(".//s:is/s:t", ns)
        return t.text if t is not None else ""
    v = cell.find("s:v", ns)
    if v is None or v.text is None:
        return ""
    if cell_type == "s":
        idx = int(v.text)
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
    return v.text


def _col_from_ref(ref: str) -> int:
    m = _CELL_REF_RE.match(ref)
    if not m:
        return 0
    return _column_letter_to_index(m.group(1))


def _column_letter_to_index(col: str) -> int:
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _to_float(value: object) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _read_chart_cache(chart_xml: str, chart_type: str) -> tuple[list[str], list[SeriesData]]:
    if not chart_type:
        return [], []
    try:
        root = ET.fromstring(chart_xml)
    except ET.ParseError:
        return [], []

    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    chart = root.find(f".//c:{chart_type}", ns)
    if chart is None:
        return [], []

    categories: list[str] = []
    series: list[SeriesData] = []
    for ser in chart.findall("c:ser", ns):
        name = _first_text(ser, "c:tx//c:v", ns)
        if chart_type == "scatterChart":
            x_vals = _collect_numbers(ser, "c:xVal//c:v", ns)
            y_vals = _collect_numbers(ser, "c:yVal//c:v", ns)
            if not categories and x_vals:
                categories = [str(v) for v in x_vals]
            series.append(SeriesData(name=name, values=y_vals, x_values=x_vals))
        else:
            if not categories:
                categories = _collect_text(ser, "c:cat//c:v", ns)
            values = _collect_numbers(ser, "c:val//c:v", ns)
            series.append(SeriesData(name=name, values=values))
    return categories, series


def _first_text(node: ET.Element, path: str, ns: dict[str, str]) -> str:
    found = node.find(path, ns)
    return found.text.strip() if found is not None and found.text else ""


def _collect_text(node: ET.Element, path: str, ns: dict[str, str]) -> list[str]:
    return [(t.text or "").strip() for t in node.findall(path, ns)]


def _collect_numbers(node: ET.Element, path: str, ns: dict[str, str]) -> list[float]:
    vals = []
    for t in node.findall(path, ns):
        if t.text:
            try:
                vals.append(float(t.text))
            except ValueError:
                vals.append(0.0)
    return vals


def _resolve_workbook_for_chart(workspace: Path, chart_index: int, chart_xml: str) -> Path:
    rel_id = _external_data_rel_id(chart_xml)
    if not rel_id:
        raise ValueError(f"chart{chart_index}.xml has no externalData relationship ID")
    rels_path = workspace / "word" / "charts" / "_rels" / f"chart{chart_index}.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    target = _find_relationship_target(rels_xml, rel_id)
    if not target:
        raise ValueError(f"relationship {rel_id} not found for chart{chart_index}")
    chart_dir = (workspace / "word" / "charts").resolve()
    resolved = _resolve_workspace_target(workspace, chart_dir, target)
    if not resolved.exists():
        raise FileNotFoundError(f"workbook not found at {resolved}")
    return resolved


def _external_data_rel_id(chart_xml: str) -> str:
    marker = "<c:externalData"
    if marker not in chart_xml:
        marker = "<externalData"
    start = chart_xml.find(marker)
    if start == -1:
        return ""
    end = chart_xml.find(">", start)
    tag = chart_xml[start:end]
    for attr in ['r:id="', 'relationships:id="']:
        idx = tag.find(attr)
        if idx != -1:
            value_start = idx + len(attr)
            value_end = tag.find('"', value_start)
            if value_end != -1:
                return tag[value_start:value_end]
    return ""


def _find_relationship_target(rels_xml: str, rel_id: str) -> str:
    m = re.search(rf'Id="{re.escape(rel_id)}"[^>]*Target="([^"]+)"', rels_xml)
    return m.group(1) if m else ""


def _resolve_worksheet_path(entries: dict[str, bytes]) -> str:
    if "xl/workbook.xml" not in entries or "xl/_rels/workbook.xml.rels" not in entries:
        return _first_worksheet(entries)
    wb = entries["xl/workbook.xml"].decode("utf-8")
    rels = entries["xl/_rels/workbook.xml.rels"].decode("utf-8")
    m = re.search(r'r:id="([^"]+)"', wb)
    if not m:
        return _first_worksheet(entries)
    rel_id = m.group(1)
    m = re.search(rf'Id="{re.escape(rel_id)}"[^>]*Target="([^"]+)"', rels)
    if not m:
        return _first_worksheet(entries)
    target = "xl/" + m.group(1).lstrip("/")
    return target if target in entries else _first_worksheet(entries)


def _first_worksheet(entries: dict[str, bytes]) -> str:
    for name in entries:
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
            return name
    return ""


def _resolve_workspace_target(workspace: Path, base_dir: Path, target: str) -> Path:
    workspace_root = workspace.resolve()
    resolved = (base_dir / Path(target)).resolve()
    if workspace_root != resolved and workspace_root not in resolved.parents:
        raise ValueError(f"relationship target escapes workspace: {target}")
    return resolved
