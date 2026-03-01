from __future__ import annotations

import re
import zipfile
from pathlib import Path

from .options import ChartData, SeriesData
from .xmlutils import xml_escape

_CHART_TYPES = ["barChart", "lineChart", "areaChart", "pieChart", "scatterChart"]


def update_chart(workspace: Path, chart_index: int, data: ChartData) -> None:
    if chart_index < 1:
        raise ValueError("chart_index must be >= 1")
    _validate_chart_data(data)

    chart_path = workspace / "word" / "charts" / f"chart{chart_index}.xml"
    if not chart_path.exists():
        raise FileNotFoundError(f"chart{chart_index}.xml not found")

    chart_xml = chart_path.read_text(encoding="utf-8")
    updated_xml = _update_chart_xml(chart_xml, data)
    chart_path.write_text(updated_xml, encoding="utf-8")

    workbook_path = _resolve_workbook_for_chart(workspace, chart_index, chart_xml)
    _update_embedded_workbook(workbook_path, data)


def _validate_chart_data(data: ChartData) -> None:
    if not data.categories:
        raise ValueError("categories cannot be empty")
    if not data.series:
        raise ValueError("series cannot be empty")
    for i, s in enumerate(data.series):
        if not s.name.strip():
            raise ValueError(f"series[{i}] name cannot be empty")
        if len(s.values) != len(data.categories):
            raise ValueError(
                f"series[{i}] values length ({len(s.values)}) must match categories length ({len(data.categories)})"
            )


def _update_chart_xml(content: str, data: ChartData) -> str:
    ns = "c:" if "<c:chart" in content else ""

    if data.chart_title:
        content = _update_title(content, data.chart_title, ns)
    if data.category_axis_title or data.value_axis_title:
        content = _update_axis_titles(content, data.category_axis_title, data.value_axis_title, ns)

    chart_type = _find_chart_type(content, ns)
    if not chart_type:
        raise ValueError("unsupported or missing chart type")

    updated = _replace_series_section(content, data, ns, chart_type)
    return _ensure_xml_declaration_newline(updated)


def _find_chart_type(content: str, ns: str) -> str:
    for t in _CHART_TYPES:
        if f"<{ns}{t}" in content:
            return t
    return ""


def _replace_series_section(content: str, data: ChartData, ns: str, chart_type: str) -> str:
    open_tag = f"<{ns}{chart_type}"
    close_tag = f"</{ns}{chart_type}>"
    start = content.find(open_tag)
    if start == -1:
        raise ValueError(f"chart type {chart_type} not found")
    end = content.find(close_tag, start)
    if end == -1:
        raise ValueError(f"chart type {chart_type} is unclosed")
    end += len(close_tag)

    section = content[start:end]
    series_xml = "".join(_build_series_xml(i, s, data.categories, ns, chart_type) for i, s in enumerate(data.series))
    tail = _copy_non_series_elements(section, ns)
    rebuilt = f"<{ns}{chart_type}>" + series_xml + tail + f"</{ns}{chart_type}>"

    return content[:start] + rebuilt + content[end:]


def _build_series_xml(index: int, series: SeriesData, categories: list[str], ns: str, chart_type: str) -> str:
    col_letter = _column_letter(index + 2)
    ser = [f'<{ns}ser><{ns}idx val="{index}"/><{ns}order val="{index}"/>']
    ser.append(
        f"<{ns}tx><{ns}strRef><{ns}f>Sheet1!${col_letter}$1</{ns}f>"
        f'<{ns}strCache><{ns}ptCount val="1"/>'
        f'<{ns}pt idx="0"><{ns}v>{xml_escape(series.name)}</{ns}v></{ns}pt>'
        f"</{ns}strCache></{ns}strRef></{ns}tx>"
    )

    if chart_type != "scatterChart":
        cat_pts = "".join(
            f'<{ns}pt idx="{i}"><{ns}v>{xml_escape(cat)}</{ns}v></{ns}pt>' for i, cat in enumerate(categories)
        )
        ser.append(
            f"<{ns}cat><{ns}strRef><{ns}f>Sheet1!$A$2:$A${len(categories) + 1}</{ns}f>"
            f'<{ns}strCache><{ns}ptCount val="{len(categories)}"/>{cat_pts}</{ns}strCache>'
            f"</{ns}strRef></{ns}cat>"
        )
    else:
        xvals = series.x_values
        if xvals is None:
            xvals = _parse_scatter_xvalues(categories)
        x_pts = "".join(f'<{ns}pt idx="{i}"><{ns}v>{v}</{ns}v></{ns}pt>' for i, v in enumerate(xvals))
        ser.append(
            f'<{ns}xVal><{ns}numRef><{ns}numCache><{ns}ptCount val="{len(xvals)}"/>'
            f"{x_pts}</{ns}numCache></{ns}numRef></{ns}xVal>"
        )

    val_tag = "yVal" if chart_type == "scatterChart" else "val"
    val_pts = "".join(f'<{ns}pt idx="{i}"><{ns}v>{v}</{ns}v></{ns}pt>' for i, v in enumerate(series.values))
    ser.append(
        f"<{ns}{val_tag}><{ns}numRef><{ns}f>Sheet1!${col_letter}$2:${col_letter}${len(categories) + 1}</{ns}f>"
        f'<{ns}numCache><{ns}formatCode>General</{ns}formatCode><{ns}ptCount val="{len(series.values)}"/>'
        f"{val_pts}</{ns}numCache></{ns}numRef></{ns}{val_tag}>"
    )

    ser.append(f"</{ns}ser>")
    return "".join(ser)


def _copy_non_series_elements(section: str, ns: str) -> str:
    buf = []
    for tag in ["axId", "dLbls", "gapWidth", "overlap", "varyColors", "scatterStyle"]:
        start_tag = f"<{ns}{tag}"
        offset = 0
        while True:
            pos = section.find(start_tag, offset)
            if pos == -1:
                break
            end_empty = section.find("/>", pos)
            if end_empty != -1:
                buf.append(section[pos : end_empty + 2])
                offset = end_empty + 2
                continue
            end_tag = f"</{ns}{tag}>"
            end = section.find(end_tag, pos)
            if end == -1:
                break
            buf.append(section[pos : end + len(end_tag)])
            offset = end + len(end_tag)
    return "".join(buf)


def _update_title(content: str, title: str, ns: str) -> str:
    title_start = content.find(f"<{ns}title>")
    if title_start == -1:
        return content
    title_end = content.find(f"</{ns}title>", title_start)
    if title_end == -1:
        return content
    title_section = content[title_start:title_end]
    t_start = title_section.find(f"<{ns}t>")
    if t_start == -1:
        return content
    t_end = title_section.find(f"</{ns}t>", t_start)
    if t_end == -1:
        return content
    abs_start = title_start + t_start + len(f"<{ns}t>")
    abs_end = title_start + t_end
    before = content[:abs_start]
    after = content[abs_end:]
    return before + xml_escape(title) + after


def _update_axis_titles(content: str, cat_title: str, val_title: str, ns: str) -> str:
    if cat_title:
        content = _update_first_axis_title(content, cat_title, ns, "catAx")
    if val_title:
        content = _update_first_axis_title(content, val_title, ns, "valAx")
    return content


def _update_first_axis_title(content: str, title: str, ns: str, axis_tag: str) -> str:
    axis_start = content.find(f"<{ns}{axis_tag}>")
    if axis_start == -1:
        return content
    axis_end = content.find(f"</{ns}{axis_tag}>", axis_start)
    if axis_end == -1:
        return content
    axis_section = content[axis_start:axis_end]
    title_start = axis_section.find(f"<{ns}title>")
    if title_start == -1:
        return content
    t_start = axis_section.find(f"<{ns}t>", title_start)
    if t_start == -1:
        return content
    t_end = axis_section.find(f"</{ns}t>", t_start)
    if t_end == -1:
        return content
    abs_start = axis_start + t_start + len(f"<{ns}t>")
    abs_end = axis_start + t_end
    before = content[:abs_start]
    after = content[abs_end:]
    return before + xml_escape(title) + after


def _ensure_xml_declaration_newline(content: str) -> str:
    idx = content.find("?>")
    if idx == -1:
        return content
    idx += 2
    if idx < len(content) and content[idx] == "\n":
        return content
    return content[:idx] + "\n" + content[idx:]


def _parse_scatter_xvalues(categories: list[str]) -> list[float]:
    vals = []
    for cat in categories:
        stripped = cat.strip()
        if not stripped:
            raise ValueError("scatter chart categories must be numeric")
        vals.append(float(stripped))
    return vals


def _column_letter(col: int) -> str:
    result = ""
    while col > 0:
        col -= 1
        result = chr(ord("A") + col % 26) + result
        col //= 26
    return result


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


def _update_embedded_workbook(xlsx_path: Path, data: ChartData) -> None:
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        entries = {name: zf.read(name) for name in zf.namelist()}

    worksheet_path = _resolve_worksheet_path(entries)
    if not worksheet_path:
        raise ValueError("no worksheet found in embedded workbook")

    use_shared_strings = "xl/sharedStrings.xml" in entries
    string_indexes: dict[str, int] = {}
    if use_shared_strings:
        entries["xl/sharedStrings.xml"], string_indexes = _update_shared_strings(entries["xl/sharedStrings.xml"], data)

    entries[worksheet_path] = _update_worksheet(entries[worksheet_path], data, use_shared_strings, string_indexes)

    with zipfile.ZipFile(xlsx_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries.items():
            zf.writestr(name, content)


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


def _update_worksheet(raw: bytes, data: ChartData, use_shared_strings: bool, string_indexes: dict[str, int]) -> bytes:
    content = raw.decode("utf-8")
    sheet_data = _build_sheet_data(data, use_shared_strings, string_indexes)
    content = re.sub(r"(?s)<sheetData\b[^>]*>.*?</sheetData>", sheet_data, content)
    last_col = _column_letter(len(data.series) + 1)
    last_row = len(data.categories) + 1
    dimension = f'<dimension ref="A1:{last_col}{last_row}"/>'
    content = re.sub(r'<dimension\b[^>]*ref="[^"]*"[^>]*/>', dimension, content)
    return content.encode("utf-8")


def _build_sheet_data(data: ChartData, use_shared_strings: bool, string_indexes: dict[str, int]) -> str:
    rows = []
    header = [""] + [s.name for s in data.series]
    rows.append(_sheet_row(1, header, use_shared_strings, string_indexes, header_row=True))
    for i, cat in enumerate(data.categories):
        values = [cat] + [s.values[i] for s in data.series]
        rows.append(_sheet_row(i + 2, values, use_shared_strings, string_indexes, header_row=False))
    return "<sheetData>" + "".join(rows) + "</sheetData>"


def _sheet_row(
    row_num: int,
    values: list[object],
    use_shared_strings: bool,
    string_indexes: dict[str, int],
    header_row: bool,
) -> str:
    cells = []
    for i, value in enumerate(values):
        ref = f"{_column_letter(i + 1)}{row_num}"
        if i == 0 or header_row:
            if use_shared_strings:
                idx = string_indexes.get(str(value), 0)
                cells.append(f'<c r="{ref}" t="s"><v>{idx}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{xml_escape(str(value))}</t></is></c>')
        else:
            cells.append(f'<c r="{ref}"><v>{value}</v></c>')
    return f'<row r="{row_num}">{"".join(cells)}</row>'


def _update_shared_strings(raw: bytes, data: ChartData) -> tuple[bytes, dict[str, int]]:
    strings = []
    strings.append("")
    for s in data.series:
        strings.append(s.name)
    for c in data.categories:
        strings.append(c)
    unique = []
    for s in strings:
        if s not in unique:
            unique.append(s)
    indexes = {s: i for i, s in enumerate(unique)}
    si = "".join(f"<si><t>{xml_escape(s)}</t></si>" for s in unique)
    updated = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(unique)}" uniqueCount="{len(unique)}">{si}</sst>'
    )
    return updated.encode("utf-8"), indexes


def _resolve_workspace_target(workspace: Path, base_dir: Path, target: str) -> Path:
    workspace_root = workspace.resolve()
    resolved = (base_dir / Path(target)).resolve()
    if workspace_root != resolved and workspace_root not in resolved.parents:
        raise ValueError(f"relationship target escapes workspace: {target}")
    return resolved
