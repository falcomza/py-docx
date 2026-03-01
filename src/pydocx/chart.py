from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path

from .captions import generate_caption_xml, insert_caption_with_element
from .document import insert_at_body_end, insert_at_body_start
from .options import (
    AxisOptions,
    BarChartOptions,
    ChartKind,
    ChartOptions,
    ChartProperties,
    DataLabelOptions,
    InsertPosition,
    LegendOptions,
    ScatterChartOptions,
    SeriesOptions,
)
from .rels import (
    CHART_CONTENT_TYPE,
    REL_CHART_TYPE,
    REL_PACKAGE_TYPE,
    ensure_content_type_override,
    insert_relationship,
    next_relationship_id,
)
from .xmlutils import xml_escape

_CHART_FILE_RE = re.compile(r"^chart(\d+)\.xml$")


def insert_chart(workspace: Path, opts: ChartOptions) -> None:
    opts = _apply_chart_defaults(opts)
    _validate_chart_options(opts)
    chart_index = _find_next_chart_index(workspace)

    chart_path = workspace / "word" / "charts" / f"chart{chart_index}.xml"
    chart_path.parent.mkdir(parents=True, exist_ok=True)
    chart_path.write_text(_generate_chart_xml(opts), encoding="utf-8")

    workbook_path = workspace / "word" / "embeddings" / f"Microsoft_Excel_Worksheet{chart_index}.xlsx"
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    _create_embedded_workbook(workbook_path, opts)

    rels_path = workspace / "word" / "charts" / "_rels" / f"chart{chart_index}.xml.rels"
    rels_path.parent.mkdir(parents=True, exist_ok=True)
    rels_path.write_text(_chart_relationships_xml(workbook_path, chart_path.parent), encoding="utf-8")

    doc_rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    doc_rels_xml = doc_rels_path.read_text(encoding="utf-8")
    rel_id = next_relationship_id(doc_rels_xml)
    rel_xml = f'\n  <Relationship Id="{rel_id}" Type="{REL_CHART_TYPE}" Target="charts/chart{chart_index}.xml"/>'
    doc_rels_path.write_text(insert_relationship(doc_rels_xml, rel_xml), encoding="utf-8")

    _insert_chart_drawing(workspace, rel_id, chart_index, opts)

    content_types_path = workspace / "[Content_Types].xml"
    content_xml = content_types_path.read_text(encoding="utf-8")
    content_xml = ensure_content_type_override(
        content_xml,
        f"/word/charts/chart{chart_index}.xml",
        CHART_CONTENT_TYPE,
    )
    content_types_path.write_text(content_xml, encoding="utf-8")


def _validate_chart_options(opts: ChartOptions) -> None:
    if not opts.categories:
        raise ValueError("categories cannot be empty")
    if not opts.series:
        raise ValueError("series cannot be empty")
    chart_kind = _normalize_chart_kind(opts.chart_kind)
    if chart_kind not in {
        ChartKind.COLUMN.value,
        ChartKind.BAR.value,
        ChartKind.LINE.value,
        ChartKind.PIE.value,
        ChartKind.AREA.value,
        ChartKind.SCATTER.value,
    }:
        raise ValueError(f"unsupported chart kind: {opts.chart_kind}")
    for i, s in enumerate(opts.series):
        if not s.name.strip():
            raise ValueError(f"series[{i}] name cannot be empty")
        if len(s.values) != len(opts.categories):
            raise ValueError(
                f"series[{i}] values length ({len(s.values)}) must match categories length ({len(opts.categories)})"
            )
        if chart_kind == ChartKind.SCATTER.value and s.x_values is not None:
            if len(s.x_values) != len(s.values):
                raise ValueError(
                    f"series[{i}] x_values length ({len(s.x_values)}) must match values length ({len(s.values)})"
                )
    if opts.bar_chart_options is not None:
        if not 0 <= opts.bar_chart_options.gap_width <= 500:
            raise ValueError("bar_chart_options.gap_width must be between 0 and 500")
        if not -100 <= opts.bar_chart_options.overlap <= 100:
            raise ValueError("bar_chart_options.overlap must be between -100 and 100")


def _apply_chart_defaults(opts: ChartOptions) -> ChartOptions:
    if not opts.chart_kind:
        opts.chart_kind = ChartKind.COLUMN

    if opts.legend is None:
        position = opts.legend_position or "r"
        opts.legend = LegendOptions(show=opts.show_legend, position=position, overlay=False)

    cat_axis_created = opts.category_axis is None
    if opts.category_axis is None:
        opts.category_axis = AxisOptions(title=opts.category_axis_title)
    opts.category_axis = _apply_axis_defaults(
        opts.category_axis, is_category_axis=True, apply_defaults=cat_axis_created
    )

    val_axis_created = opts.value_axis is None
    if opts.value_axis is None:
        opts.value_axis = AxisOptions(title=opts.value_axis_title)
    opts.value_axis = _apply_axis_defaults(opts.value_axis, is_category_axis=False, apply_defaults=val_axis_created)

    if opts.properties is None:
        opts.properties = ChartProperties()
    if not opts.properties.style:
        opts.properties.style = 2
    if not opts.properties.language:
        opts.properties.language = "en-US"
    if not opts.properties.display_blanks_as:
        opts.properties.display_blanks_as = "gap"

    chart_kind = _normalize_chart_kind(opts.chart_kind)
    if chart_kind in (ChartKind.COLUMN.value, ChartKind.BAR.value):
        if opts.bar_chart_options is None:
            opts.bar_chart_options = BarChartOptions()
        if not opts.bar_chart_options.direction:
            opts.bar_chart_options.direction = "col" if chart_kind == ChartKind.COLUMN.value else "bar"
        if not opts.bar_chart_options.grouping:
            opts.bar_chart_options.grouping = "clustered"
        if opts.bar_chart_options.gap_width == 0:
            opts.bar_chart_options.gap_width = 150

    if opts.data_labels is not None and not opts.data_labels.position:
        opts.data_labels.position = "bestFit"

    return opts


def _apply_axis_defaults(axis: AxisOptions, is_category_axis: bool, apply_defaults: bool) -> AxisOptions:
    if not axis.position:
        axis.position = "b" if is_category_axis else "l"
    if not axis.major_tick_mark:
        axis.major_tick_mark = "out"
    if not axis.minor_tick_mark:
        axis.minor_tick_mark = "none"
    if not axis.tick_label_pos:
        axis.tick_label_pos = "nextTo"
    if not axis.number_format:
        axis.number_format = "General"
    if not is_category_axis and apply_defaults:
        axis.major_gridlines = True
    return axis


def _normalize_chart_kind(kind: ChartKind | str) -> str:
    if isinstance(kind, ChartKind):
        return kind.value
    return str(kind)


def _find_next_chart_index(workspace: Path) -> int:
    charts_dir = workspace / "word" / "charts"
    if not charts_dir.exists():
        return 1
    max_idx = 0
    for entry in charts_dir.iterdir():
        if entry.is_file():
            match = _CHART_FILE_RE.match(entry.name)
            if match:
                idx = int(match.group(1))
                max_idx = max(max_idx, idx)
    return max_idx + 1


def _generate_chart_xml(opts: ChartOptions) -> str:
    chart_kind = _normalize_chart_kind(opts.chart_kind)
    title_xml = _chart_title_xml(opts.title, opts.title_overlay)
    legend_xml = _legend_xml(opts.legend)
    plot_area_xml = _plot_area_xml(opts, chart_kind)
    props = opts.properties or ChartProperties()

    style_xml = f'<c:style val="{props.style}"/>' if props.style else ""
    rounded_xml = '<c:roundedCorners val="1"/>' if props.rounded_corners else ""
    show_over_max_xml = '<c:showDLblsOverMax val="1"/>' if props.show_data_labels_over_max else ""
    plot_vis_only = "1" if props.plot_visible_only else "0"
    disp_blanks_as = props.display_blanks_as or "gap"
    lang = props.language or "en-US"
    date1904_xml = '<c:date1904 val="1"/>' if props.date1904 else ""

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'{date1904_xml}<c:lang val="{lang}"/>'
        f'<c:chart>{title_xml}<c:autoTitleDeleted val="0"/>'
        f"<c:plotArea><c:layout/>{plot_area_xml}</c:plotArea>"
        f'{legend_xml}{style_xml}<c:plotVisOnly val="{plot_vis_only}"/>'
        f'<c:dispBlanksAs val="{disp_blanks_as}"/>{show_over_max_xml}{rounded_xml}'
        "</c:chart>"
        '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'
        "</c:chartSpace>"
    )


def _plot_area_xml(opts: ChartOptions, chart_kind: str) -> str:
    if chart_kind in (ChartKind.COLUMN.value, ChartKind.BAR.value):
        chart_xml = _bar_chart_xml(opts)
        axes_xml = _category_axis_xml(opts.category_axis) + _value_axis_xml(opts.value_axis)
        return chart_xml + axes_xml
    if chart_kind == ChartKind.LINE.value:
        chart_xml = _line_chart_xml(opts)
        axes_xml = _category_axis_xml(opts.category_axis) + _value_axis_xml(opts.value_axis)
        return chart_xml + axes_xml
    if chart_kind == ChartKind.AREA.value:
        chart_xml = _area_chart_xml(opts)
        axes_xml = _category_axis_xml(opts.category_axis) + _value_axis_xml(opts.value_axis)
        return chart_xml + axes_xml
    if chart_kind == ChartKind.SCATTER.value:
        chart_xml = _scatter_chart_xml(opts)
        x_axis_xml = _value_axis_xml(opts.category_axis, ax_id=2071991400, cross_ax_id=2071991240, is_x_axis=True)
        y_axis_xml = _value_axis_xml(opts.value_axis, ax_id=2071991240, cross_ax_id=2071991400, is_x_axis=False)
        return chart_xml + x_axis_xml + y_axis_xml
    if chart_kind == ChartKind.PIE.value:
        return _pie_chart_xml(opts)
    raise ValueError(f"unsupported chart kind: {chart_kind}")


def _chart_title_xml(title: str, overlay: bool) -> str:
    if not title:
        return ""
    overlay_val = "1" if overlay else "0"
    return (
        "<c:title>"
        "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>"
        "<a:pPr><a:defRPr/></a:pPr>"
        f"<a:r><a:t>{xml_escape(title)}</a:t></a:r>"
        "</a:p></c:rich></c:tx>"
        f'<c:layout/><c:overlay val="{overlay_val}"/>'
        "</c:title>"
    )


def _legend_xml(legend: LegendOptions | None) -> str:
    if legend is None or not legend.show:
        return ""
    overlay_val = "1" if legend.overlay else "0"
    pos = legend.position or "r"
    return f'<c:legend><c:legendPos val="{pos}"/><c:layout/><c:overlay val="{overlay_val}"/></c:legend>'


def _bar_chart_xml(opts: ChartOptions) -> str:
    series_xml = [_series_xml(i, s, opts.categories, chart_kind="bar") for i, s in enumerate(opts.series)]
    bar_opts = opts.bar_chart_options or BarChartOptions()
    dlabels_xml = _data_labels_xml(opts.data_labels)
    return (
        "<c:barChart>"
        f'<c:barDir val="{bar_opts.direction}"/>'
        f'<c:grouping val="{bar_opts.grouping}"/>'
        f'<c:varyColors val="{1 if bar_opts.vary_colors else 0}"/>'
        + "".join(series_xml)
        + dlabels_xml
        + f'<c:gapWidth val="{bar_opts.gap_width}"/>'
        + f'<c:overlap val="{bar_opts.overlap}"/>'
        + '<c:axId val="2071991400"/>'
        + '<c:axId val="2071991240"/>'
        + "</c:barChart>"
    )


def _line_chart_xml(opts: ChartOptions) -> str:
    series_xml = [_series_xml(i, s, opts.categories, chart_kind="line") for i, s in enumerate(opts.series)]
    dlabels_xml = _data_labels_xml(opts.data_labels)
    return (
        "<c:lineChart>"
        '<c:grouping val="standard"/>'
        '<c:varyColors val="0"/>'
        + "".join(series_xml)
        + dlabels_xml
        + '<c:axId val="2071991400"/>'
        + '<c:axId val="2071991240"/>'
        + "</c:lineChart>"
    )


def _area_chart_xml(opts: ChartOptions) -> str:
    series_xml = [_series_xml(i, s, opts.categories, chart_kind="area") for i, s in enumerate(opts.series)]
    dlabels_xml = _data_labels_xml(opts.data_labels)
    return (
        "<c:areaChart>"
        '<c:grouping val="standard"/>'
        '<c:varyColors val="0"/>'
        + "".join(series_xml)
        + dlabels_xml
        + '<c:axId val="2071991400"/>'
        + '<c:axId val="2071991240"/>'
        + "</c:areaChart>"
    )


def _pie_chart_xml(opts: ChartOptions) -> str:
    series_xml = [_series_xml(i, s, opts.categories, chart_kind="pie") for i, s in enumerate(opts.series)]
    dlabels_xml = _data_labels_xml(opts.data_labels)
    return '<c:pieChart><c:varyColors val="1"/>' + "".join(series_xml) + dlabels_xml + "</c:pieChart>"


def _scatter_chart_xml(opts: ChartOptions) -> str:
    series_xml = [_scatter_series_xml(i, s, opts.categories) for i, s in enumerate(opts.series)]
    scatter_opts = opts.scatter_chart_options or ScatterChartOptions()
    dlabels_xml = _data_labels_xml(opts.data_labels)
    return (
        "<c:scatterChart>"
        f'<c:scatterStyle val="{scatter_opts.scatter_style}"/>'
        f'<c:varyColors val="{1 if scatter_opts.vary_colors else 0}"/>'
        + "".join(series_xml)
        + dlabels_xml
        + '<c:axId val="2071991400"/>'
        + '<c:axId val="2071991240"/>'
        + "</c:scatterChart>"
    )


def _series_xml(index: int, series: SeriesOptions, categories: list[str], chart_kind: str) -> str:
    cat_cache = "".join(f'<c:pt idx="{i}"><c:v>{xml_escape(cat)}</c:v></c:pt>' for i, cat in enumerate(categories))
    val_cache = "".join(f'<c:pt idx="{i}"><c:v>{val}</c:v></c:pt>' for i, val in enumerate(series.values))
    col_letter = _column_letter(index + 2)
    sp_pr_xml = _series_shape_xml(series)
    invert_xml = '<c:invertIfNegative val="1"/>' if series.invert_if_negative and chart_kind == "bar" else ""
    smooth_xml = '<c:smooth val="1"/>' if series.smooth and chart_kind in ("line", "scatter") else ""
    marker_xml = _marker_xml(series) if chart_kind in ("line", "scatter") else ""
    dlabels_xml = _data_labels_xml(series.data_labels)
    return (
        f'<c:ser><c:idx val="{index}"/><c:order val="{index}"/>'
        f"<c:tx><c:strRef><c:f>Sheet1!${col_letter}$1</c:f>"
        f'<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{xml_escape(series.name)}</c:v></c:pt>'
        f"</c:strCache></c:strRef></c:tx>"
        f"{sp_pr_xml}{invert_xml}{dlabels_xml}{marker_xml}{smooth_xml}"
        f"<c:cat><c:strRef><c:f>Sheet1!$A$2:$A${len(categories) + 1}</c:f>"
        f'<c:strCache><c:ptCount val="{len(categories)}"/>{cat_cache}</c:strCache>'
        f"</c:strRef></c:cat>"
        f"<c:val><c:numRef><c:f>Sheet1!${col_letter}$2:${col_letter}${len(categories) + 1}</c:f>"
        f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(series.values)}"/>{val_cache}'
        f"</c:numCache></c:numRef></c:val>"
        f"</c:ser>"
    )


def _scatter_series_xml(index: int, series: SeriesOptions, categories: list[str]) -> str:
    x_values = series.x_values or [float(i + 1) for i in range(len(series.values))]
    x_cache = "".join(f'<c:pt idx="{i}"><c:v>{val}</c:v></c:pt>' for i, val in enumerate(x_values))
    y_cache = "".join(f'<c:pt idx="{i}"><c:v>{val}</c:v></c:pt>' for i, val in enumerate(series.values))
    x_col = _column_letter(2 + index * 2)
    y_col = _column_letter(3 + index * 2)
    sp_pr_xml = _series_shape_xml(series)
    smooth_xml = '<c:smooth val="1"/>' if series.smooth else ""
    marker_xml = _marker_xml(series)
    dlabels_xml = _data_labels_xml(series.data_labels)
    return (
        f'<c:ser><c:idx val="{index}"/><c:order val="{index}"/>'
        f"<c:tx><c:strRef><c:f>Sheet1!${y_col}$1</c:f>"
        f'<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{xml_escape(series.name)}</c:v></c:pt>'
        f"</c:strCache></c:strRef></c:tx>"
        f"{sp_pr_xml}{dlabels_xml}{marker_xml}{smooth_xml}"
        f"<c:xVal><c:numRef><c:f>Sheet1!${x_col}$2:${x_col}${len(x_values) + 1}</c:f>"
        f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(x_values)}"/>{x_cache}'
        f"</c:numCache></c:numRef></c:xVal>"
        f"<c:yVal><c:numRef><c:f>Sheet1!${y_col}$2:${y_col}${len(series.values) + 1}</c:f>"
        f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(series.values)}"/>{y_cache}'
        f"</c:numCache></c:numRef></c:yVal>"
        f"</c:ser>"
    )


def _series_shape_xml(series: SeriesOptions) -> str:
    if not series.color:
        return ""
    return f'<c:spPr><a:solidFill><a:srgbClr val="{series.color}"/></a:solidFill></c:spPr>'


def _marker_xml(series: SeriesOptions) -> str:
    if series.show_markers:
        return ""
    return '<c:marker><c:symbol val="none"/></c:marker>'


def _data_labels_xml(opts: DataLabelOptions | None) -> str:
    if opts is None:
        return ""
    pos_xml = f'<c:dLblPos val="{opts.position}"/>' if opts.position else ""
    leader_xml = f'<c:showLeaderLines val="{1 if opts.show_leader_lines else 0}"/>' if opts.show_leader_lines else ""
    return (
        "<c:dLbls>"
        f'<c:showLegendKey val="{1 if opts.show_legend_key else 0}"/>'
        f'<c:showVal val="{1 if opts.show_value else 0}"/>'
        f'<c:showCatName val="{1 if opts.show_category_name else 0}"/>'
        f'<c:showSerName val="{1 if opts.show_series_name else 0}"/>'
        f'<c:showPercent val="{1 if opts.show_percent else 0}"/>'
        f"{leader_xml}{pos_xml}"
        "</c:dLbls>"
    )


def _axis_title_xml(axis: AxisOptions) -> str:
    if not axis.title:
        return ""
    overlay_val = "1" if axis.title_overlay else "0"
    return (
        "<c:title>"
        "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>"
        "<a:pPr><a:defRPr/></a:pPr>"
        f"<a:r><a:t>{xml_escape(axis.title)}</a:t></a:r>"
        "</a:p></c:rich></c:tx>"
        f'<c:layout/><c:overlay val="{overlay_val}"/>'
        "</c:title>"
    )


def _category_axis_xml(axis: AxisOptions) -> str:
    delete_val = "0" if axis.visible else "1"
    title_xml = _axis_title_xml(axis)
    return (
        '<c:catAx><c:axId val="2071991400"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        f'<c:delete val="{delete_val}"/><c:axPos val="{axis.position}"/>'
        f"{title_xml}"
        f'<c:majorTickMark val="{axis.major_tick_mark}"/><c:minorTickMark val="{axis.minor_tick_mark}"/>'
        f'<c:tickLblPos val="{axis.tick_label_pos}"/>'
        '<c:crossAx val="2071991240"/><c:crosses val="autoZero"/>'
        '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
        "</c:catAx>"
    )


def _value_axis_xml(
    axis: AxisOptions, ax_id: int = 2071991240, cross_ax_id: int = 2071991400, is_x_axis: bool = False
) -> str:
    delete_val = "0" if axis.visible else "1"
    title_xml = _axis_title_xml(axis)
    scaling_parts = ['<c:scaling><c:orientation val="minMax"/>']
    if axis.min is not None:
        scaling_parts.append(f'<c:min val="{axis.min}"/>')
    if axis.max is not None:
        scaling_parts.append(f'<c:max val="{axis.max}"/>')
    scaling_parts.append("</c:scaling>")
    scaling_xml = "".join(scaling_parts)
    grid_xml = ""
    if axis.major_gridlines:
        grid_xml += "<c:majorGridlines/>"
    if axis.minor_gridlines:
        grid_xml += "<c:minorGridlines/>"
    major_unit_xml = f'<c:majorUnit val="{axis.major_unit}"/>' if axis.major_unit is not None else ""
    minor_unit_xml = f'<c:minorUnit val="{axis.minor_unit}"/>' if axis.minor_unit is not None else ""
    crosses_xml = (
        f'<c:crossesAt val="{axis.crosses_at}"/>' if axis.crosses_at is not None else '<c:crosses val="autoZero"/>'
    )
    cross_between_xml = "" if is_x_axis else '<c:crossBetween val="between"/>'
    return (
        f'<c:valAx><c:axId val="{ax_id}"/>'
        f"{scaling_xml}"
        f'<c:delete val="{delete_val}"/><c:axPos val="{axis.position}"/>'
        f"{grid_xml}"
        f'{title_xml}<c:numFmt formatCode="{axis.number_format}" sourceLinked="0"/>'
        f"{major_unit_xml}{minor_unit_xml}"
        f'<c:majorTickMark val="{axis.major_tick_mark}"/><c:minorTickMark val="{axis.minor_tick_mark}"/>'
        f'<c:tickLblPos val="{axis.tick_label_pos}"/>'
        f'<c:crossAx val="{cross_ax_id}"/>{crosses_xml}'
        f"{cross_between_xml}"
        "</c:valAx>"
    )


def _column_letter(col: int) -> str:
    result = ""
    while col > 0:
        col -= 1
        result = chr(ord("A") + col % 26) + result
        col //= 26
    return result


def _create_embedded_workbook(path: Path, opts: ChartOptions) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _workbook_content_types())
        zf.writestr("_rels/.rels", _workbook_rels())
        zf.writestr("xl/workbook.xml", _workbook_xml())
        zf.writestr("xl/_rels/workbook.xml.rels", _workbook_xml_rels())
        zf.writestr("xl/worksheets/sheet1.xml", _sheet_xml(opts))
        zf.writestr("xl/styles.xml", _styles_xml())


def _workbook_content_types() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )


def _workbook_rels() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def _workbook_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )


def _workbook_xml_rels() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )


def _sheet_xml(opts: ChartOptions) -> str:
    rows = []
    chart_kind = _normalize_chart_kind(opts.chart_kind)
    if chart_kind == ChartKind.SCATTER.value:
        header = [""]
        for s in opts.series:
            header.extend(["x", s.name])
        rows.append(_sheet_row(1, header, string_row=True))
        for i, category in enumerate(opts.categories):
            row_vals: list[object] = [category]
            for s in opts.series:
                x_vals = s.x_values or [float(j + 1) for j in range(len(s.values))]
                row_vals.append(x_vals[i])
                row_vals.append(s.values[i])
            rows.append(_sheet_row(i + 2, row_vals, string_row=False))
    else:
        rows.append(_sheet_row(1, [""] + [s.name for s in opts.series], string_row=True))
        for i, category in enumerate(opts.categories):
            values = [category] + [s.values[i] for s in opts.series]
            rows.append(_sheet_row(i + 2, values, string_row=False))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheetData>" + "".join(rows) + "</sheetData></worksheet>"
    )


def _sheet_row(row_num: int, values: list[object], string_row: bool) -> str:
    cells = []
    for i, value in enumerate(values):
        ref = f"{_column_letter(i + 1)}{row_num}"
        if i == 0 or string_row:
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{xml_escape(str(value))}</t></is></c>')
        else:
            cells.append(f'<c r="{ref}"><v>{value}</v></c>')
    return f'<row r="{row_num}">{"".join(cells)}</row>'


def _styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<numFmts count="0"/>'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>'
        "</styleSheet>"
    )


def _chart_relationships_xml(workbook_path: Path, charts_dir: Path) -> str:
    rel_path = os.path.relpath(workbook_path, charts_dir)
    rel_path = rel_path.replace(os.sep, "/")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{REL_PACKAGE_TYPE}" Target="{rel_path}"/>'
        "</Relationships>"
    )


def _insert_chart_drawing(workspace: Path, rel_id: str, chart_index: int, opts: ChartOptions) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    doc_pr_id = _next_docpr_id(doc_xml)
    drawing_xml = (
        "<w:p><w:r><w:drawing>"
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{opts.width_emu}" cy="{opts.height_emu}"/>'
        f'<wp:docPr id="{doc_pr_id}" name="Chart {chart_index}"/>'
        "<wp:cNvGraphicFramePr/>"
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'r:id="{rel_id}"/>'
        "</a:graphicData></a:graphic>"
        "</wp:inline></w:drawing></w:r></w:p>"
    )
    if opts.caption is not None:
        if getattr(opts.caption, "type", None) is None:
            from .options import CaptionType

            opts.caption.type = CaptionType.FIGURE
        caption_xml = generate_caption_xml(opts.caption)
        drawing_xml = insert_caption_with_element(caption_xml, drawing_xml, opts.caption.position)

    if opts.position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, drawing_xml)
    elif opts.position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, drawing_xml)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")

    doc_path.write_text(updated, encoding="utf-8")


_DOCPR_RE = re.compile(r'wp:docPr id="(\d+)"')


def _next_docpr_id(doc_xml: str) -> int:
    max_id = 0
    for match in _DOCPR_RE.finditer(doc_xml):
        val = int(match.group(1))
        if val > max_id:
            max_id = val
    return max_id + 1
