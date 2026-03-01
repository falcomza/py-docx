from __future__ import annotations

from pathlib import Path

from .captions import generate_caption_xml, insert_caption_with_element
from .document import insert_at_body_end, insert_at_body_start
from .options import (
    BorderStyle,
    CellTextStyle,
    ConditionalStyle,
    InsertPosition,
    RowHeightRule,
    TableOptions,
    TableWidthType,
)
from .xmlutils import xml_escape


def insert_table(workspace: Path, opts: TableOptions) -> None:
    if not opts.columns:
        raise ValueError("columns cannot be empty")
    if any(len(r) > len(opts.columns) for r in opts.rows):
        raise ValueError("row length exceeds number of columns")
    if opts.column_widths is not None and len(opts.column_widths) != len(opts.columns):
        raise ValueError("column_widths length must match columns length")

    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")

    table_xml = _build_table_xml(opts)
    if opts.caption is not None:
        if getattr(opts.caption, "type", None) is None:
            from .options import CaptionType

            opts.caption.type = CaptionType.TABLE
        caption_xml = generate_caption_xml(opts.caption)
        table_xml = insert_caption_with_element(caption_xml, table_xml, opts.caption.position)

    if opts.position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, table_xml)
    elif opts.position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, table_xml)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")

    doc_path.write_text(updated, encoding="utf-8")


def _build_table_xml(opts: TableOptions) -> str:
    grid_cols = []
    if opts.column_widths is not None:
        for width in opts.column_widths:
            grid_cols.append(f'<w:gridCol w:w="{width}"/>')
    elif opts.proportional_column_widths:
        widths = _calculate_proportional_widths(opts)
        for width in widths:
            grid_cols.append(f'<w:gridCol w:w="{width}"/>')
    else:
        for col in opts.columns:
            if col.width:
                grid_cols.append(f'<w:gridCol w:w="{col.width}"/>')
            else:
                grid_cols.append("<w:gridCol/>")

    tbl_width_type = opts.table_width_type.value
    tbl_width_val = 0 if opts.table_width_type == TableWidthType.AUTO else opts.table_width
    table_pr = [
        "<w:tblPr>",
        f'<w:tblStyle w:val="{xml_escape(opts.table_style)}"/>',
        f'<w:tblW w:w="{tbl_width_val}" w:type="{tbl_width_type}"/>',
        f'<w:jc w:val="{opts.table_alignment.value}"/>',
        _table_borders_xml(opts),
        f'<w:tblCellMar><w:top w:w="{opts.cell_padding}" w:type="dxa"/>'
        f'<w:left w:w="{opts.cell_padding}" w:type="dxa"/>'
        f'<w:bottom w:w="{opts.cell_padding}" w:type="dxa"/>'
        f'<w:right w:w="{opts.cell_padding}" w:type="dxa"/></w:tblCellMar>',
        "</w:tblPr>",
    ]

    table_grid = "<w:tblGrid>" + "".join(grid_cols) + "</w:tblGrid>"

    rows_xml = []
    rows_xml.append(_build_row_xml([c.title for c in opts.columns], opts, header=True, row_index=1))
    for i, row in enumerate(opts.rows):
        background = opts.row_background
        if opts.alternate_row_background and (i % 2 == 1):
            background = opts.alternate_row_background
        padded = row + [""] * (len(opts.columns) - len(row))
        rows_xml.append(_build_row_xml(padded, opts, header=False, background=background, row_index=i + 2))

    return "<w:tbl>" + "".join(table_pr) + table_grid + "".join(rows_xml) + "</w:tbl>"


def _build_row_xml(
    cells: list[str],
    opts: TableOptions,
    header: bool,
    background: str | None = None,
    row_index: int = 0,
) -> str:
    cell_xml = []
    tr_pr = []
    if header and opts.header_repeat:
        tr_pr.append("<w:tblHeader/>")
    if header:
        if opts.header_row_height > 0 or opts.header_height_rule != RowHeightRule.AUTO:
            height = opts.header_row_height if opts.header_row_height > 0 else 360
            tr_pr.append(f'<w:trHeight w:val="{height}" w:hRule="{opts.header_height_rule.value}"/>')
    else:
        if opts.row_height > 0 or opts.row_height_rule != RowHeightRule.AUTO:
            height = opts.row_height if opts.row_height > 0 else 360
            tr_pr.append(f'<w:trHeight w:val="{height}" w:hRule="{opts.row_height_rule.value}"/>')

    tr_pr_xml = f"<w:trPr>{''.join(tr_pr)}</w:trPr>" if tr_pr else ""
    for idx, text in enumerate(cells):
        cell_key = (row_index, idx + 1) if row_index > 0 else None
        if header:
            align = opts.header_alignment
            if opts.columns[idx].alignment is not None:
                align = opts.columns[idx].alignment
            bold = opts.header_bold or opts.columns[idx].bold
            bg = opts.header_background
            style = _merge_text_style(opts.header_text_style, None)
        else:
            align = opts.row_alignment
            if opts.columns[idx].alignment is not None:
                align = opts.columns[idx].alignment
            bold = False
            bg = background or ""
            style = _merge_text_style(opts.row_text_style, None)

        if cell_key and opts.cell_text_styles and cell_key in opts.cell_text_styles:
            style = _merge_text_style(style, opts.cell_text_styles[cell_key])
        if opts.conditional_styles:
            condition = _match_conditional_style(text, opts.conditional_styles)
            if condition is not None:
                style = _merge_text_style(style, condition.text_style)
                if condition.background:
                    bg = condition.background

        runs = _build_run(text, bold=bold, style=style)

        cell_pr = [
            "<w:tcPr>",
            f'<w:vAlign w:val="{opts.vertical_alignment.value}"/>',
        ]
        if bg:
            cell_pr.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{xml_escape(bg)}"/>')
        cell_pr.append("</w:tcPr>")

        para_pr = "<w:pPr>"
        if opts.paragraph_style:
            para_pr += f'<w:pStyle w:val="{xml_escape(opts.paragraph_style)}"/>'
        para_pr += f'<w:jc w:val="{align.value}"/></w:pPr>'
        cell_xml.append(f"<w:tc>{''.join(cell_pr)}<w:p>{para_pr}{runs}</w:p></w:tc>")
    return "<w:tr>" + tr_pr_xml + "".join(cell_xml) + "</w:tr>"


def _table_borders_xml(opts: TableOptions) -> str:
    if opts.border_style == BorderStyle.NONE:
        return (
            "<w:tblBorders>"
            '<w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>'
            '<w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>'
            "</w:tblBorders>"
        )
    style = opts.border_style.value
    size = opts.border_size
    color = opts.border_color
    return (
        "<w:tblBorders>"
        f'<w:top w:val="{style}" w:sz="{size}" w:color="{color}"/>'
        f'<w:left w:val="{style}" w:sz="{size}" w:color="{color}"/>'
        f'<w:bottom w:val="{style}" w:sz="{size}" w:color="{color}"/>'
        f'<w:right w:val="{style}" w:sz="{size}" w:color="{color}"/>'
        f'<w:insideH w:val="{style}" w:sz="{size}" w:color="{color}"/>'
        f'<w:insideV w:val="{style}" w:sz="{size}" w:color="{color}"/>'
        "</w:tblBorders>"
    )


def _calculate_proportional_widths(opts: TableOptions) -> list[int]:
    total_width = _resolve_table_width(opts)
    lengths = []
    total_len = 0
    for i, col in enumerate(opts.columns):
        length = len(col.title)
        for row in opts.rows:
            if i < len(row):
                length = max(length, len(row[i]))
        lengths.append(length)
        total_len += length
    if total_len == 0:
        total_len = 1
    widths = []
    distributed = 0
    for length in lengths:
        w = (total_width * length) // total_len
        widths.append(w)
        distributed += w
    if widths and distributed < total_width:
        widths[-1] += total_width - distributed
    return widths


def _resolve_table_width(opts: TableOptions) -> int:
    if opts.table_width_type == TableWidthType.FIXED:
        return opts.table_width
    if opts.table_width_type == TableWidthType.PERCENTAGE:
        available = opts.available_width if opts.available_width > 0 else 9360
        return (available * opts.table_width) // 5000
    return 11520


def _match_conditional_style(text: str, styles: dict[str, ConditionalStyle]) -> ConditionalStyle | None:
    normalized = text.strip().lower()
    for key, style in styles.items():
        if normalized == key.strip().lower():
            return style
    return None


def _merge_text_style(base: CellTextStyle, override: CellTextStyle | None) -> CellTextStyle:
    if override is None:
        return base
    return CellTextStyle(
        bold=base.bold or override.bold,
        italic=base.italic or override.italic,
        font_size=override.font_size if override.font_size is not None else base.font_size,
        font_color=override.font_color if override.font_color is not None else base.font_color,
    )


def _build_run(text: str, bold: bool, style: CellTextStyle) -> str:
    r_pr = []
    if bold or style.bold:
        r_pr.append("<w:b/>")
    if style.italic:
        r_pr.append("<w:i/>")
    if style.font_size:
        r_pr.append(f'<w:sz w:val="{style.font_size}"/>')
        r_pr.append(f'<w:szCs w:val="{style.font_size}"/>')
    if style.font_color:
        r_pr.append(f'<w:color w:val="{xml_escape(style.font_color)}"/>')
    r_pr_xml = f"<w:rPr>{''.join(r_pr)}</w:rPr>" if r_pr else ""
    return f"<w:r>{r_pr_xml}<w:t>{xml_escape(text)}</w:t></w:r>"
