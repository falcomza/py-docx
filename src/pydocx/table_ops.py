from __future__ import annotations

from pathlib import Path

from .xmlops import find_nth_xml_block, inject_tcpr_element, replace_cell_text


def update_table_cell(workspace: Path, table_index: int, row: int, col: int, value: str) -> None:
    if table_index < 1:
        raise ValueError("table_index must be >= 1")
    if row < 1:
        raise ValueError("row must be >= 1")
    if col < 1:
        raise ValueError("col must be >= 1")

    doc_path = workspace / "word" / "document.xml"
    content = doc_path.read_text(encoding="utf-8")

    tbl_start, tbl_end = find_nth_xml_block(content, "w:tbl", table_index)
    tbl_content = content[tbl_start:tbl_end]

    tr_start, tr_end = find_nth_xml_block(tbl_content, "w:tr", row)
    tr_content = tbl_content[tr_start:tr_end]

    tc_start, tc_end = find_nth_xml_block(tr_content, "w:tc", col)
    tc_content = tr_content[tc_start:tc_end]

    updated_tc = replace_cell_text(tc_content, value)

    new_tr = tr_content[:tc_start] + updated_tc + tr_content[tc_end:]
    new_tbl = tbl_content[:tr_start] + new_tr + tbl_content[tr_end:]
    result = content[:tbl_start] + new_tbl + content[tbl_end:]

    doc_path.write_text(result, encoding="utf-8")


def merge_table_cells_horizontal(workspace: Path, table_index: int, row: int, start_col: int, end_col: int) -> None:
    if table_index < 1:
        raise ValueError("table_index must be >= 1")
    if row < 1:
        raise ValueError("row must be >= 1")
    if start_col < 1:
        raise ValueError("start_col must be >= 1")
    if end_col <= start_col:
        raise ValueError("end_col must be greater than start_col")

    doc_path = workspace / "word" / "document.xml"
    content = doc_path.read_text(encoding="utf-8")

    span = end_col - start_col + 1

    tbl_start, tbl_end = find_nth_xml_block(content, "w:tbl", table_index)
    tbl_content = content[tbl_start:tbl_end]

    tr_start, tr_end = find_nth_xml_block(tbl_content, "w:tr", row)
    tr_content = tbl_content[tr_start:tr_end]

    tc_start_first, tc_end_first = find_nth_xml_block(tr_content, "w:tc", start_col)
    first_cell = tr_content[tc_start_first:tc_end_first]

    _, tc_end_last = find_nth_xml_block(tr_content, "w:tc", end_col)

    modified_first = inject_tcpr_element(first_cell, f'<w:gridSpan w:val="{span}"/>')

    new_tr = tr_content[:tc_start_first] + modified_first + tr_content[tc_end_last:]
    new_tbl = tbl_content[:tr_start] + new_tr + tbl_content[tr_end:]
    result = content[:tbl_start] + new_tbl + content[tbl_end:]

    doc_path.write_text(result, encoding="utf-8")


def merge_table_cells_vertical(workspace: Path, table_index: int, start_row: int, end_row: int, col: int) -> None:
    if table_index < 1:
        raise ValueError("table_index must be >= 1")
    if col < 1:
        raise ValueError("col must be >= 1")
    if start_row < 1:
        raise ValueError("start_row must be >= 1")
    if end_row <= start_row:
        raise ValueError("end_row must be greater than start_row")

    doc_path = workspace / "word" / "document.xml"
    result = doc_path.read_text(encoding="utf-8")

    for row in range(start_row, end_row + 1):
        tbl_start, tbl_end = find_nth_xml_block(result, "w:tbl", table_index)
        tbl_content = result[tbl_start:tbl_end]

        tr_start, tr_end = find_nth_xml_block(tbl_content, "w:tr", row)
        tr_content = tbl_content[tr_start:tr_end]

        tc_start, tc_end = find_nth_xml_block(tr_content, "w:tc", col)
        cell_content = tr_content[tc_start:tc_end]

        merge_element = '<w:vMerge w:val="restart"/>' if row == start_row else "<w:vMerge/>"
        modified_cell = inject_tcpr_element(cell_content, merge_element)

        new_tr = tr_content[:tc_start] + modified_cell + tr_content[tc_end:]
        new_tbl = tbl_content[:tr_start] + new_tr + tbl_content[tr_end:]
        result = result[:tbl_start] + new_tbl + result[tbl_end:]

    doc_path.write_text(result, encoding="utf-8")
