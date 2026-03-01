from __future__ import annotations

import re
from html import unescape as html_unescape
from pathlib import Path

from .document import insert_at_body_end, insert_at_body_start
from .options import BreakOptions, InsertPosition, PageLayoutOptions, PageOrientation, SectionBreakType


def insert_page_break(workspace: Path, opts: BreakOptions) -> None:
    page_break_xml = _page_break_xml()
    _insert_break(workspace, page_break_xml, opts)


def insert_section_break(workspace: Path, opts: BreakOptions) -> None:
    section_type = opts.section_type or SectionBreakType.NEXT_PAGE
    _validate_section_type(section_type)
    section_break_xml = _section_break_xml(section_type)
    _insert_break(workspace, section_break_xml, opts)
    if opts.page_layout is not None:
        set_page_layout(workspace, opts.page_layout)


def set_page_layout(workspace: Path, layout: PageLayoutOptions) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")

    sect_start = doc_xml.rfind("<w:sectPr")
    if sect_start == -1:
        body_end = doc_xml.rfind("</w:body>")
        if body_end == -1:
            raise ValueError("document body not found")
        sect_xml = _section_properties_xml(layout)
        updated = doc_xml[:body_end] + sect_xml + doc_xml[body_end:]
        doc_path.write_text(updated, encoding="utf-8")
        return

    sect_end = doc_xml.find("</w:sectPr>", sect_start)
    if sect_end == -1:
        raise ValueError("malformed section properties")
    sect_end += len("</w:sectPr>")
    sect_xml = _section_properties_xml(layout)
    updated = doc_xml[:sect_start] + sect_xml + doc_xml[sect_end:]
    doc_path.write_text(updated, encoding="utf-8")


def _page_break_xml() -> str:
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def _section_break_xml(break_type: SectionBreakType) -> str:
    return "<w:p><w:pPr>" + _section_type_xml(break_type) + "</w:pPr></w:p>"


def _section_properties_xml(layout: PageLayoutOptions, break_type: SectionBreakType | None = None) -> str:
    orient_attr = ""
    if layout.orientation == PageOrientation.LANDSCAPE:
        orient_attr = ' w:orient="landscape"'

    sect = ["<w:sectPr>"]
    if break_type is not None:
        sect.append(f'<w:type w:val="{break_type}"/>')
    sect.append(f'<w:pgSz w:w="{layout.page_width}" w:h="{layout.page_height}"{orient_attr}/>')
    sect.append(
        f'<w:pgMar w:top="{layout.margin_top}" w:right="{layout.margin_right}" '
        f'w:bottom="{layout.margin_bottom}" w:left="{layout.margin_left}" '
        f'w:header="{layout.margin_header}" w:footer="{layout.margin_footer}" '
        f'w:gutter="{layout.margin_gutter}"/>'
    )
    sect.append('<w:cols w:space="720"/>')
    sect.append("</w:sectPr>")
    return "".join(sect)


def _section_type_xml(break_type: SectionBreakType) -> str:
    return f'<w:sectPr><w:type w:val="{break_type}"/></w:sectPr>'


def _insert_break(workspace: Path, break_xml: str, opts: BreakOptions) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")

    if opts.position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, break_xml)
    elif opts.position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, break_xml)
    elif opts.position == InsertPosition.AFTER_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for after_text insertion")
        updated = _insert_after_anchor(doc_xml, break_xml, opts.anchor)
    elif opts.position == InsertPosition.BEFORE_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for before_text insertion")
        updated = _insert_before_anchor(doc_xml, break_xml, opts.anchor)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")

    doc_path.write_text(updated, encoding="utf-8")


def _validate_section_type(break_type: SectionBreakType) -> None:
    if break_type not in (
        SectionBreakType.NEXT_PAGE,
        SectionBreakType.CONTINUOUS,
        SectionBreakType.EVEN_PAGE,
        SectionBreakType.ODD_PAGE,
    ):
        raise ValueError(f"invalid section break type: {break_type}")


def _insert_after_anchor(doc_xml: str, fragment: str, anchor: str) -> str:
    start, end = _find_paragraph_range(doc_xml, anchor)
    if start == -1:
        raise ValueError(f"anchor text {anchor!r} not found in document")
    return doc_xml[:end] + fragment + doc_xml[end:]


def _insert_before_anchor(doc_xml: str, fragment: str, anchor: str) -> str:
    start, _end = _find_paragraph_range(doc_xml, anchor)
    if start == -1:
        raise ValueError(f"anchor text {anchor!r} not found in document")
    return doc_xml[:start] + fragment + doc_xml[start:]


def _find_paragraph_range(doc_xml: str, anchor: str) -> tuple[int, int]:
    pos = 0
    while True:
        start = doc_xml.find("<w:p", pos)
        if start == -1:
            return -1, -1
        end = doc_xml.find("</w:p>", start)
        if end == -1:
            return -1, -1
        end += len("</w:p>")
        para = doc_xml[start:end]
        text = _extract_paragraph_text(para)
        if anchor in text or _normalize_ws(anchor) in _normalize_ws(text):
            return start, end
        pos = end


def _extract_paragraph_text(para_xml: str) -> str:
    parts = []
    for match in re.finditer(r"<w:t[^>]*>(.*?)</w:t>", para_xml, flags=re.DOTALL):
        parts.append(html_unescape(match.group(1)))
    return "".join(parts)


def _normalize_ws(text: str) -> str:
    return " ".join(text.split())
