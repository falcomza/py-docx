from __future__ import annotations

import re
from pathlib import Path

from .document import insert_at_body_end, insert_at_body_start
from .options import BreakOptions, InsertPosition, PageLayoutOptions, PageOrientation, SectionBreakType
from .xmlops import insert_after_anchor, insert_before_anchor

_PGSZ_RE = re.compile(r"<w:pgSz\b[^>]*/>")
_PGMAR_RE = re.compile(r"<w:pgMar\b[^>]*/>")
_COLS_RE = re.compile(r"<w:cols\b[^>]*/>")


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


def set_page_layout(workspace: Path, layout: PageLayoutOptions, section_index: int = -1) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")

    matches = list(re.finditer(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc_xml, flags=re.DOTALL))
    if not matches:
        body_end = doc_xml.rfind("</w:body>")
        if body_end == -1:
            raise ValueError("document body not found")
        sect_xml = _section_properties_xml(layout)
        updated = doc_xml[:body_end] + sect_xml + doc_xml[body_end:]
        doc_path.write_text(updated, encoding="utf-8")
        return

    selected = _resolve_section_index(matches, section_index)
    sect_xml = _merge_section_layout(selected.group(0), layout)
    updated = doc_xml[: selected.start()] + sect_xml + doc_xml[selected.end() :]
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


def _merge_section_layout(sectpr_xml: str, layout: PageLayoutOptions) -> str:
    orient_attr = ""
    if layout.orientation == PageOrientation.LANDSCAPE:
        orient_attr = ' w:orient="landscape"'

    pg_sz = f'<w:pgSz w:w="{layout.page_width}" w:h="{layout.page_height}"{orient_attr}/>'
    pg_mar = (
        f'<w:pgMar w:top="{layout.margin_top}" w:right="{layout.margin_right}" '
        f'w:bottom="{layout.margin_bottom}" w:left="{layout.margin_left}" '
        f'w:header="{layout.margin_header}" w:footer="{layout.margin_footer}" '
        f'w:gutter="{layout.margin_gutter}"/>'
    )
    cols = '<w:cols w:space="720"/>'

    merged = _upsert_singleton_tag(sectpr_xml, _PGSZ_RE, pg_sz)
    merged = _upsert_singleton_tag(merged, _PGMAR_RE, pg_mar)
    merged = _upsert_singleton_tag(merged, _COLS_RE, cols)
    return merged


def _upsert_singleton_tag(sectpr_xml: str, pattern: re.Pattern[str], replacement: str) -> str:
    if pattern.search(sectpr_xml):
        return pattern.sub(replacement, sectpr_xml, count=1)
    return sectpr_xml.replace("</w:sectPr>", replacement + "</w:sectPr>", 1)


def _resolve_section_index(matches: list[re.Match[str]], section_index: int) -> re.Match[str]:
    if section_index < 0:
        return matches[-1]
    if section_index >= len(matches):
        raise ValueError(f"section_index out of range: {section_index}; document has {len(matches)} sections")
    return matches[section_index]


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
        updated = insert_after_anchor(doc_xml, break_xml, opts.anchor)
    elif opts.position == InsertPosition.BEFORE_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for before_text insertion")
        updated = insert_before_anchor(doc_xml, break_xml, opts.anchor)
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


