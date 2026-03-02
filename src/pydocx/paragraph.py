from __future__ import annotations

import re
from html import unescape as html_unescape
from pathlib import Path

from .document import insert_at_body_end, insert_at_body_start
from .lists import ensure_numbering_xml
from .options import InsertPosition, ListType, ParagraphOptions
from .xmlops import write_run_text
from .xmlutils import xml_escape


def insert_paragraph(workspace: Path, opts: ParagraphOptions) -> None:
    if not opts.text:
        raise ValueError("paragraph text cannot be empty")
    if opts.list_type is not None:
        ensure_numbering_xml(workspace)

    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    para_xml = _build_paragraph_xml(opts)

    if opts.position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, para_xml)
    elif opts.position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, para_xml)
    elif opts.position == InsertPosition.AFTER_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for after_text insertion")
        updated = _insert_after_anchor(doc_xml, para_xml, opts.anchor)
    elif opts.position == InsertPosition.BEFORE_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for before_text insertion")
        updated = _insert_before_anchor(doc_xml, para_xml, opts.anchor)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")

    doc_path.write_text(updated, encoding="utf-8")


def insert_paragraphs(workspace: Path, paragraphs: list[ParagraphOptions]) -> None:
    for idx, opts in enumerate(paragraphs):
        try:
            insert_paragraph(workspace, opts)
        except Exception as exc:
            raise ValueError(f"insert paragraph {idx} failed: {exc}") from exc


def add_heading(workspace: Path, level: int, text: str, position: InsertPosition) -> None:
    if level not in range(1, 10):
        raise ValueError("heading level must be between 1 and 9")
    style = f"Heading{level}"
    insert_paragraph(workspace, ParagraphOptions(
        text=text, style=style, position=position))


def add_text(workspace: Path, text: str, position: InsertPosition) -> None:
    insert_paragraph(workspace, ParagraphOptions(
        text=text, style="Normal", position=position))


def _build_paragraph_xml(opts: ParagraphOptions) -> str:
    p_pr = "<w:pPr>"
    if opts.style and opts.style != "Normal":
        p_pr += f'<w:pStyle w:val="{xml_escape(opts.style)}"/>'
    if opts.alignment:
        p_pr += f'<w:jc w:val="{opts.alignment.value}"/>'
    if opts.list_type is not None:
        level = max(0, min(opts.list_level, 8))
        num_id = "1" if opts.list_type == ListType.BULLET else "2"
        p_pr += "<w:numPr>"
        p_pr += f'<w:ilvl w:val="{level}"/>'
        p_pr += f'<w:numId w:val="{num_id}"/>'
        p_pr += "</w:numPr>"
    p_pr += "</w:pPr>"

    r_pr = []
    if opts.bold:
        r_pr.append("<w:b/>")
    if opts.italic:
        r_pr.append("<w:i/>")
    if opts.underline:
        r_pr.append('<w:u w:val="single"/>')
    r_pr_xml = f"<w:rPr>{''.join(r_pr)}</w:rPr>" if r_pr else ""

    run_xml = "<w:r>" + r_pr_xml + write_run_text(opts.text) + "</w:r>"
    return "<w:p>" + p_pr + run_xml + "</w:p>"


def _insert_after_anchor(doc_xml: str, para_xml: str, anchor: str) -> str:
    start, end = _find_paragraph_range(doc_xml, anchor)
    if start == -1:
        raise ValueError(f"anchor text {anchor!r} not found in document")
    insert_pos = end
    return doc_xml[:insert_pos] + para_xml + doc_xml[insert_pos:]


def _insert_before_anchor(doc_xml: str, para_xml: str, anchor: str) -> str:
    start, _end = _find_paragraph_range(doc_xml, anchor)
    if start == -1:
        raise ValueError(f"anchor text {anchor!r} not found in document")
    insert_pos = start
    return doc_xml[:insert_pos] + para_xml + doc_xml[insert_pos:]


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
