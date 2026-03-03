from __future__ import annotations

from pathlib import Path

from .document import insert_at_body_end, insert_at_body_start
from .lists import ensure_numbering_xml
from .options import InsertPosition, ListType, ParagraphOptions
from .xmlops import build_rpr_xml, insert_after_anchor, insert_before_anchor, write_run_text
from .xmlutils import xml_escape


def _apply_paragraph(doc_xml: str, opts: ParagraphOptions) -> str:
    para_xml = _build_paragraph_xml(opts)

    if opts.position == InsertPosition.BEGINNING:
        return insert_at_body_start(doc_xml, para_xml)
    elif opts.position == InsertPosition.END:
        return insert_at_body_end(doc_xml, para_xml)
    elif opts.position == InsertPosition.AFTER_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for after_text insertion")
        return insert_after_anchor(doc_xml, para_xml, opts.anchor)
    elif opts.position == InsertPosition.BEFORE_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for before_text insertion")
        return insert_before_anchor(doc_xml, para_xml, opts.anchor)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")


def insert_paragraph(workspace: Path, opts: ParagraphOptions) -> None:
    if not opts.text:
        raise ValueError("paragraph text cannot be empty")
    if opts.list_type is not None:
        ensure_numbering_xml(workspace)

    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated = _apply_paragraph(doc_xml, opts)
    doc_path.write_text(updated, encoding="utf-8")


def insert_paragraphs(workspace: Path, paragraphs: list[ParagraphOptions]) -> None:
    if not paragraphs:
        return

    # Ensure numbering once for any list items
    if any(opts.list_type is not None for opts in paragraphs):
        ensure_numbering_xml(workspace)

    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    for idx, opts in enumerate(paragraphs):
        if not opts.text:
            raise ValueError(f"insert paragraph {idx} failed: paragraph text cannot be empty")
        try:
            doc_xml = _apply_paragraph(doc_xml, opts)
        except Exception as exc:
            raise ValueError(f"insert paragraph {idx} failed: {exc}") from exc
    doc_path.write_text(doc_xml, encoding="utf-8")


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
        if opts.restart:
            p_pr += '<w:numRestart w:val="0"/>'
        p_pr += f'<w:ilvl w:val="{level}"/>'
        p_pr += f'<w:numId w:val="{num_id}"/>'
        p_pr += "</w:numPr>"
    p_pr += "</w:pPr>"

    r_pr_xml = build_rpr_xml(opts.bold, opts.italic, opts.underline)
    run_xml = "<w:r>" + r_pr_xml + write_run_text(opts.text) + "</w:r>"
    return "<w:p>" + p_pr + run_xml + "</w:p>"


