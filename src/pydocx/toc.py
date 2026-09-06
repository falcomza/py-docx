from __future__ import annotations

import re
from pathlib import Path

from .document import insert_at_body_end, insert_at_body_start
from .options import CaptionListOptions, InsertPosition, TOCEntry, TOCOptions
from .settings import ensure_update_fields, mark_fields_dirty
from .xmlops import extract_paragraph_text
from .xmlutils import xml_escape

_PARA_RE = re.compile(r"(?s)<w:p(?:\s[^>]*)?>.*?</w:p>")
_TOC_STYLE_RE = re.compile(r'<w:pStyle w:val="(?:TOC|toc)(\d+)"')


def insert_toc(workspace: Path, opts: TOCOptions) -> None:
    instr = f'TOC \\o "{xml_escape(opts.outline_levels)}" \\h \\z \\u'
    _insert_field(workspace, opts.title, instr, opts.position, opts.update_on_open)


def insert_caption_list(workspace: Path, opts: CaptionListOptions) -> None:
    """Insert a Table of Figures / Table of Tables field (label from ``opts``)."""
    instr = f'TOC \\h \\z \\c "{xml_escape(opts.caption_label or "Figure")}"'
    _insert_field(workspace, opts.title, instr, opts.position, opts.update_on_open)


# Kept as distinct public names for the Updater API; behaviour is label-driven.
insert_table_of_figures = insert_caption_list
insert_table_of_tables = insert_caption_list


def get_toc_entries(workspace: Path) -> list[TOCEntry]:
    doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
    entries: list[TOCEntry] = []
    for match in _PARA_RE.finditer(doc_xml):
        para = match.group(0)
        style = _TOC_STYLE_RE.search(para)
        if not style:
            continue
        text = extract_paragraph_text(para).strip()
        if text:
            entries.append(TOCEntry(level=int(style.group(1)), text=text))
    return entries


def update_toc(workspace: Path) -> None:
    """Mark TOC/caption-list fields dirty so Word recalculates them on open."""
    ensure_update_fields(workspace)
    mark_fields_dirty(workspace / "word" / "document.xml")


def _insert_field(
    workspace: Path,
    title: str,
    instr: str,
    position: InsertPosition,
    update_on_open: bool,
) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")

    parts: list[str] = []
    if title:
        parts.append(
            f'<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>{xml_escape(title)}</w:t></w:r></w:p>'
        )
    parts.append(
        "<w:p>"
        '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve">{instr}</w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t> </w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "</w:p>"
    )
    field_xml = "".join(parts)

    if position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, field_xml)
    elif position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, field_xml)
    else:
        raise ValueError(f"unsupported insert position: {position}")
    doc_path.write_text(updated, encoding="utf-8")

    if update_on_open:
        ensure_update_fields(workspace)
