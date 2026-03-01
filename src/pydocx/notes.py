from __future__ import annotations

import re
from html import unescape as html_unescape
from pathlib import Path

from .options import EndnoteOptions, FootnoteOptions
from .rels import insert_relationship, next_relationship_id
from .xmlutils import xml_escape


def insert_footnote(workspace: Path, opts: FootnoteOptions) -> None:
    if not opts.text:
        raise ValueError("footnote text cannot be empty")
    if not opts.anchor:
        raise ValueError("anchor text cannot be empty")
    note_id = _ensure_notes_xml(workspace, "footnotes")
    _add_note_content(workspace, "footnotes", note_id, opts.text)
    _insert_note_reference(workspace, opts.anchor, note_id, "footnote")


def insert_endnote(workspace: Path, opts: EndnoteOptions) -> None:
    if not opts.text:
        raise ValueError("endnote text cannot be empty")
    if not opts.anchor:
        raise ValueError("anchor text cannot be empty")
    note_id = _ensure_notes_xml(workspace, "endnotes")
    _add_note_content(workspace, "endnotes", note_id, opts.text)
    _insert_note_reference(workspace, opts.anchor, note_id, "endnote")


def _ensure_notes_xml(workspace: Path, kind: str) -> int:
    filename = "footnotes.xml" if kind == "footnotes" else "endnotes.xml"
    path = workspace / "word" / filename
    if not path.exists():
        content = _initial_notes_xml(kind)
        path.write_text(content, encoding="utf-8")
        _add_note_relationship(workspace, filename, kind)
        _add_note_content_type(workspace, filename, kind)
        return 1
    raw = path.read_text(encoding="utf-8")
    return _next_note_id(raw, "footnote" if kind == "footnotes" else "endnote")


def _initial_notes_xml(kind: str) -> str:
    root = "w:footnotes" if kind == "footnotes" else "w:endnotes"
    note = "w:footnote" if kind == "footnotes" else "w:endnote"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        f'<{note} w:type="separator" w:id="-1">'
        '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        f"<w:r><w:separator/></w:r></w:p></{note}>\n"
        f'<{note} w:type="continuationSeparator" w:id="0">'
        '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        f"<w:r><w:continuationSeparator/></w:r></w:p></{note}>\n"
        f"</{root}>"
    )


def _add_note_content(workspace: Path, kind: str, note_id: int, text: str) -> None:
    filename = "footnotes.xml" if kind == "footnotes" else "endnotes.xml"
    path = workspace / "word" / filename
    raw = path.read_text(encoding="utf-8")
    close_tag = "</w:footnotes>" if kind == "footnotes" else "</w:endnotes>"
    idx = raw.rfind(close_tag)
    if idx == -1:
        raise ValueError(f"could not find closing tag for {kind}")
    entry = _note_entry(kind, note_id, text)
    updated = raw[:idx] + entry + "\n" + raw[idx:]
    path.write_text(updated, encoding="utf-8")


def _note_entry(kind: str, note_id: int, text: str) -> str:
    note = "footnote" if kind == "footnotes" else "endnote"
    style = "FootnoteText" if kind == "footnotes" else "EndnoteText"
    ref_style = "FootnoteReference" if kind == "footnotes" else "EndnoteReference"
    ref_tag = "footnoteRef" if kind == "footnotes" else "endnoteRef"
    return (
        f'<w:{note} w:id="{note_id}">'
        "<w:p>"
        f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
        "<w:r>"
        f'<w:rPr><w:rStyle w:val="{ref_style}"/></w:rPr>'
        f"<w:{ref_tag}/>"
        "</w:r>"
        "<w:r>"
        f'<w:t xml:space="preserve"> {xml_escape(text)}</w:t>'
        "</w:r>"
        "</w:p>"
        f"</w:{note}>"
    )


def _insert_note_reference(workspace: Path, anchor: str, note_id: int, note_type: str) -> None:
    doc_path = workspace / "word" / "document.xml"
    raw = doc_path.read_text(encoding="utf-8")
    start, end = _find_paragraph_range(raw, anchor)
    if start == -1:
        raise ValueError(f"anchor text {anchor!r} not found in document")
    insert_pos = end - len("</w:p>")
    if note_type == "footnote":
        ref = f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{note_id}"/></w:r>'
    else:
        ref = f'<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="{note_id}"/></w:r>'
    updated = raw[:insert_pos] + ref + raw[insert_pos:]
    doc_path.write_text(updated, encoding="utf-8")


def _find_paragraph_range(doc_xml: str, anchor: str) -> tuple[int, int]:
    pos = 0
    while True:
        start = _find_next_paragraph_start(doc_xml, pos)
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


def _find_next_paragraph_start(doc_xml: str, start: int) -> int:
    idx = doc_xml.find("<w:p", start)
    if idx == -1:
        return -1
    return idx


def _extract_paragraph_text(para_xml: str) -> str:
    out = []
    for match in re.finditer(r"<w:t[^>]*>(.*?)</w:t>", para_xml, flags=re.DOTALL):
        out.append(html_unescape(match.group(1)))
    return "".join(out)


def _normalize_ws(text: str) -> str:
    return " ".join(text.split())


def _next_note_id(raw: str, note_tag: str) -> int:
    matches = re.findall(rf"<w:{note_tag}[^>]*w:id=\"(\d+)\"", raw)
    max_id = 0
    for m in matches:
        try:
            max_id = max(max_id, int(m))
        except ValueError:
            continue
    return max_id + 1


def _add_note_relationship(workspace: Path, filename: str, rel_type: str) -> None:
    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    if filename in rels_xml:
        return
    rel_id = next_relationship_id(rels_xml)
    rel = (
        f'<Relationship Id="{rel_id}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rel_type}" '
        f'Target="{filename}"/>'
    )
    rels_path.write_text(insert_relationship(rels_xml, "\n  " + rel), encoding="utf-8")


def _add_note_content_type(workspace: Path, filename: str, note_type: str) -> None:
    ct_path = workspace / "[Content_Types].xml"
    content = ct_path.read_text(encoding="utf-8")
    if filename in content:
        return
    ctype = f"application/vnd.openxmlformats-officedocument.wordprocessingml.{note_type}+xml"
    override = f'<Override PartName="/word/{filename}" ContentType="{ctype}"/>'
    content = content.replace("</Types>", override + "</Types>", 1)
    ct_path.write_text(content, encoding="utf-8")
