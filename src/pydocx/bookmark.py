from __future__ import annotations

import re
from html import unescape as html_unescape
from pathlib import Path

from .document import insert_at_body_end, insert_at_body_start
from .options import BookmarkOptions, InsertPosition
from .xmlutils import xml_escape

_VALID_BOOKMARK_RE = re.compile(r"^[a-zA-Z][a-zA-Z0-9_]*$")
_BOOKMARK_ID_RE = re.compile(r"<w:bookmark(?:Start|End)[^>]*w:id=\"(\d+)\"")
_RESERVED_BOOKMARK_PREFIXES = ("_Toc", "_Hlt", "_Ref", "_GoBack")


def create_bookmark(workspace: Path, name: str, opts: BookmarkOptions) -> None:
    _validate_bookmark_name(name)
    bookmark_id = _next_bookmark_id(workspace)
    bookmark_xml = _bookmark_marker_xml(name, bookmark_id, opts)
    _insert_bookmark(workspace, bookmark_xml, opts)


def create_bookmark_with_text(workspace: Path, name: str, text: str, opts: BookmarkOptions) -> None:
    if not text:
        raise ValueError("bookmark text cannot be empty")
    _validate_bookmark_name(name)
    bookmark_id = _next_bookmark_id(workspace)
    bookmark_xml = _bookmark_with_text_xml(name, text, bookmark_id, opts)
    _insert_bookmark(workspace, bookmark_xml, opts)


def wrap_text_in_bookmark(workspace: Path, name: str, anchor_text: str) -> None:
    if not anchor_text:
        raise ValueError("anchor text cannot be empty")
    _validate_bookmark_name(name)
    bookmark_id = _next_bookmark_id(workspace)

    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated = _wrap_existing_text_in_bookmark(doc_xml, name, anchor_text, bookmark_id)
    doc_path.write_text(updated, encoding="utf-8")


def _validate_bookmark_name(name: str) -> None:
    if not name:
        raise ValueError("bookmark name cannot be empty")
    if len(name) > 40:
        raise ValueError("bookmark name must be 40 characters or less")
    if not name[0].isalpha():
        raise ValueError("bookmark name must start with a letter")
    if not _VALID_BOOKMARK_RE.match(name):
        raise ValueError("bookmark name can only contain letters, digits, and underscores")
    for prefix in _RESERVED_BOOKMARK_PREFIXES:
        if name.startswith(prefix):
            raise ValueError(f"bookmark name cannot start with reserved prefix '{prefix}'")


def _next_bookmark_id(workspace: Path) -> int:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    max_id = 0
    for match in _BOOKMARK_ID_RE.findall(doc_xml):
        try:
            max_id = max(max_id, int(match))
        except ValueError:
            continue
    return max_id + 1


def _bookmark_marker_xml(name: str, bookmark_id: int, opts: BookmarkOptions) -> str:
    p_pr = ""
    if opts.style and opts.style != "Normal":
        p_pr = f'<w:pPr><w:pStyle w:val="{xml_escape(opts.style)}"/></w:pPr>'
    return (
        "<w:p>"
        f"{p_pr}"
        f'<w:bookmarkStart w:id="{bookmark_id}" w:name="{xml_escape(name)}"/>'
        f'<w:bookmarkEnd w:id="{bookmark_id}"/>'
        "</w:p>"
    )


def _bookmark_with_text_xml(name: str, text: str, bookmark_id: int, opts: BookmarkOptions) -> str:
    p_pr = ""
    if opts.style and opts.style != "Normal":
        p_pr = f'<w:pPr><w:pStyle w:val="{xml_escape(opts.style)}"/></w:pPr>'
    run = _write_run_text(text)
    return (
        "<w:p>"
        f"{p_pr}"
        f'<w:bookmarkStart w:id="{bookmark_id}" w:name="{xml_escape(name)}"/>'
        f"<w:r>{run}</w:r>"
        f'<w:bookmarkEnd w:id="{bookmark_id}"/>'
        "</w:p>"
    )


def _insert_bookmark(workspace: Path, bookmark_xml: str, opts: BookmarkOptions) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")

    if opts.position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, bookmark_xml)
    elif opts.position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, bookmark_xml)
    elif opts.position == InsertPosition.AFTER_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for after_text insertion")
        updated = _insert_after_anchor(doc_xml, bookmark_xml, opts.anchor)
    elif opts.position == InsertPosition.BEFORE_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for before_text insertion")
        updated = _insert_before_anchor(doc_xml, bookmark_xml, opts.anchor)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")

    doc_path.write_text(updated, encoding="utf-8")


def _wrap_existing_text_in_bookmark(doc_xml: str, name: str, anchor_text: str, bookmark_id: int) -> str:
    escaped_anchor = xml_escape(anchor_text)
    text_idx = doc_xml.find(escaped_anchor)
    if text_idx == -1:
        raise ValueError(f"anchor text not found in document: {anchor_text}")

    run_start_idx = doc_xml.rfind("<w:r", 0, text_idx)
    if run_start_idx == -1:
        raise ValueError("could not find run tag before anchor text")

    run_end_idx = doc_xml.find("</w:r>", text_idx)
    if run_end_idx == -1:
        raise ValueError("could not find closing run tag after anchor text")
    run_end_idx += len("</w:r>")

    bookmark_start = f'<w:bookmarkStart w:id="{bookmark_id}" w:name="{xml_escape(name)}"/>'
    bookmark_end = f'<w:bookmarkEnd w:id="{bookmark_id}"/>'

    return (
        doc_xml[:run_start_idx]
        + bookmark_start
        + doc_xml[run_start_idx:run_end_idx]
        + bookmark_end
        + doc_xml[run_end_idx:]
    )


def _write_run_text(text: str) -> str:
    parts: list[str] = []
    start = 0

    def flush(seg: str) -> None:
        if seg == "":
            return
        t = "<w:t"
        if seg.startswith(" ") or seg.endswith(" "):
            t += ' xml:space="preserve"'
        t += ">" + xml_escape(seg) + "</w:t>"
        parts.append(t)

    for i, ch in enumerate(text):
        if ch == "\n":
            flush(text[start:i])
            parts.append("<w:br/>")
            start = i + 1
        elif ch == "\t":
            flush(text[start:i])
            parts.append("<w:tab/>")
            start = i + 1
    flush(text[start:])
    return "".join(parts)


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
