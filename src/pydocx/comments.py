from __future__ import annotations

import re
from dataclasses import replace
from datetime import UTC, datetime
from html import unescape as html_unescape
from pathlib import Path

from .options import Comment, CommentOptions
from .rels import ensure_content_type_override, insert_relationship, next_relationship_id
from .xmlutils import xml_escape

_COMMENT_ID_RE = re.compile(r'<w:comment[^>]*w:id="(\d+)"')
_COMMENT_AUTHOR_RE = re.compile(r'w:author="([^"]*)"')
_COMMENT_INITIALS_RE = re.compile(r'w:initials="([^"]*)"')
_COMMENT_DATE_RE = re.compile(r'w:date="([^"]*)"')
_COMMENT_BLOCK_RE = re.compile(r"(?s)<w:comment\s+[^>]*>.*?</w:comment>")
_COMMENT_TEXT_RE = re.compile(r"<w:t[^>]*>([^<]*)</w:t>")


def insert_comment(workspace: Path, opts: CommentOptions) -> None:
    if not opts.text:
        raise ValueError("comment text cannot be empty")
    if not opts.anchor:
        raise ValueError("anchor text cannot be empty")

    author = opts.author or "Author"
    initials = opts.initials or (author[0] if author else "")
    opts = replace(opts, author=author, initials=initials)

    comment_id = _ensure_comments_xml(workspace)
    _add_comment_content(workspace, comment_id, opts)
    _insert_comment_markers(workspace, opts.anchor, comment_id)


def get_comments(workspace: Path) -> list[Comment]:
    comments_path = workspace / "word" / "comments.xml"
    if not comments_path.exists():
        return []
    raw = comments_path.read_text(encoding="utf-8")
    return _parse_comments(raw)


def _ensure_comments_xml(workspace: Path) -> int:
    comments_path = workspace / "word" / "comments.xml"
    if not comments_path.exists():
        comments_path.write_text(_initial_comments_xml(), encoding="utf-8")
        _add_comments_relationship(workspace)
        _add_comments_content_type(workspace)
        return 1

    raw = comments_path.read_text(encoding="utf-8")
    return _next_comment_id(raw)


def _initial_comments_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        "</w:comments>"
    )


def _add_comments_relationship(workspace: Path) -> None:
    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    if "comments.xml" in rels_xml:
        return
    rel_id = next_relationship_id(rels_xml)
    rel = (
        f'<Relationship Id="{rel_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
        'Target="comments.xml"/>'
    )
    rels_path.write_text(insert_relationship(rels_xml, "\n  " + rel), encoding="utf-8")


def _add_comments_content_type(workspace: Path) -> None:
    ct_path = workspace / "[Content_Types].xml"
    content = ct_path.read_text(encoding="utf-8")
    updated = ensure_content_type_override(
        content,
        "/word/comments.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    )
    ct_path.write_text(updated, encoding="utf-8")


def _add_comment_content(workspace: Path, comment_id: int, opts: CommentOptions) -> None:
    comments_path = workspace / "word" / "comments.xml"
    raw = comments_path.read_text(encoding="utf-8")
    close_tag = "</w:comments>"
    idx = raw.rfind(close_tag)
    if idx == -1:
        raise ValueError("could not find closing </w:comments> tag")
    entry = _comment_entry(comment_id, opts)
    updated = raw[:idx] + entry + "\n" + raw[idx:]
    comments_path.write_text(updated, encoding="utf-8")


def _comment_entry(comment_id: int, opts: CommentOptions) -> str:
    date_str = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        f'<w:comment w:id="{comment_id}" w:author="{xml_escape(opts.author)}" '
        f'w:date="{date_str}" w:initials="{xml_escape(opts.initials)}">'
        "<w:p>"
        '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>'
        "<w:r>"
        '<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        "<w:annotationRef/>"
        "</w:r>"
        "<w:r>"
        f'<w:t xml:space="preserve"> {xml_escape(opts.text)}</w:t>'
        "</w:r>"
        "</w:p>"
        "</w:comment>"
    )


def _insert_comment_markers(workspace: Path, anchor: str, comment_id: int) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    para_start, para_end = _find_paragraph_range(doc_xml, anchor)
    if para_start == -1:
        raise ValueError(f"anchor text {anchor!r} not found in document")

    para_xml = doc_xml[para_start:para_end]
    ppr_end = para_xml.find("</w:pPr>")
    if ppr_end >= 0:
        insert_start_offset = ppr_end + len("</w:pPr>")
    else:
        p_open_end = para_xml.find(">")
        if p_open_end < 0:
            raise ValueError("invalid paragraph XML")
        insert_start_offset = p_open_end + 1

    insert_start_pos = para_start + insert_start_offset
    insert_end_pos = para_end - len("</w:p>")

    range_start = f'<w:commentRangeStart w:id="{comment_id}"/>'
    range_end = (
        f'<w:commentRangeEnd w:id="{comment_id}"/>'
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:commentReference w:id="{comment_id}"/></w:r>'
    )

    updated = (
        doc_xml[:insert_start_pos]
        + range_start
        + doc_xml[insert_start_pos:insert_end_pos]
        + range_end
        + doc_xml[insert_end_pos:]
    )
    doc_path.write_text(updated, encoding="utf-8")


def _next_comment_id(raw: str) -> int:
    max_id = 0
    for match in _COMMENT_ID_RE.findall(raw):
        try:
            max_id = max(max_id, int(match))
        except ValueError:
            continue
    return max_id + 1


def _parse_comments(raw: str) -> list[Comment]:
    comments: list[Comment] = []
    blocks = _COMMENT_BLOCK_RE.findall(raw)
    for block in blocks:
        comment_id = 0
        author = ""
        initials = ""
        date = ""

        m = _COMMENT_ID_RE.search(block)
        if m:
            try:
                comment_id = int(m.group(1))
            except ValueError:
                comment_id = 0
        m = _COMMENT_AUTHOR_RE.search(block)
        if m:
            author = m.group(1)
        m = _COMMENT_INITIALS_RE.search(block)
        if m:
            initials = m.group(1)
        m = _COMMENT_DATE_RE.search(block)
        if m:
            date = m.group(1)

        texts = []
        for tm in _COMMENT_TEXT_RE.findall(block):
            texts.append(html_unescape(tm))
        text = "".join(texts)

        if comment_id > 0:
            comments.append(
                Comment(
                    id=comment_id,
                    author=author,
                    initials=initials,
                    date=date,
                    text=text,
                )
            )
    return comments


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
