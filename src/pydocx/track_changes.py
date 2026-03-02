from __future__ import annotations

import re
from datetime import UTC, datetime
from html import unescape as html_unescape
from pathlib import Path

from .document import insert_at_body_end, insert_at_body_start
from .options import InsertPosition, TrackedDeleteOptions, TrackedInsertOptions
from .xmlops import write_run_text
from .xmlutils import xml_escape

_REVISION_ID_RE = re.compile(r'w:id="(\d+)"')


def insert_tracked_text(workspace: Path, opts: TrackedInsertOptions) -> None:
    if not opts.text:
        raise ValueError("text cannot be empty")

    author = opts.author or "Author"
    date = opts.date or datetime.now(UTC)

    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    start_id = _next_revision_id(doc_xml)
    tracked_xml = _build_tracked_insert_xml(opts, author, date, start_id)
    updated = _insert_tracked_at_position(doc_xml, tracked_xml, opts)
    doc_path.write_text(updated, encoding="utf-8")


def delete_tracked_text(workspace: Path, opts: TrackedDeleteOptions) -> None:
    if not opts.anchor:
        raise ValueError("anchor text cannot be empty")

    author = opts.author or "Author"
    date = opts.date or datetime.now(UTC)

    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    start_id = _next_revision_id(doc_xml)
    updated = _mark_paragraph_as_deleted(doc_xml, opts.anchor, author, date, start_id)
    doc_path.write_text(updated, encoding="utf-8")


def _next_revision_id(doc_xml: str) -> int:
    max_id = 0
    for match in _REVISION_ID_RE.findall(doc_xml):
        try:
            max_id = max(max_id, int(match))
        except ValueError:
            continue
    return max_id + 1


def _build_tracked_insert_xml(
    opts: TrackedInsertOptions,
    author: str,
    date: datetime,
    start_id: int,
) -> str:
    date_str = date.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")

    p_pr = "<w:pPr>"
    if opts.style and opts.style != "Normal":
        p_pr += f'<w:pStyle w:val="{xml_escape(opts.style)}"/>'
    p_pr += f'<w:rPr><w:ins w:id="{start_id}" w:author="{xml_escape(author)}" w:date="{date_str}"/></w:rPr>'
    p_pr += "</w:pPr>"

    r_pr = []
    if opts.bold:
        r_pr.append("<w:b/>")
    if opts.italic:
        r_pr.append("<w:i/>")
    if opts.underline:
        r_pr.append('<w:u w:val="single"/>')
    r_pr_xml = f"<w:rPr>{''.join(r_pr)}</w:rPr>" if r_pr else ""

    run_xml = f"<w:r>{r_pr_xml}{write_run_text(opts.text)}</w:r>"

    return (
        "<w:p>"
        f"{p_pr}"
        f'<w:ins w:id="{start_id + 1}" w:author="{xml_escape(author)}" w:date="{date_str}">'
        f"{run_xml}"
        "</w:ins>"
        "</w:p>"
    )


def _insert_tracked_at_position(
    doc_xml: str,
    tracked_xml: str,
    opts: TrackedInsertOptions,
) -> str:
    if opts.position == InsertPosition.BEGINNING:
        return insert_at_body_start(doc_xml, tracked_xml)
    if opts.position == InsertPosition.END:
        return insert_at_body_end(doc_xml, tracked_xml)
    if opts.position == InsertPosition.AFTER_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for after_text insertion")
        return _insert_after_anchor(doc_xml, tracked_xml, opts.anchor)
    if opts.position == InsertPosition.BEFORE_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for before_text insertion")
        return _insert_before_anchor(doc_xml, tracked_xml, opts.anchor)
    raise ValueError(f"unsupported insert position: {opts.position}")


def _mark_paragraph_as_deleted(
    doc_xml: str,
    anchor: str,
    author: str,
    date: datetime,
    start_id: int,
) -> str:
    para_start, para_end = _find_paragraph_range(doc_xml, anchor)
    if para_start == -1:
        raise ValueError(f"anchor text {anchor!r} not found in document")
    para_xml = doc_xml[para_start:para_end]
    date_str = date.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    modified = _convert_runs_to_deleted_with_id(para_xml, author, date_str, start_id)
    return doc_xml[:para_start] + modified + doc_xml[para_end:]


def _convert_runs_to_deleted_with_id(
    para_xml: str,
    author: str,
    date_str: str,
    start_id: int,
) -> str:
    result: list[str] = []
    del_id = start_id
    pos = 0
    while pos < len(para_xml):
        run_start = para_xml.find("<w:r>", pos)
        run_start_attr = para_xml.find("<w:r ", pos)
        next_run = -1
        if run_start >= 0:
            next_run = run_start
        if run_start_attr >= 0 and (next_run < 0 or run_start_attr < next_run):
            next_run = run_start_attr
        if next_run < 0:
            result.append(para_xml[pos:])
            break
        result.append(para_xml[pos:next_run])

        run_end = para_xml.find("</w:r>", next_run)
        if run_end < 0:
            result.append(para_xml[next_run:])
            break
        run_end += len("</w:r>")
        run_content = para_xml[next_run:run_end]

        if "<w:t" in run_content:
            del_run = run_content
            del_run = del_run.replace("<w:t>", '<w:delText xml:space="preserve">')
            del_run = del_run.replace("<w:t ", "<w:delText ")
            del_run = del_run.replace("</w:t>", "</w:delText>")
            result.append(f'<w:del w:id="{del_id}" w:author="{xml_escape(author)}" w:date="{date_str}">')
            result.append(del_run)
            result.append("</w:del>")
            del_id += 1
        else:
            result.append(run_content)

        pos = run_end
    return "".join(result)


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
