from __future__ import annotations

import re
from collections.abc import Callable
from html import unescape as html_unescape
from pathlib import Path

from .options import DeleteMatchMode, DeleteOptions

_PARA_PATTERN = re.compile(r"(?s)<w:p[^>]*>.*?</w:p>")
_TABLE_PATTERN = re.compile(r"(?s)<w:tbl>.*?</w:tbl>")
_IMAGE_PATTERN = re.compile(r'(?s)<w:p[^>]*>.*?<wp:inline.*?<a:blip[^>]*r:embed="[^"]*"[^>]*>.*?</wp:inline>.*?</w:p>')
_CHART_PATTERN = re.compile(r'(?s)<w:p[^>]*>.*?<wp:inline.*?<c:chart[^>]*r:id="[^"]*"[^>]*>.*?</wp:inline>.*?</w:p>')
_TEXT_TOKEN_PATTERN = re.compile(r"(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>|<w:tab[^>]*/>|<w:br[^>]*/>")


def delete_paragraphs(workspace: Path, text: str, opts: DeleteOptions) -> int:
    if not text:
        raise ValueError("text cannot be empty")
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated, count = _delete_paragraphs_containing(doc_xml, text, opts)
    if count > 0:
        doc_path.write_text(updated, encoding="utf-8")
    return count


def delete_table(workspace: Path, table_index: int) -> None:
    if table_index < 1:
        raise ValueError("table index must be >= 1")
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated = _delete_nth_pattern(doc_xml, _TABLE_PATTERN, table_index, "table")
    doc_path.write_text(updated, encoding="utf-8")


def delete_image(workspace: Path, image_index: int) -> None:
    if image_index < 1:
        raise ValueError("image index must be >= 1")
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated = _delete_nth_pattern(doc_xml, _IMAGE_PATTERN, image_index, "image")
    doc_path.write_text(updated, encoding="utf-8")


def delete_chart(workspace: Path, chart_index: int) -> None:
    if chart_index < 1:
        raise ValueError("chart index must be >= 1")
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated = _delete_nth_pattern(doc_xml, _CHART_PATTERN, chart_index, "chart")
    doc_path.write_text(updated, encoding="utf-8")


def _delete_paragraphs_containing(doc_xml: str, text: str, opts: DeleteOptions) -> tuple[str, int]:
    matcher = _build_matcher(text, opts)
    count = 0
    result: list[str] = []
    last_end = 0
    for match in _PARA_PATTERN.finditer(doc_xml):
        para_xml = match.group(0)
        para_text = _extract_paragraph_plain_text(para_xml)
        if (opts.max_deletions == 0 or count < opts.max_deletions) and matcher(para_text):
            result.append(doc_xml[last_end : match.start()])
            last_end = match.end()
            count += 1
            continue
        result.append(doc_xml[last_end : match.end()])
        last_end = match.end()
    result.append(doc_xml[last_end:])
    return "".join(result), count


def _delete_nth_pattern(doc_xml: str, pattern: re.Pattern[str], index: int, label: str) -> str:
    matches = list(pattern.finditer(doc_xml))
    if index > len(matches):
        raise ValueError(f"{label} {index} not found (document has {len(matches)} {label}s)")
    target = matches[index - 1]
    return doc_xml[: target.start()] + doc_xml[target.end() :]


def _build_matcher(text: str, opts: DeleteOptions) -> Callable[[str], bool]:
    flags = 0 if opts.match_case else re.IGNORECASE

    if opts.mode == DeleteMatchMode.EXACT and not opts.whole_word:
        pattern = re.compile(rf"\A{re.escape(text)}\Z", flags)
        return lambda para: pattern.search(para.strip()) is not None

    if opts.mode == DeleteMatchMode.REGEX:
        body = text
    else:
        body = re.escape(text)
        if opts.whole_word:
            body = rf"\b{body}\b"
    pattern = re.compile(body, flags)
    return lambda para: pattern.search(para) is not None


def _extract_paragraph_plain_text(para_xml: str) -> str:
    parts: list[str] = []
    for match in _TEXT_TOKEN_PATTERN.finditer(para_xml):
        if match.group(1) is not None:
            parts.append(html_unescape(match.group(1)))
        else:
            parts.append(" ")
    return "".join(parts)
