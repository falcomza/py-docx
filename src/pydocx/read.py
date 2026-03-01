from __future__ import annotations

import re
from collections.abc import Iterable
from html import unescape as html_unescape
from pathlib import Path

from .options import FindOptions, TextMatch

_TEXT_PATTERN = re.compile(r"<w:t(?:[ \t][^>]*)?>([^<]*)</w:t>")
_PARA_PATTERN = re.compile(r"(?s)<w:p[^>]*>.*?</w:p>")
_TABLE_PATTERN = re.compile(r"(?s)<w:tbl>.*?</w:tbl>")
_ROW_PATTERN = re.compile(r"(?s)<w:tr[^>]*>.*?</w:tr>")
_CELL_PATTERN = re.compile(r"(?s)<w:tc>.*?</w:tc>")


def get_text(workspace: Path) -> str:
    doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
    return _extract_text_from_xml(doc_xml)


def get_paragraph_text(workspace: Path) -> list[str]:
    doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
    return _extract_paragraphs_from_xml(doc_xml)


def get_table_text(workspace: Path) -> list[list[list[str]]]:
    doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
    return _extract_tables_from_xml(doc_xml)


def find_text(workspace: Path, pattern: str, opts: FindOptions) -> list[TextMatch]:
    if not pattern:
        raise ValueError("search pattern cannot be empty")

    search_pattern = _compile_search_pattern(pattern, opts)
    matches: list[TextMatch] = []

    if opts.in_paragraphs or opts.in_tables:
        doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
        scoped_doc_xml = _scope_document_xml(doc_xml, opts.in_paragraphs, opts.in_tables)
        matches.extend(_find_in_xml(scoped_doc_xml, search_pattern, opts, opts.max_results))

    if opts.in_headers and (opts.max_results == 0 or len(matches) < opts.max_results):
        for header_path in sorted((workspace / "word").glob("header*.xml")):
            header_xml = header_path.read_text(encoding="utf-8")
            remaining = _remaining_limit(opts.max_results, len(matches))
            matches.extend(_find_in_xml(header_xml, search_pattern, opts, remaining))
            if opts.max_results > 0 and len(matches) >= opts.max_results:
                break

    if opts.in_footers and (opts.max_results == 0 or len(matches) < opts.max_results):
        for footer_path in sorted((workspace / "word").glob("footer*.xml")):
            footer_xml = footer_path.read_text(encoding="utf-8")
            remaining = _remaining_limit(opts.max_results, len(matches))
            matches.extend(_find_in_xml(footer_xml, search_pattern, opts, remaining))
            if opts.max_results > 0 and len(matches) >= opts.max_results:
                break

    return matches


def _compile_search_pattern(pattern: str, opts: FindOptions) -> re.Pattern[str]:
    if opts.use_regex:
        flags = 0 if opts.match_case else re.IGNORECASE
        return re.compile(pattern, flags)

    escaped = re.escape(pattern)
    if opts.whole_word:
        escaped = r"\b" + escaped + r"\b"
    flags = 0 if opts.match_case else re.IGNORECASE
    return re.compile(escaped, flags)


def _find_in_xml(
    xml: str,
    pattern: re.Pattern[str],
    opts: FindOptions,
    max_results: int,
) -> list[TextMatch]:
    full_text = _extract_text_from_xml(xml)
    paragraphs = _extract_paragraphs_from_xml(xml)
    matches: list[TextMatch] = []

    for match in pattern.finditer(full_text):
        if max_results > 0 and len(matches) >= max_results:
            break
        start, end = match.span()
        match_text = full_text[start:end]
        para_index = _find_paragraph_index(full_text, start, paragraphs)
        before = full_text[max(0, start - 50) : start]
        after = full_text[end : min(len(full_text), end + 50)]
        matches.append(
            TextMatch(
                text=match_text,
                paragraph=para_index,
                position=start,
                before=before,
                after=after,
            )
        )
    return matches


def _extract_text_from_xml(xml: str) -> str:
    parts: list[str] = []
    for match in _TEXT_PATTERN.findall(xml):
        parts.append(html_unescape(match))
    return "".join(parts)


def _extract_paragraphs_from_xml(xml: str) -> list[str]:
    paragraphs: list[str] = []
    for para in _PARA_PATTERN.findall(xml):
        text = _extract_text_from_xml(para)
        if text:
            paragraphs.append(text)
    return paragraphs


def _extract_tables_from_xml(xml: str) -> list[list[list[str]]]:
    tables: list[list[list[str]]] = []
    for table in _TABLE_PATTERN.findall(xml):
        rows: list[list[str]] = []
        for row in _ROW_PATTERN.findall(table):
            cells: list[str] = []
            for cell in _CELL_PATTERN.findall(row):
                cells.append(_extract_text_from_xml(cell))
            if cells:
                rows.append(cells)
        if rows:
            tables.append(rows)
    return tables


def _find_paragraph_index(full_text: str, position: int, paragraphs: Iterable[str]) -> int:
    current = 0
    for idx, para in enumerate(paragraphs):
        end = current + len(para)
        if current <= position < end:
            return idx
        current = end
    return -1


def _remaining_limit(max_results: int, current: int) -> int:
    if max_results == 0:
        return 0
    return max(0, max_results - current)


def _scope_document_xml(doc_xml: str, include_paragraphs: bool, include_tables: bool) -> str:
    if include_paragraphs and include_tables:
        return doc_xml
    if include_paragraphs:
        return _TABLE_PATTERN.sub("", doc_xml)
    if include_tables:
        return "".join(_TABLE_PATTERN.findall(doc_xml))
    return ""
