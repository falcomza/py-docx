from __future__ import annotations

import re
from collections.abc import Callable
from pathlib import Path

from .options import ReplaceOptions
from .xmlutils import xml_escape

_TEXT_RUN_RE = re.compile(r"(<w:t(?:\s[^>]*)?>.*?</w:t>)", re.DOTALL)
_TEXT_CONTENT_RE = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.DOTALL)
_TABLE_PATTERN = re.compile(r"(?s)<w:tbl>.*?</w:tbl>")


def replace_text(workspace: Path, old: str, new: str, opts: ReplaceOptions) -> int:
    if not old:
        raise ValueError("old text cannot be empty")

    count = 0

    if opts.in_paragraphs or opts.in_tables:
        doc_path = workspace / "word" / "document.xml"
        count += _replace_in_file(doc_path, old, new, opts, count)

    if opts.in_headers:
        for header_path in sorted((workspace / "word").glob("header*.xml")):
            count += _replace_in_file(header_path, old, new, opts, count)
            if opts.max_replacements > 0 and count >= opts.max_replacements:
                break

    if opts.in_footers:
        for footer_path in sorted((workspace / "word").glob("footer*.xml")):
            count += _replace_in_file(footer_path, old, new, opts, count)
            if opts.max_replacements > 0 and count >= opts.max_replacements:
                break

    return count


def replace_text_regex(
    workspace: Path,
    pattern: str,
    replacement: str,
    opts: ReplaceOptions,
) -> int:
    if not pattern:
        raise ValueError("regex pattern cannot be empty")
    regex = re.compile(pattern)

    count = 0
    if opts.in_paragraphs or opts.in_tables:
        doc_path = workspace / "word" / "document.xml"
        count += _replace_regex_in_file(doc_path, regex, replacement, opts, count)

    if opts.in_headers:
        for header_path in sorted((workspace / "word").glob("header*.xml")):
            count += _replace_regex_in_file(header_path, regex, replacement, opts, count)
            if opts.max_replacements > 0 and count >= opts.max_replacements:
                break

    if opts.in_footers:
        for footer_path in sorted((workspace / "word").glob("footer*.xml")):
            count += _replace_regex_in_file(footer_path, regex, replacement, opts, count)
            if opts.max_replacements > 0 and count >= opts.max_replacements:
                break

    return count


def _replace_in_file(path: Path, old: str, new: str, opts: ReplaceOptions, current: int) -> int:
    xml = path.read_text(encoding="utf-8")
    updated, replaced = _replace_in_document_xml(
        xml=xml,
        opts=opts,
        current=current,
        replacer=lambda segment, offset: _replace_text_in_xml(segment, old, new, opts, offset),
    )
    if replaced > 0:
        path.write_text(updated, encoding="utf-8")
    return replaced


def _replace_regex_in_file(
    path: Path,
    pattern: re.Pattern[str],
    replacement: str,
    opts: ReplaceOptions,
    current: int,
) -> int:
    xml = path.read_text(encoding="utf-8")
    updated, replaced = _replace_in_document_xml(
        xml=xml,
        opts=opts,
        current=current,
        replacer=lambda segment, offset: _replace_regex_in_xml(segment, pattern, replacement, opts, offset),
    )
    if replaced > 0:
        path.write_text(updated, encoding="utf-8")
    return replaced


def _replace_in_document_xml(
    *,
    xml: str,
    opts: ReplaceOptions,
    current: int,
    replacer: Callable[[str, int], tuple[str, int]],
) -> tuple[str, int]:
    if opts.in_paragraphs and opts.in_tables:
        return replacer(xml, current)
    if not opts.in_paragraphs and not opts.in_tables:
        return xml, 0

    replaced_total = 0
    result: list[str] = []
    last_end = 0

    for table_match in _TABLE_PATTERN.finditer(xml):
        before = xml[last_end : table_match.start()]
        table_xml = table_match.group(0)

        if opts.in_paragraphs:
            updated_before, replaced = replacer(before, current + replaced_total)
            result.append(updated_before)
            result.append(table_xml)
        else:
            result.append(before)
            updated_table, replaced = replacer(table_xml, current + replaced_total)
            result.append(updated_table)

        replaced_total += replaced
        last_end = table_match.end()

    tail = xml[last_end:]
    if opts.in_paragraphs:
        updated_tail, replaced = replacer(tail, current + replaced_total)
        result.append(updated_tail)
        replaced_total += replaced
    else:
        result.append(tail)

    return "".join(result), replaced_total


def _replace_text_in_xml(
    xml: str,
    old: str,
    new: str,
    opts: ReplaceOptions,
    current: int,
) -> tuple[str, int]:
    replaced = 0
    escaped_old = re.escape(old)
    word_re = None
    case_insensitive_re = None

    if opts.whole_word:
        pattern = r"\b" + escaped_old + r"\b"
        flags = 0 if opts.match_case else re.IGNORECASE
        word_re = re.compile(pattern, flags)
    elif not opts.match_case:
        case_insensitive_re = re.compile(escaped_old, re.IGNORECASE)

    def repl(match: re.Match[str]) -> str:
        nonlocal replaced, current
        if opts.max_replacements > 0 and current >= opts.max_replacements:
            return match.group(0)

        text_match = _TEXT_CONTENT_RE.search(match.group(0))
        if not text_match:
            return match.group(0)
        raw_text = text_match.group(1)
        text = _xml_unescape(raw_text)

        if opts.whole_word and word_re is not None:

            def sub_word(m: re.Match[str]) -> str:
                nonlocal replaced, current
                if opts.max_replacements > 0 and current >= opts.max_replacements:
                    return m.group(0)
                replaced += 1
                current += 1
                return new

            replaced_text = word_re.sub(sub_word, text)
        elif opts.match_case:
            occurrences = text.count(old)
            if opts.max_replacements > 0:
                limit = max(opts.max_replacements - current, 0)
                occurrences = min(occurrences, limit)
                replaced_text = text.replace(old, new, limit)
            else:
                replaced_text = text.replace(old, new)
            if occurrences > 0:
                replaced += occurrences
                current += occurrences
        else:

            def sub_case(m: re.Match[str]) -> str:
                nonlocal replaced, current
                if opts.max_replacements > 0 and current >= opts.max_replacements:
                    return m.group(0)
                replaced += 1
                current += 1
                return new

            replaced_text = case_insensitive_re.sub(sub_case, text) if case_insensitive_re else text

        if replaced_text != text:
            return match.group(0).replace(raw_text, xml_escape(replaced_text), 1)
        return match.group(0)

    updated = _TEXT_RUN_RE.sub(repl, xml)
    return updated, replaced


def _replace_regex_in_xml(
    xml: str,
    pattern: re.Pattern[str],
    replacement: str,
    opts: ReplaceOptions,
    current: int,
) -> tuple[str, int]:
    replaced = 0

    def repl(match: re.Match[str]) -> str:
        nonlocal replaced, current
        if opts.max_replacements > 0 and current >= opts.max_replacements:
            return match.group(0)
        text_match = _TEXT_CONTENT_RE.search(match.group(0))
        if not text_match:
            return match.group(0)
        raw_text = text_match.group(1)
        text = _xml_unescape(raw_text)

        def sub(m: re.Match[str]) -> str:
            nonlocal replaced, current
            if opts.max_replacements > 0 and current >= opts.max_replacements:
                return m.group(0)
            replaced += 1
            current += 1
            return replacement

        replaced_text = pattern.sub(sub, text)
        if replaced_text != text:
            return match.group(0).replace(raw_text, xml_escape(replaced_text), 1)
        return match.group(0)

    updated = _TEXT_RUN_RE.sub(repl, xml)
    return updated, replaced


def _xml_unescape(value: str) -> str:
    value = value.replace("&amp;", "&")
    value = value.replace("&lt;", "<")
    value = value.replace("&gt;", ">")
    value = value.replace("&quot;", '"')
    value = value.replace("&apos;", "'")
    return value
