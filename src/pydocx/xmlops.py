from __future__ import annotations

import re
from html import unescape as html_unescape

from .xmlutils import xml_escape, xml_unescape

_PARA_TEXT_RE = re.compile(r"<w:t[^>]*>(.*?)</w:t>", re.DOTALL)
_PARA_BLOCK_RE = re.compile(r"(?s)<w:p(?:\s[^>]*)?>.*?</w:p>")
_RUN_BLOCK_RE = re.compile(r"(?s)<w:r(?:\s[^>]*)?>.*?</w:r>")
_RUN_RPR_RE = re.compile(r"(?s)<w:rPr>.*?</w:rPr>")
_RUN_TEXT_RE = re.compile(r"(?s)<w:t(?:\s[^>]*)?>(.*?)</w:t>")


def write_run_text(text: str) -> str:
    """Build OOXML run text nodes, handling spaces, newlines, and tabs."""
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


def find_nth_xml_block(content: str, tag: str, n: int) -> tuple[int, int]:
    open_exact = f"<{tag}>"
    open_attr = f"<{tag} "
    close_tag = f"</{tag}>"

    count = 0
    pos = 0
    while True:
        ie = content.find(open_exact, pos)
        ia = content.find(open_attr, pos)
        if ie < 0 and ia < 0:
            raise ValueError(f"only {count} {tag} element(s) found")
        if ie >= 0 and ia >= 0:
            idx = ie if ie <= ia else ia
        elif ie >= 0:
            idx = ie
        else:
            idx = ia
        count += 1
        close_idx = content.find(close_tag, idx)
        if close_idx < 0:
            raise ValueError(f"unclosed <{tag}>")
        end = close_idx + len(close_tag)
        if count == n:
            return idx, end
        pos = end


def extract_paragraph_text(para_xml: str) -> str:
    return "".join(html_unescape(m.group(1)) for m in _PARA_TEXT_RE.finditer(para_xml))


def normalize_ws(text: str) -> str:
    return " ".join(text.split())


def find_paragraph_range(doc_xml: str, anchor: str) -> tuple[int, int]:
    norm_anchor = normalize_ws(anchor)
    pos = 0
    while True:
        start = doc_xml.find("<w:p", pos)
        if start == -1:
            return -1, -1
        end = doc_xml.find("</w:p>", start)
        if end == -1:
            return -1, -1
        end += len("</w:p>")
        text = extract_paragraph_text(doc_xml[start:end])
        if anchor in text or norm_anchor in normalize_ws(text):
            return start, end
        pos = end


def insert_after_anchor(doc_xml: str, fragment: str, anchor: str) -> str:
    start, end = find_paragraph_range(doc_xml, anchor)
    if start == -1:
        raise ValueError(f"anchor text {anchor!r} not found in document")
    return doc_xml[:end] + fragment + doc_xml[end:]


def insert_before_anchor(doc_xml: str, fragment: str, anchor: str) -> str:
    start, _ = find_paragraph_range(doc_xml, anchor)
    if start == -1:
        raise ValueError(f"anchor text {anchor!r} not found in document")
    return doc_xml[:start] + fragment + doc_xml[start:]


def build_rpr_xml(
    bold: bool,
    italic: bool,
    underline: bool,
    *,
    font_family: str = "",
    font_size: int = 0,
    font_color: str = "",
    strikethrough: bool = False,
    highlight: str = "",
    all_caps: bool = False,
    small_caps: bool = False,
) -> str:
    # CT_RPr child order (ECMA-376 §17.3.2.28): rFonts, b, i, caps, smallCaps,
    # strike, color, sz, szCs, highlight, u.
    parts = []
    if font_family:
        esc = xml_escape(font_family)
        parts.append(f'<w:rFonts w:ascii="{esc}" w:hAnsi="{esc}" w:cs="{esc}"/>')
    if bold:
        parts.append("<w:b/>")
    if italic:
        parts.append("<w:i/>")
    if all_caps:
        parts.append("<w:caps/>")
    if small_caps:
        parts.append("<w:smallCaps/>")
    if strikethrough:
        parts.append("<w:strike/>")
    if font_color:
        parts.append(f'<w:color w:val="{xml_escape(font_color)}"/>')
    if font_size > 0:
        parts.append(f'<w:sz w:val="{font_size}"/><w:szCs w:val="{font_size}"/>')
    if highlight:
        parts.append(f'<w:highlight w:val="{xml_escape(highlight)}"/>')
    if underline:
        parts.append('<w:u w:val="single"/>')
    return f"<w:rPr>{''.join(parts)}</w:rPr>" if parts else ""


def merge_adjacent_runs(xml: str) -> str:
    """Merge consecutive <w:r> runs that share identical <w:rPr> within each
    paragraph, so text placeholders split across runs by Word become matchable.
    Runs separated by any markup (hyperlink/bookmark/ins boundaries) are left
    alone.
    """
    return _PARA_BLOCK_RE.sub(lambda m: _merge_runs_in_paragraph(m.group(0)), xml)


_MERGE_UNSAFE_RE = re.compile(r"<w:(?:br|tab|drawing|object|fldChar|instrText)\b")


def _run_is_plain(run: str) -> bool:
    return "<w:t" in run and not _MERGE_UNSAFE_RE.search(run)


def _merge_runs_in_paragraph(para: str) -> str:
    runs = list(_RUN_BLOCK_RE.finditer(para))
    if len(runs) <= 1:
        return para

    replacements: list[tuple[int, int, str]] = []  # (start, end, replacement-run) for merged streaks
    streak: list[re.Match[str]] = []
    streak_rpr = ""

    def flush() -> None:
        if len(streak) >= 2:
            text = "".join(xml_unescape(t) for m in streak for t in _RUN_TEXT_RE.findall(m.group(0)))
            t_open = '<w:t xml:space="preserve">' if text != text.strip() else "<w:t>"
            replacements.append(
                (streak[0].start(), streak[-1].end(), f"<w:r>{streak_rpr}{t_open}{xml_escape(text)}</w:t></w:r>")
            )
        streak.clear()

    for run in runs:
        rpr_match = _RUN_RPR_RE.search(run.group(0))
        rpr = rpr_match.group(0) if rpr_match else ""
        gap = para[streak[-1].end() : run.start()] if streak else ""
        if streak and _run_is_plain(run.group(0)) and rpr == streak_rpr and "<" not in gap and ">" not in gap:
            streak.append(run)
            continue
        flush()
        if _run_is_plain(run.group(0)):
            streak.append(run)
            streak_rpr = rpr
    flush()

    if not replacements:
        return para
    out: list[str] = []
    pos = 0
    for start, end, replacement in replacements:
        out.append(para[pos:start])
        out.append(replacement)
        pos = end
    out.append(para[pos:])
    return "".join(out)


def inject_tcpr_element(cell_content: str, element: str) -> str:
    if "<w:tcPr/>" in cell_content:
        return cell_content.replace("<w:tcPr/>", f"<w:tcPr>{element}</w:tcPr>", 1)
    idx = cell_content.find("<w:tcPr>")
    if idx >= 0:
        insert_pos = idx + len("<w:tcPr>")
        return cell_content[:insert_pos] + element + cell_content[insert_pos:]
    idx = cell_content.find("<w:tcPr ")
    if idx >= 0:
        close_idx = cell_content.find(">", idx)
        if close_idx >= 0:
            insert_pos = close_idx + 1
            return cell_content[:insert_pos] + element + cell_content[insert_pos:]
    open_end = cell_content.find(">")
    if open_end < 0:
        return cell_content
    insert_pos = open_end + 1
    return cell_content[:insert_pos] + f"<w:tcPr>{element}</w:tcPr>" + cell_content[insert_pos:]


def replace_cell_text(tc_content: str, value: str) -> str:
    escaped = xml_escape(value)

    tc_pr = ""
    if "<w:tcPr>" in tc_content:
        start = tc_content.find("<w:tcPr>")
        end = tc_content.find("</w:tcPr>", start)
        if end >= 0:
            tc_pr = tc_content[start : end + len("</w:tcPr>")]
    elif "<w:tcPr " in tc_content:
        start = tc_content.find("<w:tcPr ")
        end = tc_content.find("</w:tcPr>", start)
        if end >= 0:
            tc_pr = tc_content[start : end + len("</w:tcPr>")]

    p_pr = ""
    try:
        p_start, p_end = find_nth_xml_block(tc_content, "w:p", 1)
        p_content = tc_content[p_start:p_end]
        pp_start = p_content.find("<w:pPr>")
        if pp_start >= 0:
            pp_end = p_content.find("</w:pPr>", pp_start)
            if pp_end >= 0:
                p_pr = p_content[pp_start : pp_end + len("</w:pPr>")]
    except ValueError:
        open_tag = "<w:tc>"
        if "<w:tc " in tc_content:
            end = tc_content.find(">")
            if end >= 0:
                open_tag = tc_content[: end + 1]
        if escaped:
            return open_tag + tc_pr + f'<w:p><w:r><w:t xml:space="preserve">{escaped}</w:t></w:r></w:p>' + "</w:tc>"
        return open_tag + tc_pr + "<w:p/>" + "</w:tc>"

    tc_open_end = tc_content.find(">")
    tc_open_tag = tc_content[: tc_open_end + 1]

    p_open_end = p_content.find(">")
    p_open_tag = p_content[: p_open_end + 1]

    result = tc_open_tag + tc_pr + p_open_tag + p_pr
    if escaped:
        result += f'<w:r><w:t xml:space="preserve">{escaped}</w:t></w:r>'
    result += "</w:p></w:tc>"
    return result
