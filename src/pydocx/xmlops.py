from __future__ import annotations

from .xmlutils import xml_escape


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
    remaining = content
    offset = 0
    while True:
        ie = remaining.find(open_exact)
        ia = remaining.find(open_attr)
        idx = -1
        if ie >= 0 and ia >= 0:
            idx = ie if ie <= ia else ia
        elif ie >= 0:
            idx = ie
        elif ia >= 0:
            idx = ia
        if idx < 0:
            raise ValueError(f"only {count} {tag} element(s) found")
        count += 1
        abs_start = offset + idx
        close_idx = remaining.find(close_tag, idx)
        if close_idx < 0:
            raise ValueError(f"unclosed <{tag}>")
        abs_end = abs_start + (close_idx - idx) + len(close_tag)
        if count == n:
            return abs_start, abs_end
        advance = (close_idx - idx) + len(close_tag) + idx
        offset += advance
        remaining = remaining[advance:]


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
