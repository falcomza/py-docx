from __future__ import annotations

import re
from pathlib import Path

from .options import PageNumberOptions


def set_page_number(workspace: Path, opts: PageNumberOptions) -> None:
    if opts.start < 0:
        raise ValueError("page number start must be >= 0")

    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated = _set_pgnum_in_sectpr(doc_xml, opts)
    doc_path.write_text(updated, encoding="utf-8")


def _set_pgnum_in_sectpr(content: str, opts: PageNumberOptions) -> str:
    parts = ["<w:pgNumType"]
    if opts.start > 0:
        parts.append(f' w:start="{opts.start}"')
    if opts.format:
        parts.append(f' w:fmt="{opts.format.value}"')
    parts.append("/>")
    pg_num = "".join(parts)

    body_end = content.rfind("</w:body>")
    if body_end == -1:
        raise ValueError("could not find </w:body>")

    sect_start = content.rfind("<w:sectPr", 0, body_end)
    if sect_start == -1:
        sect = f"<w:sectPr>{pg_num}</w:sectPr>"
        return content[:body_end] + sect + content[body_end:]

    sect_end = content.find("</w:sectPr>", sect_start)
    if sect_end == -1:
        raise ValueError("malformed sectPr element")
    sect_end += len("</w:sectPr>")
    sect = content[sect_start:sect_end]

    if re.search(r"<w:pgNumType[^/]*/>", sect):
        sect = re.sub(r"<w:pgNumType[^/]*/>", pg_num, sect)
    else:
        sect = sect.replace("</w:sectPr>", pg_num + "</w:sectPr>", 1)

    return content[:sect_start] + sect + content[sect_end:]
