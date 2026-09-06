from __future__ import annotations

import re
from pathlib import Path

from .rels import insert_relationship, next_relationship_id

BULLET_ABSTRACT_ID = 1
NUMBERED_ABSTRACT_ID = 2
BULLET_NUM_ID = 1
NUMBERED_NUM_ID = 2

_NUM_ID_RE = re.compile(r'<w:num\s+w:numId="(\d+)"')


def ensure_numbering_xml(workspace: Path) -> None:
    numbering_path = workspace / "word" / "numbering.xml"
    if numbering_path.exists():
        return
    numbering_path.parent.mkdir(parents=True, exist_ok=True)
    numbering_path.write_text(_numbering_xml(), encoding="utf-8")
    _add_numbering_relationship(workspace)
    _add_numbering_content_type(workspace)


def allocate_restart_num_id(workspace: Path, level: int) -> int:
    """Append a <w:num> that reuses the numbered abstract list but overrides the
    given level to restart at 1, and return its new numId. This is the correct
    OOXML mechanism for restarting numbered-list numbering (ECMA-376 §17.9.20).
    """
    ensure_numbering_xml(workspace)
    level = max(0, min(level, 8))
    numbering_path = workspace / "word" / "numbering.xml"
    content = numbering_path.read_text(encoding="utf-8")

    next_id = 1
    for match in _NUM_ID_RE.finditer(content):
        next_id = max(next_id, int(match.group(1)) + 1)

    new_num = (
        f'<w:num w:numId="{next_id}">'
        f'<w:abstractNumId w:val="{NUMBERED_ABSTRACT_ID}"/>'
        f'<w:lvlOverride w:ilvl="{level}"><w:startOverride w:val="1"/></w:lvlOverride>'
        "</w:num>"
    )
    content = content.replace("</w:numbering>", new_num + "</w:numbering>", 1)
    numbering_path.write_text(content, encoding="utf-8")
    return next_id


def _add_numbering_relationship(workspace: Path) -> None:
    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    if "numbering.xml" in rels_xml:
        return
    rel_id = next_relationship_id(rels_xml)
    rel = (
        f'<Relationship Id="{rel_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
        'Target="numbering.xml"/>'
    )
    rels_path.write_text(insert_relationship(rels_xml, "\n  " + rel), encoding="utf-8")


def _add_numbering_content_type(workspace: Path) -> None:
    ct_path = workspace / "[Content_Types].xml"
    content = ct_path.read_text(encoding="utf-8")
    if "/word/numbering.xml" in content:
        return
    override = (
        '<Override PartName="/word/numbering.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    )
    content = content.replace("</Types>", override + "</Types>", 1)
    ct_path.write_text(content, encoding="utf-8")


# Private-use codepoints that map to bullet glyphs in Symbol / Wingdings fonts.
_BULLET_SYMBOLS = ("\uf0b7", "o", "\uf0a7")
_BULLET_FONTS = ("Symbol", "Courier New", "Wingdings")


def _levels(*, numbered: bool) -> str:
    parts = []
    for level in range(9):
        if numbered:
            num_fmt = "decimal"
            lvl_text = "".join(f"%{i}." for i in range(1, level + 2))
            rpr = ""
        else:
            font = _BULLET_FONTS[level % 3]
            num_fmt = "bullet"
            lvl_text = _BULLET_SYMBOLS[level % 3]
            rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr>'
        parts.append(
            f'<w:lvl w:ilvl="{level}">'
            '<w:start w:val="1"/>'
            f'<w:numFmt w:val="{num_fmt}"/>'
            f'<w:lvlText w:val="{lvl_text}"/>'
            '<w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{720 * (level + 1)}" w:hanging="360"/></w:pPr>'
            f"{rpr}"
            "</w:lvl>"
        )
    return "".join(parts)


def _numbering_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:abstractNum w:abstractNumId="{BULLET_ABSTRACT_ID}">'
        '<w:multiLevelType w:val="hybridMultilevel"/>'
        f"{_levels(numbered=False)}"
        "</w:abstractNum>"
        f'<w:abstractNum w:abstractNumId="{NUMBERED_ABSTRACT_ID}">'
        '<w:multiLevelType w:val="hybridMultilevel"/>'
        f"{_levels(numbered=True)}"
        "</w:abstractNum>"
        f'<w:num w:numId="{BULLET_NUM_ID}"><w:abstractNumId w:val="{BULLET_ABSTRACT_ID}"/></w:num>'
        f'<w:num w:numId="{NUMBERED_NUM_ID}"><w:abstractNumId w:val="{NUMBERED_ABSTRACT_ID}"/></w:num>'
        "</w:numbering>"
    )
