from __future__ import annotations

from pathlib import Path

from .rels import insert_relationship, next_relationship_id


def ensure_numbering_xml(workspace: Path) -> None:
    numbering_path = workspace / "word" / "numbering.xml"
    if numbering_path.exists():
        return
    numbering_path.parent.mkdir(parents=True, exist_ok=True)
    numbering_path.write_text(_numbering_xml(), encoding="utf-8")
    _add_numbering_relationship(workspace)
    _add_numbering_content_type(workspace)


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


def _numbering_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        # Bullet abstractNum (id=1)
        '<w:abstractNum w:abstractNumId="1">'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="•"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>'
        "</w:lvl>"
        "</w:abstractNum>"
        # Numbered abstractNum (id=2)
        '<w:abstractNum w:abstractNumId="2">'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        "</w:lvl>"
        "</w:abstractNum>"
        # Num instances
        '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>'
        "</w:numbering>"
    )
