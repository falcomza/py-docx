from __future__ import annotations

import re
from pathlib import Path

from .options import FooterOptions, HeaderFooterContent, HeaderOptions
from .rels import insert_relationship, next_relationship_id
from .xmlutils import xml_escape

_SECTPR_RE = re.compile(r"<w:sectPr[^>]*>.*?</w:sectPr>", re.DOTALL)


def set_header(workspace: Path, content: HeaderFooterContent, opts: HeaderOptions) -> None:
    filename = _header_filename(opts.type)
    header_path = workspace / "word" / filename
    header_xml = _generate_header_footer_xml(content, is_header=True)
    header_path.write_text(header_xml, encoding="utf-8")

    rel_id = _add_header_footer_relationship(workspace, filename, "header")
    _update_document_for_header_footer(
        workspace, opts.type.value, "header", rel_id, opts.different_first, opts.different_odd_even
    )
    _add_header_footer_content_type(workspace, filename, "header")


def set_footer(workspace: Path, content: HeaderFooterContent, opts: FooterOptions) -> None:
    filename = _footer_filename(opts.type)
    footer_path = workspace / "word" / filename
    footer_xml = _generate_header_footer_xml(content, is_header=False)
    footer_path.write_text(footer_xml, encoding="utf-8")

    rel_id = _add_header_footer_relationship(workspace, filename, "footer")
    _update_document_for_header_footer(
        workspace, opts.type.value, "footer", rel_id, opts.different_first, opts.different_odd_even
    )
    _add_header_footer_content_type(workspace, filename, "footer")


def _header_filename(header_type: str) -> str:
    if header_type == "first":
        return "header1.xml"
    if header_type == "even":
        return "header2.xml"
    if header_type == "default":
        return "header3.xml"
    return "header.xml"


def _footer_filename(footer_type: str) -> str:
    if footer_type == "first":
        return "footer1.xml"
    if footer_type == "even":
        return "footer2.xml"
    if footer_type == "default":
        return "footer3.xml"
    return "footer.xml"


def _generate_header_footer_xml(content: HeaderFooterContent, is_header: bool) -> str:
    root = "w:hdr" if is_header else "w:ftr"
    xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]
    if content.left_text or content.center_text or content.right_text:
        xml.append(_three_column_table(content))
    xml.append(f"</{root}>")
    return "\n".join(xml)


def _three_column_table(content: HeaderFooterContent) -> str:
    return (
        "<w:tbl>"
        "<w:tblPr>"
        '<w:tblW w:w="5000" w:type="pct"/>'
        '<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>'
        '<w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>'
        "</w:tblPr>"
        "<w:tblGrid>"
        '<w:gridCol w:w="3000"/>'
        '<w:gridCol w:w="3000"/>'
        '<w:gridCol w:w="3000"/>'
        "</w:tblGrid>"
        "<w:tr>"
        + _table_cell(content.left_text, "left")
        + _table_cell(content.center_text, "center")
        + _table_cell(content.right_text, "right")
        + "</w:tr>"
        "</w:tbl>"
    )


def _table_cell(text: str, align: str) -> str:
    if text:
        return (
            '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:jc w:val="{align}"/></w:pPr>'
            f"<w:r><w:t>{xml_escape(text)}</w:t></w:r></w:p></w:tc>"
        )
    return '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p/></w:tc>'


def _add_header_footer_relationship(workspace: Path, filename: str, hdr_ftr: str) -> str:
    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    existing = re.search(rf'<Relationship Id="([^"]+)"[^>]*Target="{re.escape(filename)}"', rels_xml)
    if existing:
        return existing.group(1)
    rel_id = next_relationship_id(rels_xml)
    rel_type = f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{hdr_ftr}"
    rel_xml = f'\n  <Relationship Id="{rel_id}" Type="{rel_type}" Target="{filename}"/>'
    rels_path.write_text(insert_relationship(rels_xml, rel_xml), encoding="utf-8")
    return rel_id


def _update_document_for_header_footer(
    workspace: Path,
    ref_type: str,
    hdr_ftr: str,
    rel_id: str,
    different_first: bool,
    different_odd_even: bool,
) -> None:
    doc_path = workspace / "word" / "document.xml"
    content = doc_path.read_text(encoding="utf-8")

    if _SECTPR_RE.search(content):
        content = _SECTPR_RE.sub(
            lambda m: _add_to_sectpr(m.group(0), ref_type, hdr_ftr, rel_id, different_first, different_odd_even),
            content,
            count=1,
        )
    else:
        sectpr = _create_sectpr(ref_type, hdr_ftr, rel_id, different_first, different_odd_even)
        content = content.replace("</w:body>", sectpr + "</w:body>", 1)
    doc_path.write_text(content, encoding="utf-8")


def _add_to_sectpr(
    sectpr: str,
    ref_type: str,
    hdr_ftr: str,
    rel_id: str,
    different_first: bool,
    different_odd_even: bool,
) -> str:
    ref = f'<w:{hdr_ftr}Reference w:type="{ref_type}" r:id="{rel_id}"/>'
    pattern = rf'<w:{hdr_ftr}Reference w:type="{ref_type}"[^>]*/>'
    if re.search(pattern, sectpr):
        sectpr = re.sub(pattern, ref, sectpr)
    else:
        sectpr = sectpr.replace("</w:sectPr>", ref + "</w:sectPr>", 1)
    if different_first and "<w:titlePg" not in sectpr:
        sectpr = sectpr.replace("</w:sectPr>", "<w:titlePg/></w:sectPr>", 1)
    if different_odd_even and "<w:evenAndOddHeaders" not in sectpr:
        sectpr = sectpr.replace("</w:sectPr>", "<w:evenAndOddHeaders/></w:sectPr>", 1)
    return sectpr


def _create_sectpr(
    ref_type: str,
    hdr_ftr: str,
    rel_id: str,
    different_first: bool,
    different_odd_even: bool,
) -> str:
    parts = ["<w:sectPr>"]
    if different_first:
        parts.append("<w:titlePg/>")
    if different_odd_even:
        parts.append("<w:evenAndOddHeaders/>")
    parts.append(f'<w:{hdr_ftr}Reference w:type="{ref_type}" r:id="{rel_id}"/>')
    parts.append("</w:sectPr>")
    return "".join(parts)


def _add_header_footer_content_type(workspace: Path, filename: str, hdr_ftr: str) -> None:
    content_types_path = workspace / "[Content_Types].xml"
    content = content_types_path.read_text(encoding="utf-8")
    if filename in content:
        return
    ctype = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
        if hdr_ftr == "header"
        else "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
    )
    override = f'<Override PartName="/word/{filename}" ContentType="{ctype}"/>'
    content = content.replace("</Types>", override + "</Types>", 1)
    content_types_path.write_text(content, encoding="utf-8")
