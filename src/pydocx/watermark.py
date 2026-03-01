from __future__ import annotations

import re
from pathlib import Path

from .options import WatermarkOptions
from .rels import insert_relationship, next_relationship_id
from .xmlutils import xml_escape


def set_text_watermark(workspace: Path, opts: WatermarkOptions) -> None:
    if not opts.text:
        raise ValueError("watermark text cannot be empty")

    font_family = opts.font_family or "Calibri"
    color = _normalize_hex_color(opts.color or "C0C0C0") or "C0C0C0"
    opacity = opts.opacity if opts.opacity > 0 else 0.5
    if opacity > 1.0:
        opacity = 1.0

    watermark_xml = _generate_watermark_xml(
        WatermarkOptions(
            text=opts.text,
            font_family=font_family,
            color=color,
            opacity=opacity,
            diagonal=opts.diagonal,
        )
    )

    header_file = _find_default_header_file(workspace)
    if header_file:
        _inject_watermark_into_header(workspace, header_file, watermark_xml)
    else:
        _create_watermark_header(workspace, watermark_xml)


def _find_default_header_file(workspace: Path) -> str:
    doc_path = workspace / "word" / "document.xml"
    content = doc_path.read_text(encoding="utf-8")
    match = re.search(r'<w:headerReference w:type="default" r:id="([^"]+)"/>', content)
    if not match:
        return ""
    rel_id = match.group(1)

    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    target_match = re.search(rf'<Relationship Id="{re.escape(rel_id)}"[^>]*Target="([^"]+)"', rels_xml)
    if not target_match:
        return ""
    return target_match.group(1)


def _inject_watermark_into_header(workspace: Path, header_file: str, watermark_xml: str) -> None:
    header_path = workspace / "word" / header_file
    content = header_path.read_text(encoding="utf-8")
    content = _ensure_vml_namespaces(content)

    hdr_idx = content.find("<w:hdr")
    if hdr_idx == -1:
        raise ValueError("could not find <w:hdr> element")
    hdr_close = content.find(">", hdr_idx)
    if hdr_close == -1:
        raise ValueError("malformed <w:hdr> element")
    insert_pos = hdr_close + 1
    updated = content[:insert_pos] + "\n" + watermark_xml + content[insert_pos:]
    header_path.write_text(updated, encoding="utf-8")


def _create_watermark_header(workspace: Path, watermark_xml: str) -> None:
    header_file = _next_header_filename(workspace)
    header_path = workspace / "word" / header_file

    header_xml = "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office">',
            watermark_xml,
            "</w:hdr>",
        ]
    )
    header_path.write_text(header_xml, encoding="utf-8")

    rel_id = _add_header_relationship(workspace, header_file)
    _update_document_header_reference(workspace, rel_id)
    _add_header_content_type(workspace, header_file)


def _next_header_filename(workspace: Path) -> str:
    for i in range(1, 11):
        name = f"header{i}.xml"
        if not (workspace / "word" / name).exists():
            return name
    return "header11.xml"


def _add_header_relationship(workspace: Path, filename: str) -> str:
    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    existing = re.search(rf'<Relationship Id="([^"]+)"[^>]*Target="{re.escape(filename)}"', rels_xml)
    if existing:
        return existing.group(1)
    rel_id = next_relationship_id(rels_xml)
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    rel_xml = f'\n  <Relationship Id="{rel_id}" Type="{rel_type}" Target="{filename}"/>'
    rels_path.write_text(insert_relationship(rels_xml, rel_xml), encoding="utf-8")
    return rel_id


def _update_document_header_reference(workspace: Path, rel_id: str) -> None:
    doc_path = workspace / "word" / "document.xml"
    content = doc_path.read_text(encoding="utf-8")

    sect_match = re.search(r"<w:sectPr[^>]*>.*?</w:sectPr>", content, re.DOTALL)
    ref = f'<w:headerReference w:type="default" r:id="{rel_id}"/>'
    if sect_match:
        sectpr = sect_match.group(0)
        if re.search(r'<w:headerReference w:type="default"[^>]*/>', sectpr):
            updated_sect = re.sub(
                r'<w:headerReference w:type="default"[^>]*/>',
                ref,
                sectpr,
            )
        else:
            updated_sect = sectpr.replace("</w:sectPr>", ref + "</w:sectPr>", 1)
        content = content.replace(sectpr, updated_sect, 1)
    else:
        content = content.replace("</w:body>", f"<w:sectPr>{ref}</w:sectPr></w:body>", 1)

    doc_path.write_text(content, encoding="utf-8")


def _add_header_content_type(workspace: Path, filename: str) -> None:
    ct_path = workspace / "[Content_Types].xml"
    content = ct_path.read_text(encoding="utf-8")
    if filename in content:
        return
    override = (
        f'<Override PartName="/word/{filename}" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    )
    content = content.replace("</Types>", override + "</Types>", 1)
    ct_path.write_text(content, encoding="utf-8")


def _generate_watermark_xml(opts: WatermarkOptions) -> str:
    rotation = "rotation:315;" if opts.diagonal else ""
    return (
        "<w:p>"
        '<w:pPr><w:pStyle w:val="Header"/></w:pPr>'
        "<w:r>"
        "<w:rPr><w:noProof/></w:rPr>"
        "<w:pict>"
        '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" '
        'path="m@7,l@8,m@5,21600l@6,21600e">'
        "<v:formulas>"
        '<v:f eqn="sum #0 0 10800"/>'
        '<v:f eqn="prod #0 2 1"/>'
        '<v:f eqn="sum 21600 0 @1"/>'
        '<v:f eqn="sum 0 0 @2"/>'
        '<v:f eqn="sum 21600 0 @3"/>'
        '<v:f eqn="if @0 @3 0"/>'
        '<v:f eqn="if @0 21600 @1"/>'
        '<v:f eqn="if @0 0 @2"/>'
        '<v:f eqn="if @0 @4 21600"/>'
        '<v:f eqn="mid @5 @6"/>'
        '<v:f eqn="mid @8 @5"/>'
        '<v:f eqn="mid @7 @8"/>'
        '<v:f eqn="mid @6 @7"/>'
        '<v:f eqn="sum @6 0 @5"/>'
        "</v:formulas>"
        '<v:path textpathok="t" o:connecttype="custom" '
        'o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" '
        'o:connectangles="270,180,90,0"/>'
        '<v:textpath on="t" fitshape="t"/>'
        '<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>'
        '<o:lock v:ext="edit" text="t" shapetype="t"/>'
        "</v:shapetype>"
        f'<v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" type="#_x0000_t136" '
        f'style="position:absolute;margin-left:0;margin-top:0;'
        f"width:468pt;height:117pt;{rotation}"
        "z-index:-251658752;"
        "mso-position-horizontal:center;mso-position-horizontal-relative:margin;"
        'mso-position-vertical:center;mso-position-vertical-relative:margin" '
        f'o:allowincell="f" fillcolor="#{opts.color}" stroked="f">'
        f'<v:fill opacity="{opts.opacity:.2f}"/>'
        f'<v:textpath style="font-family:&quot;{xml_escape(opts.font_family)}&quot;;font-size:1pt" '
        f'string="{xml_escape(opts.text)}"/>'
        "</v:shape>"
        "</w:pict>"
        "</w:r>"
        "</w:p>"
    )


def _ensure_vml_namespaces(header_xml: str) -> str:
    if 'xmlns:v="' not in header_xml:
        header_xml = header_xml.replace(
            "<w:hdr ",
            '<w:hdr xmlns:v="urn:schemas-microsoft-com:vml" ',
            1,
        )
    if 'xmlns:o="' not in header_xml:
        header_xml = header_xml.replace(
            "<w:hdr ",
            '<w:hdr xmlns:o="urn:schemas-microsoft-com:office:office" ',
            1,
        )
    return header_xml


def _normalize_hex_color(color: str) -> str:
    c = color.strip().lstrip("#")
    if len(c) != 6:
        return ""
    if not all(ch in "0123456789abcdefABCDEF" for ch in c):
        return ""
    return c.upper()
