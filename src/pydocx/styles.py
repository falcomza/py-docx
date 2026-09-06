from __future__ import annotations

from pathlib import Path

from .options import StyleDefinition, StyleType
from .rels import ensure_content_type_override, insert_relationship, next_relationship_id
from .xmlops import build_rpr_xml
from .xmlutils import xml_escape

_STYLES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
_STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"


def add_style(workspace: Path, defn: StyleDefinition) -> None:
    if not defn.id:
        raise ValueError("style id cannot be empty")

    style_xml = _generate_style_xml(defn)
    styles_path = workspace / "word" / "styles.xml"

    if styles_path.exists():
        content = styles_path.read_text(encoding="utf-8")
        idx = content.rfind("</w:styles>")
        if idx == -1:
            raise ValueError("styles.xml missing </w:styles>")
        content = content[:idx] + style_xml + content[idx:]
        styles_path.write_text(content, encoding="utf-8")
        return

    styles_path.parent.mkdir(parents=True, exist_ok=True)
    styles_path.write_text(_generate_styles_document(style_xml), encoding="utf-8")
    _ensure_styles_relationship(workspace)
    _ensure_styles_content_type(workspace)


def add_styles(workspace: Path, defns: list[StyleDefinition]) -> None:
    for i, defn in enumerate(defns):
        try:
            add_style(workspace, defn)
        except Exception as exc:
            raise ValueError(f"add style {i} ({defn.id}) failed: {exc}") from exc


def _generate_style_xml(defn: StyleDefinition) -> str:
    name = defn.name or defn.id
    parts = [f'<w:style w:type="{defn.type.value}" w:styleId="{xml_escape(defn.id)}">']
    parts.append(f'<w:name w:val="{xml_escape(name)}"/>')
    if defn.based_on:
        parts.append(f'<w:basedOn w:val="{xml_escape(defn.based_on)}"/>')
    if defn.next_style:
        parts.append(f'<w:next w:val="{xml_escape(defn.next_style)}"/>')
    if defn.type == StyleType.PARAGRAPH:
        parts.append(_style_paragraph_props(defn))
    parts.append(_style_run_props(defn))
    parts.append("</w:style>")
    return "".join(parts)


def _style_paragraph_props(defn: StyleDefinition) -> str:
    inner: list[str] = []
    if defn.keep_next:
        inner.append("<w:keepNext/>")
    if defn.keep_lines:
        inner.append("<w:keepLines/>")
    if defn.page_break_before:
        inner.append("<w:pageBreakBefore/>")
    if defn.space_before or defn.space_after or defn.line_spacing:
        spacing = "<w:spacing"
        if defn.space_before:
            spacing += f' w:before="{defn.space_before}"'
        if defn.space_after:
            spacing += f' w:after="{defn.space_after}"'
        if defn.line_spacing:
            spacing += f' w:line="{defn.line_spacing}" w:lineRule="auto"'
        inner.append(spacing + "/>")
    if defn.indent_left or defn.indent_right or defn.indent_first:
        ind = "<w:ind"
        if defn.indent_left:
            ind += f' w:left="{defn.indent_left}"'
        if defn.indent_right:
            ind += f' w:right="{defn.indent_right}"'
        if defn.indent_first:
            ind += f' w:firstLine="{defn.indent_first}"'
        inner.append(ind + "/>")
    if defn.alignment is not None:
        inner.append(f'<w:jc w:val="{defn.alignment.value}"/>')
    if 1 <= defn.outline_level <= 9:
        inner.append(f'<w:outlineLvl w:val="{defn.outline_level - 1}"/>')
    return f"<w:pPr>{''.join(inner)}</w:pPr>" if inner else ""


def _style_run_props(defn: StyleDefinition) -> str:
    return build_rpr_xml(
        defn.bold,
        defn.italic,
        defn.underline,
        font_family=defn.font_family,
        font_size=defn.font_size,
        font_color=defn.color,
        strikethrough=defn.strikethrough,
        all_caps=defn.all_caps,
        small_caps=defn.small_caps,
    )


def _generate_styles_document(style_xml: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"{style_xml}"
        "</w:styles>"
    )


def _ensure_styles_relationship(workspace: Path) -> None:
    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    content = rels_path.read_text(encoding="utf-8")
    if "styles.xml" in content:
        return
    rel_id = next_relationship_id(content)
    rel = f'\n  <Relationship Id="{rel_id}" Type="{_STYLES_REL_TYPE}" Target="styles.xml"/>'
    rels_path.write_text(insert_relationship(content, rel), encoding="utf-8")


def _ensure_styles_content_type(workspace: Path) -> None:
    ct_path = workspace / "[Content_Types].xml"
    ct_path.write_text(
        ensure_content_type_override(ct_path.read_text(encoding="utf-8"), "/word/styles.xml", _STYLES_CONTENT_TYPE),
        encoding="utf-8",
    )
