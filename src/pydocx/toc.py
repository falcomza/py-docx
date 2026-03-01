from __future__ import annotations

from pathlib import Path

from .document import insert_at_body_end, insert_at_body_start
from .options import InsertPosition, TOCOptions
from .xmlutils import xml_escape


def insert_toc(workspace: Path, opts: TOCOptions) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    toc_xml = _build_toc_xml(opts)

    if opts.position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, toc_xml)
    elif opts.position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, toc_xml)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")

    doc_path.write_text(updated, encoding="utf-8")

    if opts.update_on_open:
        _set_update_fields(workspace)


def _build_toc_xml(opts: TOCOptions) -> str:
    parts = []
    if opts.title:
        parts.append(
            f'<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>{xml_escape(opts.title)}</w:t></w:r></w:p>'
        )
    instr = f'TOC \\o "{xml_escape(opts.outline_levels)}" \\h \\z \\u'
    field = (
        "<w:p>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f"<w:r><w:instrText>{instr}</w:instrText></w:r>"
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t> </w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "</w:p>"
    )
    parts.append(field)
    return "".join(parts)


def _set_update_fields(workspace: Path) -> None:
    settings_path = workspace / "word" / "settings.xml"
    if settings_path.exists():
        content = settings_path.read_text(encoding="utf-8")
        if "<w:updateFields" in content:
            content = _replace_update_fields(content)
        else:
            content = content.replace("</w:settings>", '<w:updateFields w:val="true"/></w:settings>', 1)
        settings_path.write_text(content, encoding="utf-8")
        return

    settings_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:updateFields w:val="true"/>'
        "</w:settings>"
    )
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    settings_path.write_text(settings_xml, encoding="utf-8")
    _ensure_settings_content_type(workspace)


def _replace_update_fields(content: str) -> str:
    start = content.find("<w:updateFields")
    if start == -1:
        return content
    end = content.find("/>", start)
    if end == -1:
        return content
    return content[:start] + '<w:updateFields w:val="true"/>' + content[end + 2 :]


def _ensure_settings_content_type(workspace: Path) -> None:
    ct_path = workspace / "[Content_Types].xml"
    content = ct_path.read_text(encoding="utf-8")
    if "/word/settings.xml" in content:
        return
    override = '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    content = content.replace("</Types>", override + "</Types>", 1)
    ct_path.write_text(content, encoding="utf-8")
