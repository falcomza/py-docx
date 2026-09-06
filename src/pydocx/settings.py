from __future__ import annotations

import re
from pathlib import Path

from .rels import ensure_content_type_override, insert_relationship, next_relationship_id

_SETTINGS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
_SETTINGS_PART_NAME = "/word/settings.xml"
_SETTINGS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"

_EMPTY_SETTINGS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
    "</w:settings>"
)

_UPDATE_FIELDS_TRUTHY_RE = re.compile(r'<w:updateFields\s[^>]*w:val="(?:1|true|on)"[^>]*?/?>')
_UPDATE_FIELDS_ANY_RE = re.compile(r"<w:updateFields[^>]*?/?>(?:</w:updateFields>)?")
_SETTINGS_SELF_CLOSING_RE = re.compile(r"<w:settings\b([^>]*)/>")
_FLDCHAR_BEGIN_RE = re.compile(r'<w:fldChar\s[^>]*w:fldCharType="begin"[^>]*?>')
_FLDSIMPLE_RE = re.compile(r"<w:fldSimple\s[^>]*?>")

_TRUE_FLAG = '<w:updateFields w:val="1"/>'


def ensure_settings_element(
    workspace: Path,
    element_xml: str,
    *,
    existing_re: re.Pattern[str] | None = None,
) -> None:
    """Ensure ``element_xml`` appears once inside ``word/settings.xml``, creating
    and wiring the part if it does not exist. If ``existing_re`` is given and
    matches, that match is replaced with ``element_xml`` (used to upgrade a falsy
    flag to a truthy one) — otherwise the element is inserted once.
    """
    settings_path = workspace / "word" / "settings.xml"
    if settings_path.exists():
        content = settings_path.read_text(encoding="utf-8")
    else:
        settings_path.parent.mkdir(parents=True, exist_ok=True)
        content = _EMPTY_SETTINGS_XML
        _add_settings_content_type(workspace)
        _add_settings_relationship(workspace)

    # A self-closing <w:settings/> root cannot hold children — expand it first.
    content = _SETTINGS_SELF_CLOSING_RE.sub(r"<w:settings\1></w:settings>", content, count=1)

    if existing_re is not None and existing_re.search(content):
        content = existing_re.sub(element_xml, content, count=1)
    elif "</w:settings>" in content:
        content = content.replace("</w:settings>", element_xml + "</w:settings>", 1)
    else:
        content += element_xml + "</w:settings>"

    settings_path.write_text(content, encoding="utf-8")


def ensure_update_fields(workspace: Path) -> None:
    """Ensure ``<w:updateFields w:val="1"/>`` in settings.xml so Word/LibreOffice
    recalculates field codes on open. Idempotent."""
    settings_path = workspace / "word" / "settings.xml"
    if settings_path.exists() and _UPDATE_FIELDS_TRUTHY_RE.search(settings_path.read_text(encoding="utf-8")):
        return
    ensure_settings_element(workspace, _TRUE_FLAG, existing_re=_UPDATE_FIELDS_ANY_RE)


def force_field_update_on_open(workspace: Path) -> None:
    """Like :func:`ensure_update_fields`, plus mark every field code in the header
    and footer parts ``w:dirty="true"`` (not all readers honour the document-level
    flag for those parts). Idempotent."""
    ensure_update_fields(workspace)
    for entry in (workspace / "word").glob("*.xml"):
        if entry.name.startswith(("header", "footer")):
            mark_fields_dirty(entry)


def mark_fields_dirty(xml_path: Path) -> None:
    """Add ``w:dirty="true"`` to every complex/simple field opening tag in one part."""
    if not xml_path.exists():
        return
    original = xml_path.read_text(encoding="utf-8")
    updated = _FLDSIMPLE_RE.sub(_add_dirty_attr, _FLDCHAR_BEGIN_RE.sub(_add_dirty_attr, original))
    if updated != original:
        xml_path.write_text(updated, encoding="utf-8")


def _add_dirty_attr(match: re.Match[str]) -> str:
    tag = match.group(0)
    if 'w:dirty="true"' in tag:
        return tag
    if tag.endswith("/>"):
        return tag[:-2] + ' w:dirty="true"/>'
    return tag[:-1] + ' w:dirty="true">'


def _add_settings_content_type(workspace: Path) -> None:
    ct_path = workspace / "[Content_Types].xml"
    ct_path.write_text(
        ensure_content_type_override(ct_path.read_text(encoding="utf-8"), _SETTINGS_PART_NAME, _SETTINGS_CONTENT_TYPE),
        encoding="utf-8",
    )


def _add_settings_relationship(workspace: Path) -> None:
    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    content = rels_path.read_text(encoding="utf-8")
    if _SETTINGS_REL_TYPE in content:
        return
    rel_id = next_relationship_id(content)
    rel = f'\n  <Relationship Id="{rel_id}" Type="{_SETTINGS_REL_TYPE}" Target="settings.xml"/>'
    rels_path.write_text(insert_relationship(content, rel), encoding="utf-8")
