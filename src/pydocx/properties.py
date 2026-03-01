from __future__ import annotations

import re
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from .options import AppProperties, CoreProperties, CustomProperty
from .rels import ensure_content_type_override, insert_relationship, next_relationship_id
from .xmlutils import xml_escape


def set_core_properties(workspace: Path, props: CoreProperties) -> None:
    core_path = workspace / "docProps" / "core.xml"
    if core_path.exists():
        content = core_path.read_text(encoding="utf-8")
    else:
        content = _default_core_xml()

    content = _update_core_property(content, "dc:title", props.title)
    content = _update_core_property(content, "dc:subject", props.subject)
    content = _update_core_property(content, "dc:creator", props.creator)
    content = _update_core_property(content, "cp:keywords", props.keywords)
    content = _update_core_property(content, "dc:description", props.description)
    content = _update_core_property(content, "cp:category", props.category)
    content = _update_core_property(content, "cp:lastModifiedBy", props.last_modified_by)
    content = _update_core_property(content, "cp:revision", props.revision)
    content = _update_core_property(content, "cp:contentStatus", props.content_status)

    if props.created is not None:
        content = _update_core_date_property(content, "dcterms:created", _to_rfc3339(props.created))
    if props.modified is not None:
        content = _update_core_date_property(content, "dcterms:modified", _to_rfc3339(props.modified))
    else:
        content = _update_core_date_property(
            content,
            "dcterms:modified",
            _to_rfc3339(datetime.now(UTC)),
        )

    core_path.parent.mkdir(parents=True, exist_ok=True)
    core_path.write_text(content, encoding="utf-8")


def set_app_properties(workspace: Path, props: AppProperties) -> None:
    app_path = workspace / "docProps" / "app.xml"
    if app_path.exists():
        content = app_path.read_text(encoding="utf-8")
    else:
        content = _default_app_xml()

    content = _update_app_property(content, "Company", props.company)
    content = _update_app_property(content, "Manager", props.manager)
    content = _update_app_property(content, "Application", props.application)
    content = _update_app_property(content, "AppVersion", props.app_version)
    content = _update_app_property(content, "Template", props.template)
    content = _update_app_property(content, "HyperlinkBase", props.hyperlink_base)

    content = _update_app_property(content, "TotalTime", _int_or_empty(props.total_time))
    content = _update_app_property(content, "Pages", _int_or_empty(props.pages))
    content = _update_app_property(content, "Words", _int_or_empty(props.words))
    content = _update_app_property(content, "Characters", _int_or_empty(props.characters))
    content = _update_app_property(content, "CharactersWithSpaces", _int_or_empty(props.characters_with_spaces))
    content = _update_app_property(content, "Lines", _int_or_empty(props.lines))
    content = _update_app_property(content, "Paragraphs", _int_or_empty(props.paragraphs))
    content = _update_app_property(content, "DocSecurity", _int_or_empty(props.doc_security))

    app_path.parent.mkdir(parents=True, exist_ok=True)
    app_path.write_text(content, encoding="utf-8")


def set_custom_properties(workspace: Path, properties: list[CustomProperty]) -> None:
    custom_path = workspace / "docProps" / "custom.xml"
    content = _custom_properties_xml(properties)
    custom_path.parent.mkdir(parents=True, exist_ok=True)
    custom_path.write_text(content, encoding="utf-8")
    _ensure_custom_content_type(workspace)
    _ensure_custom_relationship(workspace)


def get_core_properties(workspace: Path) -> CoreProperties:
    core_path = workspace / "docProps" / "core.xml"
    if not core_path.exists():
        return CoreProperties()
    content = core_path.read_text(encoding="utf-8")
    props = CoreProperties()
    props.title = _extract_core_property(content, "dc:title")
    props.subject = _extract_core_property(content, "dc:subject")
    props.creator = _extract_core_property(content, "dc:creator")
    props.keywords = _extract_core_property(content, "cp:keywords")
    props.description = _extract_core_property(content, "dc:description")
    props.category = _extract_core_property(content, "cp:category")
    props.last_modified_by = _extract_core_property(content, "cp:lastModifiedBy")
    props.revision = _extract_core_property(content, "cp:revision")
    props.content_status = _extract_core_property(content, "cp:contentStatus")
    created = _extract_core_property(content, "dcterms:created")
    modified = _extract_core_property(content, "dcterms:modified")
    props.created = _parse_rfc3339(created)
    props.modified = _parse_rfc3339(modified)
    return props


def get_app_properties(workspace: Path) -> AppProperties:
    app_path = workspace / "docProps" / "app.xml"
    if not app_path.exists():
        return AppProperties()
    content = app_path.read_text(encoding="utf-8")
    props = AppProperties()
    props.company = _extract_app_property(content, "Company")
    props.manager = _extract_app_property(content, "Manager")
    props.application = _extract_app_property(content, "Application")
    props.app_version = _extract_app_property(content, "AppVersion")
    props.template = _extract_app_property(content, "Template")
    props.hyperlink_base = _extract_app_property(content, "HyperlinkBase")
    props.total_time = _int_from(_extract_app_property(content, "TotalTime"))
    props.pages = _int_from(_extract_app_property(content, "Pages"))
    props.words = _int_from(_extract_app_property(content, "Words"))
    props.characters = _int_from(_extract_app_property(content, "Characters"))
    props.characters_with_spaces = _int_from(_extract_app_property(content, "CharactersWithSpaces"))
    props.lines = _int_from(_extract_app_property(content, "Lines"))
    props.paragraphs = _int_from(_extract_app_property(content, "Paragraphs"))
    props.doc_security = _int_from(_extract_app_property(content, "DocSecurity"))
    return props


def get_custom_properties(workspace: Path) -> list[CustomProperty]:
    custom_path = workspace / "docProps" / "custom.xml"
    if not custom_path.exists():
        return []
    content = custom_path.read_text(encoding="utf-8")
    return _parse_custom_properties_xml(content)


def _update_core_property(content: str, prop: str, value: str) -> str:
    pattern = re.compile(rf"<{re.escape(prop)}[^>]*>.*?</{re.escape(prop)}>")
    if value == "":
        return pattern.sub("", content)
    escaped = xml_escape(value)
    replacement = f"<{prop}>{escaped}</{prop}>"
    if pattern.search(content):
        return pattern.sub(replacement, content)
    return content.replace("</cp:coreProperties>", replacement + "</cp:coreProperties>", 1)


def _update_core_date_property(content: str, prop: str, value: str) -> str:
    if value == "":
        return content
    if prop in ("dcterms:created", "dcterms:modified"):
        element = f'<{prop} xsi:type="dcterms:W3CDTF">{value}</{prop}>'
    else:
        element = f"<{prop}>{value}</{prop}>"
    pattern = re.compile(rf"<{re.escape(prop)}[^>]*>.*?</{re.escape(prop)}>")
    if pattern.search(content):
        return pattern.sub(element, content)
    return content.replace("</cp:coreProperties>", element + "</cp:coreProperties>", 1)


def _update_app_property(content: str, prop: str, value: str) -> str:
    pattern = re.compile(rf"<{re.escape(prop)}(?:\s[^>]*)?>.*?</{re.escape(prop)}>")
    if value == "":
        return pattern.sub("", content)
    escaped = xml_escape(value)
    replacement = f"<{prop}>{escaped}</{prop}>"
    if pattern.search(content):
        return pattern.sub(replacement, content)
    return content.replace("</Properties>", replacement + "</Properties>", 1)


def _extract_core_property(content: str, prop: str) -> str:
    pattern = re.compile(rf"<{re.escape(prop)}[^>]*>(.*?)</{re.escape(prop)}>")
    match = pattern.search(content)
    return match.group(1) if match else ""


def _extract_app_property(content: str, prop: str) -> str:
    pattern = re.compile(rf"<{re.escape(prop)}(?:\s[^>]*)?>([^<]*)</{re.escape(prop)}>")
    match = pattern.search(content)
    return match.group(1) if match else ""


def _default_core_xml() -> str:
    now = _to_rfc3339(datetime.now(UTC))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
        f"<cp:revision>1</cp:revision>\n"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>\n'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>\n'
        "</cp:coreProperties>"
    )


def _default_app_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\n'
        "<Application>Microsoft Word</Application>\n"
        "<DocSecurity>0</DocSecurity>\n"
        "<ScaleCrop>false</ScaleCrop>\n"
        "<LinksUpToDate>false</LinksUpToDate>\n"
        "<SharedDoc>false</SharedDoc>\n"
        "<HyperlinksChanged>false</HyperlinksChanged>\n"
        "<AppVersion>16.0000</AppVersion>\n"
        "</Properties>"
    )


def _custom_properties_xml(properties: list[CustomProperty]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
    ]
    for idx, prop in enumerate(properties):
        pid = idx + 2
        prop_type = prop.type or _infer_custom_type(prop.value)
        lines.append(
            f'<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{pid}" name="{xml_escape(prop.name)}">'
        )
        lines.append(_format_custom_value(prop.value, prop_type))
        lines.append("</property>")
    lines.append("</Properties>")
    return "\n".join(lines)


def _infer_custom_type(value: Any) -> str:
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, int):
        return "i4"
    if isinstance(value, float):
        return "r8"
    if isinstance(value, datetime):
        return "date"
    return "lpwstr"


def _format_custom_value(value: Any, prop_type: str) -> str:
    if prop_type == "lpwstr":
        return f"<vt:lpwstr>{xml_escape(str(value))}</vt:lpwstr>"
    if prop_type == "i4":
        return f"<vt:i4>{value}</vt:i4>"
    if prop_type == "r8":
        return f"<vt:r8>{value}</vt:r8>"
    if prop_type == "bool":
        return f"<vt:bool>{value}</vt:bool>"
    if prop_type == "date":
        if isinstance(value, datetime):
            return f"<vt:filetime>{_to_rfc3339(value)}</vt:filetime>"
        return f"<vt:lpwstr>{xml_escape(str(value))}</vt:lpwstr>"
    return f"<vt:lpwstr>{xml_escape(str(value))}</vt:lpwstr>"


def _parse_custom_properties_xml(content: str) -> list[CustomProperty]:
    properties: list[CustomProperty] = []
    prop_pattern = re.compile(r'(?s)<property[^>]*name="([^"]*?)"[^>]*>(.*?)</property>')
    for name, body in prop_pattern.findall(content):
        prop = CustomProperty(name=name, value="")
        if (value := _extract_vt_value(body, "lpwstr")) != "":
            prop.value = value
            prop.type = "lpwstr"
        elif (value := _extract_vt_value(body, "i4")) != "":
            prop.type = "i4"
            prop.value = _int_from(value)
        elif (value := _extract_vt_value(body, "r8")) != "":
            prop.type = "r8"
            prop.value = _float_from(value)
        elif (value := _extract_vt_value(body, "bool")) != "":
            prop.type = "bool"
            prop.value = value.lower() == "true"
        elif (value := _extract_vt_value(body, "filetime")) != "":
            prop.type = "date"
            prop.value = _parse_rfc3339(value) or value
        properties.append(prop)
    return properties


def _extract_vt_value(body: str, vt_type: str) -> str:
    pattern = re.compile(rf"<vt:{re.escape(vt_type)}[^>]*>(.*?)</vt:{re.escape(vt_type)}>")
    match = pattern.search(body)
    return match.group(1) if match else ""


def _ensure_custom_content_type(workspace: Path) -> None:
    ct_path = workspace / "[Content_Types].xml"
    content = ct_path.read_text(encoding="utf-8")
    updated = ensure_content_type_override(
        content,
        "/docProps/custom.xml",
        "application/vnd.openxmlformats-officedocument.custom-properties+xml",
    )
    ct_path.write_text(updated, encoding="utf-8")


def _ensure_custom_relationship(workspace: Path) -> None:
    rels_path = workspace / "_rels" / ".rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    if "docProps/custom.xml" in rels_xml:
        return
    rel_id = next_relationship_id(rels_xml)
    rel = (
        f'<Relationship Id="{rel_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" '
        'Target="docProps/custom.xml"/>'
    )
    rels_path.write_text(insert_relationship(rels_xml, "\n  " + rel), encoding="utf-8")


def _to_rfc3339(value: datetime) -> str:
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_rfc3339(value: str) -> datetime | None:
    if not value:
        return None
    try:
        if value.endswith("Z"):
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def _int_or_empty(value: int) -> str:
    return str(value) if value > 0 else ""


def _int_from(value: str) -> int:
    try:
        return int(value)
    except ValueError:
        return 0


def _float_from(value: str) -> float:
    try:
        return float(value)
    except ValueError:
        return 0.0
