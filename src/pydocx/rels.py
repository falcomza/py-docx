from __future__ import annotations

import re

REL_CHART_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
REL_IMAGE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_PACKAGE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
IMAGE_CONTENT_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
}


_REL_ID_RE = re.compile(r'Id="rId(\d+)"')


def next_relationship_id(rels_xml: str) -> str:
    max_id = 0
    for match in _REL_ID_RE.finditer(rels_xml):
        num = int(match.group(1))
        if num > max_id:
            max_id = num
    return f"rId{max_id + 1}"


def insert_relationship(rels_xml: str, relationship_xml: str) -> str:
    marker = "</Relationships>"
    idx = rels_xml.rfind(marker)
    if idx == -1:
        raise ValueError("document.xml.rels missing </Relationships>")
    return rels_xml[:idx] + relationship_xml + rels_xml[idx:]


def ensure_content_type_override(content_types_xml: str, part_name: str, content_type: str) -> str:
    if part_name in content_types_xml:
        return content_types_xml
    marker = "</Types>"
    idx = content_types_xml.rfind(marker)
    if idx == -1:
        raise ValueError("[Content_Types].xml missing </Types>")
    override = f'\n  <Override PartName="{part_name}" ContentType="{content_type}"/>'
    return content_types_xml[:idx] + override + content_types_xml[idx:]


def ensure_default_content_type(content_types_xml: str, extension: str, content_type: str) -> str:
    ext = extension.lstrip(".")
    if f'Extension="{ext}"' in content_types_xml:
        return content_types_xml
    marker = "</Types>"
    idx = content_types_xml.rfind(marker)
    if idx == -1:
        raise ValueError("[Content_Types].xml missing </Types>")
    default = f'\n  <Default Extension="{ext}" ContentType="{content_type}"/>'
    return content_types_xml[:idx] + default + content_types_xml[idx:]
