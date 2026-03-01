from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET

from .errors import InvalidPackageError

_REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
_CT_NS = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}


def validate_workspace(workspace: Path) -> None:
    _ensure_required_files(workspace)
    _validate_content_types(workspace)
    _validate_relationships(workspace)


def _ensure_required_files(workspace: Path) -> None:
    required = [
        workspace / "[Content_Types].xml",
        workspace / "_rels" / ".rels",
        workspace / "word" / "document.xml",
        workspace / "word" / "_rels" / "document.xml.rels",
    ]
    missing = [str(p.relative_to(workspace)) for p in required if not p.exists()]
    if missing:
        raise InvalidPackageError(f"missing required package files: {', '.join(missing)}")


def _validate_content_types(workspace: Path) -> None:
    path = workspace / "[Content_Types].xml"
    root = _parse_xml(path)
    overrides = {
        node.attrib.get("PartName", "")
        for node in root.findall("ct:Override", _CT_NS)
        if node.attrib.get("PartName")
    }
    if "/word/document.xml" not in overrides:
        raise InvalidPackageError("[Content_Types].xml is missing override for /word/document.xml")


def _validate_relationships(workspace: Path) -> None:
    rels_files = [
        (workspace / "_rels" / ".rels", workspace),
        (workspace / "word" / "_rels" / "document.xml.rels", workspace / "word"),
    ]

    for rels_path, base_dir in rels_files:
        root = _parse_xml(rels_path)
        seen_ids: set[str] = set()
        for rel in root.findall("r:Relationship", _REL_NS):
            rel_id = rel.attrib.get("Id", "")
            if not rel_id:
                raise InvalidPackageError(f"{rels_path.name} contains relationship without Id")
            if rel_id in seen_ids:
                raise InvalidPackageError(f"{rels_path.name} contains duplicate relationship Id: {rel_id}")
            seen_ids.add(rel_id)

            if rel.attrib.get("TargetMode") == "External":
                continue

            target = rel.attrib.get("Target", "")
            if not target:
                raise InvalidPackageError(f"{rels_path.name} relationship {rel_id} is missing Target")

            resolved = (base_dir / target).resolve()
            ws_resolved = workspace.resolve()
            if ws_resolved not in resolved.parents and resolved != ws_resolved:
                raise InvalidPackageError(f"{rels_path.name} relationship {rel_id} target escapes workspace: {target}")
            if not resolved.exists():
                rel_file = str(rels_path.relative_to(workspace))
                raise InvalidPackageError(f"{rel_file} relationship {rel_id} points to missing target: {target}")


def _parse_xml(path: Path) -> ET.Element:
    try:
        return ET.fromstring(path.read_text(encoding="utf-8"))
    except ET.ParseError as exc:
        raise InvalidPackageError(f"malformed XML in {path.name}: {exc}") from exc
