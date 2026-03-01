from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


def load_xml(path: str | Path) -> ET.ElementTree:
    return ET.parse(path)


def save_xml(tree: ET.ElementTree, path: str | Path) -> None:
    tree.write(path, encoding="utf-8", xml_declaration=True)


def xml_escape(value: str) -> str:
    return escape(value, {"'": "&apos;", '"': "&quot;"})
