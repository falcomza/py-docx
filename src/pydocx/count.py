from __future__ import annotations

import re
from pathlib import Path

_TABLE_PATTERN = re.compile(r"(?s)<w:tbl>")
_PARA_PATTERN = re.compile(r"(?s)<w:p[^>]*>")
_BLIP_PATTERN = re.compile(r'<a:blip[^>]*r:embed="[^"]*"[^>]*>')
_CHART_PATTERN = re.compile(r'(?s)<c:chart[^>]*r:id="[^"]*"[^>]*>')


def get_table_count(workspace: Path) -> int:
    doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
    return len(_TABLE_PATTERN.findall(doc_xml))


def get_paragraph_count(workspace: Path) -> int:
    doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
    return len(_PARA_PATTERN.findall(doc_xml))


def get_image_count(workspace: Path) -> int:
    doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
    return len(_BLIP_PATTERN.findall(doc_xml))


def get_chart_count(workspace: Path) -> int:
    doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
    return len(_CHART_PATTERN.findall(doc_xml))
