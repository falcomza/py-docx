from __future__ import annotations

from pathlib import Path

import pytest

from pydocx.options import ParagraphAlignment, StyleDefinition, StyleType
from pydocx.styles import add_style
from pydocx.updater import Updater


def _styles(u: Updater) -> str:
    return (u.workspace / "word" / "styles.xml").read_text(encoding="utf-8")


def test_add_style_injects_into_existing_styles(blank: Updater, tmp_path: Path) -> None:
    add_style(
        blank.workspace,
        StyleDefinition(
            id="Callout",
            name="Callout Box",
            based_on="Normal",
            bold=True,
            font_size=24,
            color="FF0000",
            alignment=ParagraphAlignment.CENTER,
            space_after=120,
            outline_level=2,
        ),
    )
    xml = _styles(blank)
    assert '<w:style w:type="paragraph" w:styleId="Callout">' in xml
    assert '<w:name w:val="Callout Box"/>' in xml
    assert "<w:b/>" in xml and '<w:sz w:val="24"/>' in xml
    assert '<w:jc w:val="center"/>' in xml
    assert '<w:outlineLvl w:val="1"/>' in xml
    blank.save(tmp_path / "o.docx")


def test_add_style_creates_part_when_absent(blank: Updater) -> None:
    (blank.workspace / "word" / "styles.xml").unlink()
    # also drop the rel + content-type so we exercise the create path fully
    add_style(blank.workspace, StyleDefinition(id="X", type=StyleType.CHARACTER, italic=True))
    xml = _styles(blank)
    assert xml.startswith("<?xml")
    assert '<w:style w:type="character" w:styleId="X">' in xml
    rels = (blank.workspace / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
    assert "styles.xml" in rels
    ct = (blank.workspace / "[Content_Types].xml").read_text(encoding="utf-8")
    assert "/word/styles.xml" in ct


def test_empty_style_id_rejected(blank: Updater) -> None:
    with pytest.raises(ValueError):
        add_style(blank.workspace, StyleDefinition(id=""))
