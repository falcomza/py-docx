from __future__ import annotations

from pydocx.options import InsertPosition, ParagraphOptions
from pydocx.paragraph import insert_paragraph
from pydocx.updater import Updater
from pydocx.xmlops import build_rpr_xml


def test_build_rpr_child_order() -> None:
    rpr = build_rpr_xml(
        True,
        True,
        True,
        font_family="Arial",
        font_size=28,
        font_color="00FF00",
        strikethrough=True,
        highlight="yellow",
        all_caps=True,
        small_caps=True,
    )
    order = [
        rpr.index("<w:rFonts"),
        rpr.index("<w:b/>"),
        rpr.index("<w:i/>"),
        rpr.index("<w:caps/>"),
        rpr.index("<w:smallCaps/>"),
        rpr.index("<w:strike/>"),
        rpr.index("<w:color"),
        rpr.index("<w:sz "),
        rpr.index("<w:highlight"),
        rpr.index("<w:u "),
    ]
    assert order == sorted(order)


def test_paragraph_applies_run_formatting(blank: Updater) -> None:
    insert_paragraph(
        blank.workspace,
        ParagraphOptions(
            text="styled",
            position=InsertPosition.END,
            font_color="123456",
            font_family="Georgia",
            font_size=24,
            highlight="green",
            strikethrough=True,
        ),
    )
    doc = (blank.workspace / "word" / "document.xml").read_text(encoding="utf-8")
    assert '<w:color w:val="123456"/>' in doc
    assert 'w:ascii="Georgia"' in doc
    assert '<w:sz w:val="24"/>' in doc
    assert '<w:highlight w:val="green"/>' in doc
    assert "<w:strike/>" in doc
