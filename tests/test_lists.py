from __future__ import annotations

from pydocx.options import InsertPosition, ListType, ParagraphOptions
from pydocx.paragraph import insert_paragraph, insert_paragraphs
from pydocx.updater import Updater


def _numbering(u: Updater) -> str:
    return (u.workspace / "word" / "numbering.xml").read_text(encoding="utf-8")


def test_numbering_defines_all_nine_levels(blank: Updater) -> None:
    insert_paragraph(blank.workspace, ParagraphOptions(text="x", list_type=ListType.BULLET))
    xml = _numbering(blank)
    for lvl in range(9):
        assert f'<w:lvl w:ilvl="{lvl}">' in xml
    # both abstractNums populated
    assert xml.count('<w:lvl w:ilvl="8">') == 2


def test_nested_bullet_references_defined_level(blank: Updater) -> None:
    insert_paragraph(
        blank.workspace,
        ParagraphOptions(text="deep", list_type=ListType.BULLET, list_level=3),
    )
    doc = (blank.workspace / "word" / "document.xml").read_text(encoding="utf-8")
    assert '<w:ilvl w:val="3"/>' in doc


def test_restart_allocates_new_num_with_start_override(blank: Updater) -> None:
    insert_paragraphs(
        blank.workspace,
        [
            ParagraphOptions(text="a", list_type=ListType.NUMBERED),
            ParagraphOptions(text="b", list_type=ListType.NUMBERED, restart=True),
        ],
    )
    numbering = _numbering(blank)
    doc = (blank.workspace / "word" / "document.xml").read_text(encoding="utf-8")
    assert "<w:numRestart" not in doc
    assert '<w:startOverride w:val="1"/>' in numbering
    assert '<w:num w:numId="3">' in numbering
    assert '<w:numId w:val="3"/>' in doc


def test_restart_ignored_for_bullets(blank: Updater) -> None:
    insert_paragraph(
        blank.workspace,
        ParagraphOptions(text="x", list_type=ListType.BULLET, restart=True, position=InsertPosition.END),
    )
    doc = (blank.workspace / "word" / "document.xml").read_text(encoding="utf-8")
    assert '<w:numId w:val="1"/>' in doc
