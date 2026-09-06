from __future__ import annotations

from pydocx.options import CaptionListOptions, InsertPosition
from pydocx.toc import get_toc_entries, insert_table_of_figures, insert_table_of_tables, update_toc
from pydocx.updater import Updater


def _doc(u: Updater) -> str:
    return (u.workspace / "word" / "document.xml").read_text(encoding="utf-8")


def test_table_of_figures_field(blank: Updater) -> None:
    insert_table_of_figures(
        blank.workspace, CaptionListOptions(caption_label="Figure", title="Figures", position=InsertPosition.BEGINNING)
    )
    doc = _doc(blank)
    assert 'TOC \\h \\z \\c "Figure"' in doc
    assert ">Figures<" in doc


def test_table_of_tables_field(blank: Updater) -> None:
    insert_table_of_tables(blank.workspace, CaptionListOptions(caption_label="Table"))
    assert 'TOC \\h \\z \\c "Table"' in _doc(blank)


def test_get_toc_entries(blank: Updater) -> None:
    doc_path = blank.workspace / "word" / "document.xml"
    doc = doc_path.read_text(encoding="utf-8")
    inject = (
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Chapter One</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:r><w:t>Section 1.1</w:t></w:r></w:p>'
    )
    doc_path.write_text(doc.replace("</w:body>", inject + "</w:body>"), encoding="utf-8")
    entries = get_toc_entries(blank.workspace)
    assert [(e.level, e.text) for e in entries] == [(1, "Chapter One"), (2, "Section 1.1")]


def test_update_toc_marks_dirty(blank: Updater) -> None:
    doc_path = blank.workspace / "word" / "document.xml"
    doc = doc_path.read_text(encoding="utf-8")
    doc_path.write_text(
        doc.replace("</w:body>", '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p></w:body>'),
        encoding="utf-8",
    )
    update_toc(blank.workspace)
    assert 'w:dirty="true"' in doc_path.read_text(encoding="utf-8")
    assert '<w:updateFields w:val="1"/>' in (blank.workspace / "word" / "settings.xml").read_text(encoding="utf-8")
