from __future__ import annotations

from pydocx.delete import delete_paragraphs
from pydocx.options import DeleteMatchMode, DeleteOptions, InsertPosition, ParagraphOptions
from pydocx.paragraph import insert_paragraph
from pydocx.updater import Updater


def _fill(u: Updater, texts: list[str]) -> None:
    for t in texts:
        insert_paragraph(u.workspace, ParagraphOptions(text=t, position=InsertPosition.END))


def _count(u: Updater) -> int:
    doc = (u.workspace / "word" / "document.xml").read_text(encoding="utf-8")
    return doc.count("<w:p>") + doc.count("<w:p ")


def test_contains_mode_default(blank: Updater) -> None:
    _fill(blank, ["alpha item", "beta item", "gamma"])
    before = _count(blank)
    n = delete_paragraphs(blank.workspace, "item", DeleteOptions())
    assert n == 2
    assert _count(blank) == before - 2


def test_exact_mode(blank: Updater) -> None:
    _fill(blank, ["done", "done now", "done"])
    n = delete_paragraphs(blank.workspace, "done", DeleteOptions(mode=DeleteMatchMode.EXACT))
    assert n == 2


def test_regex_mode(blank: Updater) -> None:
    _fill(blank, ["ticket-12", "ticket-3", "note"])
    n = delete_paragraphs(blank.workspace, r"ticket-\d+", DeleteOptions(mode=DeleteMatchMode.REGEX))
    assert n == 2


def test_max_deletions_caps(blank: Updater) -> None:
    _fill(blank, ["row", "row", "row", "row"])
    n = delete_paragraphs(blank.workspace, "row", DeleteOptions(max_deletions=2))
    assert n == 2
