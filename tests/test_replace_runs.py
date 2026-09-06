from __future__ import annotations

from pathlib import Path

from pydocx.options import ReplaceOptions
from pydocx.replace import replace_text
from pydocx.xmlops import merge_adjacent_runs


def _doc(body: str) -> str:
    return (
        '<?xml version="1.0"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{body}</w:body></w:document>"
    )


def test_split_placeholder_is_replaced(tmp_path: Path) -> None:
    doc = tmp_path / "word" / "document.xml"
    doc.parent.mkdir(parents=True)
    doc.write_text(
        _doc("<w:p><w:r><w:t>Hello {{</w:t></w:r><w:r><w:t>name}} there</w:t></w:r></w:p>"),
        encoding="utf-8",
    )
    count = replace_text(tmp_path, "{{name}}", "World", ReplaceOptions())
    assert count == 1
    assert "Hello World there" in doc.read_text(encoding="utf-8")


def test_runs_with_different_rpr_not_merged() -> None:
    para = "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>foo</w:t></w:r><w:r><w:t>bar</w:t></w:r></w:p>"
    merged = merge_adjacent_runs(_doc(para))
    assert "<w:b/>" in merged
    assert merged.count("<w:r>") + merged.count("<w:r ") >= 1
    assert ">foobar<" not in merged


def test_runs_across_markup_boundary_not_merged() -> None:
    para = "<w:p><w:hyperlink><w:r><w:t>foo</w:t></w:r></w:hyperlink><w:r><w:t>bar</w:t></w:r></w:p>"
    merged = merge_adjacent_runs(_doc(para))
    assert ">foobar<" not in merged


def test_merge_preserves_leading_space() -> None:
    para = '<w:p><w:r><w:t xml:space="preserve"> a</w:t></w:r><w:r><w:t>b </w:t></w:r></w:p>'
    merged = merge_adjacent_runs(_doc(para))
    assert 'xml:space="preserve"> ab </w:t>' in merged
