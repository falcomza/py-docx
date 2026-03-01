from __future__ import annotations

import os
import zipfile
from pathlib import Path

import pytest

from pydocx.chart_read import _resolve_workbook_for_chart as resolve_chart_read_workbook
from pydocx.chart_update import (
    _resolve_workbook_for_chart as resolve_chart_update_workbook,
)
from pydocx.chart_update import _update_title, _update_worksheet
from pydocx.errors import InvalidDocxError
from pydocx.options import ChartData, FindOptions, ReplaceOptions, SeriesData
from pydocx.read import find_text
from pydocx.replace import replace_text
from pydocx.updater import new_from_bytes
from pydocx.ziputils import extract_zip


def test_new_from_bytes_does_not_embed_source_artifacts(tmp_path: Path) -> None:
    template = Path(__file__).resolve().parents[1] / "templates" / "docx_template.docx"
    updater = new_from_bytes(template.read_bytes())
    output = tmp_path / "saved.docx"
    try:
        updater.save(output)
    finally:
        updater.cleanup()

    with zipfile.ZipFile(output, "r") as zf:
        names = set(zf.namelist())

    assert "input.docx" not in names
    assert "source.docx" not in names


def test_extract_zip_rejects_path_traversal(tmp_path: Path) -> None:
    archive = tmp_path / "unsafe.docx"
    output_dir = tmp_path / "out"
    output_dir.mkdir()

    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("../escape.txt", "bad")

    with pytest.raises(InvalidDocxError, match="unsafe path"):
        extract_zip(archive, output_dir)

    assert not (tmp_path / "escape.txt").exists()


def test_update_worksheet_replaces_sheet_data_and_dimension() -> None:
    raw = (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<dimension ref="A1:A1"/>'
        b'<sheetData><row r="1"><c r="A1"><v>OLD</v></c></row></sheetData>'
        b"</worksheet>"
    )
    data = ChartData(
        categories=["Q1", "Q2"],
        series=[SeriesData(name="Revenue", values=[10.0, 20.0])],
    )

    updated = _update_worksheet(raw, data, use_shared_strings=False, string_indexes={}).decode("utf-8")

    assert "OLD" not in updated
    assert '<dimension ref="A1:B3"/>' in updated
    assert '<sheetData><row r="1">' in updated
    assert "Revenue" in updated


def test_update_title_replaces_content_without_corrupting_xml() -> None:
    xml = (
        "<c:chart><c:title><c:tx><c:rich><c:p><c:r><c:t>Old Title</c:t></c:r></c:p></c:rich></c:tx>"
        "</c:title><c:plotArea/></c:chart>"
    )

    updated = _update_title(xml, "New Title", "c:")

    assert "<c:t>New Title</c:t>" in updated
    assert "Old Title" not in updated
    assert "</c:title><c:plotArea/></c:chart>" in updated


def test_chart_workbook_resolution_rejects_target_outside_workspace(tmp_path: Path) -> None:
    workspace = tmp_path / "workspace"
    chart_dir = workspace / "word" / "charts"
    rels_dir = chart_dir / "_rels"
    rels_dir.mkdir(parents=True)

    outside = tmp_path / "outside.xlsx"
    outside.write_bytes(b"not-an-xlsx")
    rel_target = os.path.relpath(outside, chart_dir)
    rels_xml = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Target="{rel_target}"/>'
        "</Relationships>"
    )
    (rels_dir / "chart1.xml.rels").write_text(rels_xml, encoding="utf-8")
    chart_xml = '<c:chartSpace><c:externalData r:id="rId1"/></c:chartSpace>'

    with pytest.raises(ValueError, match="escapes workspace"):
        resolve_chart_read_workbook(workspace, 1, chart_xml)
    with pytest.raises(ValueError, match="escapes workspace"):
        resolve_chart_update_workbook(workspace, 1, chart_xml)


def test_find_text_respects_paragraph_vs_table_scope(tmp_path: Path) -> None:
    workspace = tmp_path / "ws"
    word_dir = workspace / "word"
    word_dir.mkdir(parents=True)
    (word_dir / "document.xml").write_text(
        (
            "<w:document><w:body>"
            "<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>"
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Beta</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            "</w:body></w:document>"
        ),
        encoding="utf-8",
    )

    in_paragraphs = find_text(workspace, "Alpha", FindOptions(in_paragraphs=True, in_tables=False))
    in_tables = find_text(workspace, "Beta", FindOptions(in_paragraphs=False, in_tables=True))
    excluded = find_text(workspace, "Beta", FindOptions(in_paragraphs=True, in_tables=False))

    assert len(in_paragraphs) == 1
    assert len(in_tables) == 1
    assert excluded == []


def test_replace_text_respects_paragraph_vs_table_scope(tmp_path: Path) -> None:
    workspace = tmp_path / "ws"
    word_dir = workspace / "word"
    word_dir.mkdir(parents=True)
    doc_path = word_dir / "document.xml"
    doc_path.write_text(
        (
            "<w:document><w:body>"
            "<w:p><w:r><w:t>Target outside</w:t></w:r></w:p>"
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Target inside</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            "</w:body></w:document>"
        ),
        encoding="utf-8",
    )

    replaced = replace_text(workspace, "Target", "Replaced", ReplaceOptions(in_paragraphs=True, in_tables=False))
    updated = doc_path.read_text(encoding="utf-8")

    assert replaced == 1
    assert "Replaced outside" in updated
    assert "Target inside" in updated
