from __future__ import annotations

import os
import zipfile
from pathlib import Path

import pytest

from pydocx.breaks import set_page_layout
from pydocx.chart_read import _resolve_workbook_for_chart as resolve_chart_read_workbook
from pydocx.chart_update import (
    _resolve_workbook_for_chart as resolve_chart_update_workbook,
)
from pydocx.chart_update import _update_title, _update_worksheet
from pydocx.errors import InvalidDocxError, InvalidPackageError
from pydocx.header_footer import set_header
from pydocx.options import (
    ChartData,
    FindOptions,
    HeaderFooterContent,
    HeaderOptions,
    HeaderType,
    InsertPosition,
    PageLayoutOptions,
    PageOrientation,
    ParagraphOptions,
    ReplaceOptions,
    SeriesData,
)
from pydocx.paragraph import insert_paragraph
from pydocx.read import find_text
from pydocx.replace import replace_text
from pydocx.updater import new_blank, new_from_bytes
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


def test_set_page_layout_preserves_existing_section_properties() -> None:
    updater = new_blank()
    try:
        workspace = updater.workspace
        doc_path = workspace / "word" / "document.xml"
        doc_path.write_text(
            (
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                "<w:body>"
                "<w:p><w:r><w:t>hello</w:t></w:r></w:p>"
                '<w:sectPr><w:pgNumType w:start="3"/><w:headerReference w:type="default" r:id="rId9"/>'
                '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
                'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
                "</w:body></w:document>"
            ),
            encoding="utf-8",
        )
        set_page_layout(
            workspace,
            PageLayoutOptions(
                page_width=11906,
                page_height=16838,
                orientation=PageOrientation.PORTRAIT,
                margin_top=1000,
                margin_right=1100,
                margin_bottom=1200,
                margin_left=1300,
                margin_header=700,
                margin_footer=800,
                margin_gutter=0,
            ),
        )
        updated = doc_path.read_text(encoding="utf-8")
    finally:
        updater.cleanup()

    assert '<w:headerReference w:type="default" r:id="rId9"/>' in updated
    assert '<w:pgNumType w:start="3"/>' in updated
    assert '<w:pgSz w:w="11906" w:h="16838"/>' in updated
    assert (
        '<w:pgMar w:top="1000" w:right="1100" w:bottom="1200" '
        'w:left="1300" w:header="700" w:footer="800" w:gutter="0"/>'
    ) in updated


def test_set_header_odd_even_writes_settings_and_not_sectpr() -> None:
    updater = new_blank()
    try:
        workspace = updater.workspace
        set_header(
            workspace,
            HeaderFooterContent(left_text="L"),
            HeaderOptions(type=HeaderType.DEFAULT, different_odd_even=True),
        )
        doc_xml = (workspace / "word" / "document.xml").read_text(encoding="utf-8")
        settings_xml = (workspace / "word" / "settings.xml").read_text(encoding="utf-8")
    finally:
        updater.cleanup()

    assert "<w:evenAndOddHeaders" not in doc_xml
    assert "<w:evenAndOddHeaders/>" in settings_xml


def test_set_header_targets_last_section_properties() -> None:
    updater = new_blank()
    try:
        workspace = updater.workspace
        doc_path = workspace / "word" / "document.xml"
        doc_path.write_text(
            (
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                "<w:body>"
                '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>'
                "<w:r><w:t>one</w:t></w:r></w:p>"
                "<w:p><w:r><w:t>two</w:t></w:r></w:p>"
                '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
                "</w:body></w:document>"
            ),
            encoding="utf-8",
        )
        set_header(
            workspace,
            HeaderFooterContent(center_text="C"),
            HeaderOptions(type=HeaderType.DEFAULT),
        )
        updated = doc_path.read_text(encoding="utf-8")
    finally:
        updater.cleanup()

    assert updated.count("<w:headerReference") == 1
    assert updated.rfind("<w:headerReference") > updated.rfind("<w:sectPr")


def test_new_blank_has_settings_and_styles_relationships() -> None:
    updater = new_blank()
    try:
        rels_xml = (updater.workspace / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
    finally:
        updater.cleanup()

    assert 'relationships/styles" Target="styles.xml"' in rels_xml
    assert 'relationships/settings" Target="settings.xml"' in rels_xml


def test_insert_paragraph_anchor_matches_unescaped_text() -> None:
    updater = new_blank()
    try:
        workspace = updater.workspace
        doc_path = workspace / "word" / "document.xml"
        doc_path.write_text(
            (
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                "<w:body>"
                "<w:p><w:r><w:t>A &amp; B</w:t></w:r></w:p>"
                '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
                "</w:body></w:document>"
            ),
            encoding="utf-8",
        )
        insert_paragraph(
            workspace,
            ParagraphOptions(text="Inserted", position=InsertPosition.AFTER_TEXT, anchor="A & B"),
        )
        updated = doc_path.read_text(encoding="utf-8")
    finally:
        updater.cleanup()

    assert "<w:t>Inserted</w:t>" in updated


def test_save_rejects_dangling_relationship_target(tmp_path: Path) -> None:
    updater = new_blank()
    output = tmp_path / "out.docx"
    try:
        rels_path = updater.workspace / "word" / "_rels" / "document.xml.rels"
        rels_xml = rels_path.read_text(encoding="utf-8")
        rels_xml = rels_xml.replace(
            "</Relationships>",
            (
                '<Relationship Id="rId999" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
                'Target="missing-header.xml"/></Relationships>'
            ),
            1,
        )
        rels_path.write_text(rels_xml, encoding="utf-8")

        with pytest.raises(InvalidPackageError, match="missing target"):
            updater.save(output)
    finally:
        updater.cleanup()


def test_set_header_with_section_index_targets_first_section() -> None:
    updater = new_blank()
    try:
        workspace = updater.workspace
        doc_path = workspace / "word" / "document.xml"
        doc_path.write_text(
            (
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                "<w:body>"
                '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>'
                "<w:r><w:t>one</w:t></w:r></w:p>"
                "<w:p><w:r><w:t>two</w:t></w:r></w:p>"
                '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
                "</w:body></w:document>"
            ),
            encoding="utf-8",
        )
        set_header(
            workspace,
            HeaderFooterContent(center_text="C"),
            HeaderOptions(type=HeaderType.DEFAULT),
            section_index=0,
        )
        updated = doc_path.read_text(encoding="utf-8")
    finally:
        updater.cleanup()

    first_sectpr_end = updated.find("</w:sectPr>")
    header_ref_pos = updated.find("<w:headerReference")
    assert header_ref_pos != -1
    assert header_ref_pos < first_sectpr_end


def test_set_page_layout_with_section_index_targets_first_section() -> None:
    updater = new_blank()
    try:
        workspace = updater.workspace
        doc_path = workspace / "word" / "document.xml"
        doc_path.write_text(
            (
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                "<w:body>"
                '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>'
                "<w:r><w:t>one</w:t></w:r></w:p>"
                "<w:p><w:r><w:t>two</w:t></w:r></w:p>"
                '<w:sectPr><w:pgSz w:w="12000" w:h="15000"/></w:sectPr>'
                "</w:body></w:document>"
            ),
            encoding="utf-8",
        )
        set_page_layout(
            workspace,
            PageLayoutOptions(
                page_width=11111,
                page_height=22222,
                orientation=PageOrientation.PORTRAIT,
                margin_top=1000,
                margin_right=1000,
                margin_bottom=1000,
                margin_left=1000,
                margin_header=700,
                margin_footer=700,
                margin_gutter=0,
            ),
            section_index=0,
        )
        updated = doc_path.read_text(encoding="utf-8")
    finally:
        updater.cleanup()

    assert '<w:pgSz w:w="11111" w:h="22222"/>' in updated
    assert '<w:pgSz w:w="12000" w:h="15000"/>' in updated
