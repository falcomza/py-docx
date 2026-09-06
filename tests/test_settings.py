from __future__ import annotations

from pathlib import Path
from xml.dom.minidom import parseString

from pydocx.settings import force_field_update_on_open
from pydocx.updater import Updater


def _settings(u: Updater) -> str:
    return (u.workspace / "word" / "settings.xml").read_text(encoding="utf-8")


def test_existing_settings_gets_flag(blank: Updater) -> None:
    # new_blank ships an empty <w:settings></w:settings> root
    force_field_update_on_open(blank.workspace)
    text = _settings(blank)
    assert '<w:updateFields w:val="1"/>' in text
    parseString(text)  # still well-formed


def test_idempotent_when_already_truthy(blank: Updater) -> None:
    force_field_update_on_open(blank.workspace)
    force_field_update_on_open(blank.workspace)
    assert _settings(blank).count("<w:updateFields") == 1


def test_falsy_value_replaced(blank: Updater) -> None:
    sp = blank.workspace / "word" / "settings.xml"
    content = sp.read_text(encoding="utf-8")
    sp.write_text(content.replace("</w:settings>", '<w:updateFields w:val="0"/></w:settings>'), encoding="utf-8")
    force_field_update_on_open(blank.workspace)
    text = _settings(blank)
    assert '<w:updateFields w:val="1"/>' in text
    assert 'w:val="0"' not in text


def test_created_and_wired_when_absent(blank: Updater, tmp_path: Path) -> None:
    (blank.workspace / "word" / "settings.xml").unlink()
    ct = blank.workspace / "[Content_Types].xml"
    ct.write_text(
        ct.read_text(encoding="utf-8").replace(
            '<Override PartName="/word/settings.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
            "",
        ),
        encoding="utf-8",
    )
    rels = blank.workspace / "word" / "_rels" / "document.xml.rels"
    import re

    rels.write_text(
        re.sub(r"<Relationship[^>]*settings[^>]*/>", "", rels.read_text(encoding="utf-8")), encoding="utf-8"
    )

    force_field_update_on_open(blank.workspace)
    assert '<w:updateFields w:val="1"/>' in _settings(blank)
    assert "settings.xml" in rels.read_text(encoding="utf-8")
    assert "/word/settings.xml" in ct.read_text(encoding="utf-8")
    blank.save(tmp_path / "o.docx")


def test_header_footer_fields_marked_dirty(blank: Updater) -> None:
    word = blank.workspace / "word"
    (word / "header1.xml").write_text(
        '<w:hdr xmlns:w="x"><w:p><w:r>'
        '<w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>'
        "</w:p></w:hdr>",
        encoding="utf-8",
    )
    force_field_update_on_open(blank.workspace)
    hdr = (word / "header1.xml").read_text(encoding="utf-8")
    assert hdr.count('w:dirty="true"') == 2
