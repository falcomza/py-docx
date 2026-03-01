from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path


def create_blank_docx(workspace: Path) -> None:
    (workspace / "_rels").mkdir(parents=True, exist_ok=True)
    (workspace / "word" / "_rels").mkdir(parents=True, exist_ok=True)
    (workspace / "docProps").mkdir(parents=True, exist_ok=True)

    (workspace / "[Content_Types].xml").write_text(_content_types_xml(), encoding="utf-8")
    (workspace / "_rels" / ".rels").write_text(_package_rels_xml(), encoding="utf-8")
    (workspace / "word" / "_rels" / "document.xml.rels").write_text(_document_rels_xml(), encoding="utf-8")
    (workspace / "word" / "document.xml").write_text(_document_xml(), encoding="utf-8")
    (workspace / "word" / "styles.xml").write_text(_styles_xml(), encoding="utf-8")
    (workspace / "word" / "settings.xml").write_text(_settings_xml(), encoding="utf-8")
    (workspace / "docProps" / "core.xml").write_text(_core_xml(), encoding="utf-8")
    (workspace / "docProps" / "app.xml").write_text(_app_xml(), encoding="utf-8")


def _content_types_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
        '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
        '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
        '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\n'
        '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n'
        "</Types>"
    )


def _package_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n'
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\n'
        '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\n'
        "</Relationships>"
    )


def _document_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n'
        "</Relationships>"
    )


def _document_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        "  <w:body>\n"
        "    <w:p><w:r><w:t/></w:r></w:p>\n"
        '    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>\n'
        "  </w:body>\n"
        "</w:document>"
    )


def _styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">\n'
        '    <w:name w:val="Normal"/>\n'
        "    <w:qFormat/>\n"
        "  </w:style>\n"
        "</w:styles>"
    )


def _settings_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>\n'
    )


def _core_xml() -> str:
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
        "  <cp:revision>1</cp:revision>\n"
        f'  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>\n'
        f'  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>\n'
        "</cp:coreProperties>"
    )


def _app_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\n'
        "  <Application>Microsoft Word</Application>\n"
        "  <DocSecurity>0</DocSecurity>\n"
        "  <ScaleCrop>false</ScaleCrop>\n"
        "  <LinksUpToDate>false</LinksUpToDate>\n"
        "  <SharedDoc>false</SharedDoc>\n"
        "  <HyperlinksChanged>false</HyperlinksChanged>\n"
        "  <AppVersion>16.0000</AppVersion>\n"
        "</Properties>"
    )
