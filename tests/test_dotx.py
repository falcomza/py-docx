from __future__ import annotations

import zipfile
from pathlib import Path

from pydocx.updater import new, new_blank

_TEMPLATE_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
_DOCUMENT_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"


def _make_dotx(tmp_path: Path) -> Path:
    src = new_blank()
    docx = tmp_path / "seed.docx"
    src.save(docx)
    src.cleanup()

    dotx = tmp_path / "template.dotx"
    with zipfile.ZipFile(docx) as zin, zipfile.ZipFile(dotx, "w") as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item == "[Content_Types].xml":
                data = data.replace(_DOCUMENT_CT.encode(), _TEMPLATE_CT.encode())
            zout.writestr(item, data)
    return dotx


def test_opening_dotx_promotes_content_type(tmp_path: Path) -> None:
    dotx = _make_dotx(tmp_path)
    u = new(dotx)
    try:
        ct = (u.workspace / "[Content_Types].xml").read_text(encoding="utf-8")
        assert _TEMPLATE_CT not in ct
        assert _DOCUMENT_CT in ct
        out = tmp_path / "out.docx"
        u.save(out)
        with zipfile.ZipFile(out) as zf:
            assert _TEMPLATE_CT.encode() not in zf.read("[Content_Types].xml")
    finally:
        u.cleanup()
