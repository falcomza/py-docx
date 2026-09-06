from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

from pydocx.options import EmbeddedObjectOptions, InsertPosition
from pydocx.updater import Updater


def _fake_xlsx() -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
    return buf.getvalue()


def test_insert_embedded_object(blank: Updater, tmp_path: Path) -> None:
    blank.insert_embedded_object(EmbeddedObjectOptions(file_bytes=_fake_xlsx(), position=InsertPosition.END))
    ws = blank.workspace
    assert (ws / "word" / "embeddings" / "embedding1.xlsx").exists()
    assert list((ws / "word" / "media").glob("image*.png"))

    doc = (ws / "word" / "document.xml").read_text(encoding="utf-8")
    assert '<o:OLEObject Type="Embed"' in doc
    assert 'DrawAspect="Icon"' in doc

    rels = (ws / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
    assert "embeddings/embedding1.xlsx" in rels
    ct = (ws / "[Content_Types].xml").read_text(encoding="utf-8")
    assert 'Extension="xlsx"' in ct

    out = tmp_path / "o.docx"
    blank.save(out)
    assert out.read_bytes()[:2] == b"PK"


def test_embed_requires_anchor_for_anchored_position(blank: Updater) -> None:
    import pytest

    with pytest.raises(ValueError):
        blank.insert_embedded_object(EmbeddedObjectOptions(file_bytes=_fake_xlsx(), position=InsertPosition.AFTER_TEXT))
