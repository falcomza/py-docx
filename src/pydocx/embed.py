from __future__ import annotations

import re
from importlib import resources
from pathlib import Path

from .document import insert_at_body_end, insert_at_body_start
from .image import _next_docpr_id, _next_image_index
from .options import EmbeddedObjectOptions, InsertPosition
from .rels import (
    IMAGE_CONTENT_TYPES,
    REL_IMAGE_TYPE,
    REL_PACKAGE_TYPE,
    ensure_default_content_type,
    insert_relationship,
    next_relationship_id,
)
from .xmlops import insert_after_anchor, insert_before_anchor

_XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_VML_NS = "urn:schemas-microsoft-com:vml"
_OFFICE_NS = "urn:schemas-microsoft-com:office:office"
_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_EMBED_RE = re.compile(r"^embedding(\d+)\.\w+$")


def insert_embedded_object(workspace: Path, opts: EmbeddedObjectOptions) -> None:
    if opts.position in (InsertPosition.AFTER_TEXT, InsertPosition.BEFORE_TEXT) and not opts.anchor:
        raise ValueError(f"anchor text required for position {opts.position}")

    file_bytes = _resolve_file_bytes(opts)
    icon_bytes = _resolve_icon_bytes(opts)

    emb_idx = _next_embedding_index(workspace)
    img_idx = _next_image_index(workspace)

    emb_dir = workspace / "word" / "embeddings"
    emb_dir.mkdir(parents=True, exist_ok=True)
    xlsx_name = f"embedding{emb_idx}.xlsx"
    (emb_dir / xlsx_name).write_bytes(file_bytes)

    media_dir = workspace / "word" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    icon_name = f"image{img_idx}.png"
    (media_dir / icon_name).write_bytes(icon_bytes)

    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    xlsx_rel_id = next_relationship_id(rels_xml)
    rels_xml = insert_relationship(
        rels_xml,
        f'\n  <Relationship Id="{xlsx_rel_id}" Type="{REL_PACKAGE_TYPE}" Target="embeddings/{xlsx_name}"/>',
    )
    image_rel_id = next_relationship_id(rels_xml)
    rels_xml = insert_relationship(
        rels_xml,
        f'\n  <Relationship Id="{image_rel_id}" Type="{REL_IMAGE_TYPE}" Target="media/{icon_name}"/>',
    )
    rels_path.write_text(rels_xml, encoding="utf-8")

    ct_path = workspace / "[Content_Types].xml"
    ct_xml = ct_path.read_text(encoding="utf-8")
    ct_xml = ensure_default_content_type(ct_xml, ".xlsx", _XLSX_CONTENT_TYPE)
    ct_xml = ensure_default_content_type(ct_xml, ".png", IMAGE_CONTENT_TYPES[".png"])
    ct_path.write_text(ct_xml, encoding="utf-8")

    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    shape_id = f"_x0000_i{_next_docpr_id(doc_xml)}"
    object_id = f"_{1000000000 + emb_idx}"
    fragment = _ole_object_xml(
        shape_id, image_rel_id, xlsx_rel_id, opts.prog_id, object_id, opts.width_pt, opts.height_pt
    )

    if opts.position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, fragment)
    elif opts.position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, fragment)
    elif opts.position == InsertPosition.AFTER_TEXT:
        updated = insert_after_anchor(doc_xml, fragment, opts.anchor)
    elif opts.position == InsertPosition.BEFORE_TEXT:
        updated = insert_before_anchor(doc_xml, fragment, opts.anchor)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")
    doc_path.write_text(updated, encoding="utf-8")


def _ole_object_xml(
    shape_id: str, image_rel_id: str, xlsx_rel_id: str, prog_id: str, object_id: str, width_pt: int, height_pt: int
) -> str:
    return (
        f'<w:p><w:r><w:object w:dxaOrig="{width_pt * 20}" w:dyaOrig="{height_pt * 20}">'
        f'<v:shape id="{shape_id}" type="#_x0000_t75"'
        f' style="width:{width_pt}pt;height:{height_pt}pt" o:ole=""'
        f' xmlns:v="{_VML_NS}" xmlns:o="{_OFFICE_NS}">'
        f'<v:imagedata r:id="{image_rel_id}" o:title="" xmlns:r="{_RELS_NS}"/>'
        "</v:shape>"
        f'<o:OLEObject Type="Embed" ProgID="{prog_id}" ShapeID="{shape_id}"'
        f' DrawAspect="Icon" ObjectID="{object_id}" r:id="{xlsx_rel_id}"'
        f' xmlns:o="{_OFFICE_NS}" xmlns:r="{_RELS_NS}"/>'
        "</w:object></w:r></w:p>"
    )


def _resolve_file_bytes(opts: EmbeddedObjectOptions) -> bytes:
    if opts.file_bytes:
        return opts.file_bytes
    if not opts.file_path:
        raise ValueError("one of file_path or file_bytes must be provided")
    return Path(opts.file_path).read_bytes()


def _resolve_icon_bytes(opts: EmbeddedObjectOptions) -> bytes:
    if opts.icon_bytes:
        return opts.icon_bytes
    if opts.icon_path:
        try:
            return Path(opts.icon_path).read_bytes()
        except OSError:
            pass
    return (resources.files("pydocx.assets") / "excel_icon.png").read_bytes()


def _next_embedding_index(workspace: Path) -> int:
    directory = workspace / "word" / "embeddings"
    if not directory.exists():
        return 1
    indices = [int(m.group(1)) for e in directory.iterdir() if (m := _EMBED_RE.match(e.name))]
    return max(indices, default=0) + 1
