from __future__ import annotations

import re
import shutil
from pathlib import Path

from .captions import generate_caption_xml, insert_caption_with_element
from .document import insert_at_body_end, insert_at_body_start
from .options import ImageOptions, InsertPosition
from .rels import (
    IMAGE_CONTENT_TYPES,
    REL_IMAGE_TYPE,
    ensure_default_content_type,
    insert_relationship,
    next_relationship_id,
)
from .xmlutils import xml_escape

_IMAGE_RE = re.compile(r"^image(\d+)\.\w+$")
_DOCPR_RE = re.compile(r'wp:docPr id="(\d+)"')

EMUS_PER_INCH = 914400
DEFAULT_DPI = 96


def insert_image(workspace: Path, opts: ImageOptions) -> None:
    if not opts.path:
        raise ValueError("image path cannot be empty")
    src = Path(opts.path)
    if not src.exists():
        raise FileNotFoundError(f"image file not found: {src}")

    if opts.width > 0 and opts.height > 0:
        final_w, final_h = opts.width, opts.height
    else:
        actual_w, actual_h = _get_image_dimensions(src)
        final_w, final_h = _calculate_dimensions((actual_w, actual_h), opts.width, opts.height)

    image_index = _next_image_index(workspace)
    ext = src.suffix.lower()
    content_type = IMAGE_CONTENT_TYPES.get(ext, IMAGE_CONTENT_TYPES.get(".png"))

    media_dir = workspace / "word" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    dest_name = f"image{image_index}{ext}"
    shutil.copyfile(src, media_dir / dest_name)

    doc_rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    rels_xml = doc_rels_path.read_text(encoding="utf-8")
    rel_id = next_relationship_id(rels_xml)
    rel_xml = f'\n  <Relationship Id="{rel_id}" Type="{REL_IMAGE_TYPE}" Target="media/{dest_name}"/>'
    doc_rels_path.write_text(insert_relationship(rels_xml, rel_xml), encoding="utf-8")

    content_types_path = workspace / "[Content_Types].xml"
    ct_xml = content_types_path.read_text(encoding="utf-8")
    ct_xml = ensure_default_content_type(ct_xml, ext, content_type)
    content_types_path.write_text(ct_xml, encoding="utf-8")

    drawing_xml = _generate_image_drawing_xml(workspace, image_index, rel_id, final_w, final_h, opts.alt_text)
    if opts.caption is not None:
        if getattr(opts.caption, "type", None) is None:
            from .options import CaptionType

            opts.caption.type = CaptionType.FIGURE
        caption_xml = generate_caption_xml(opts.caption)
        drawing_xml = insert_caption_with_element(caption_xml, drawing_xml, opts.caption.position)
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    if opts.position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, drawing_xml)
    elif opts.position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, drawing_xml)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")
    doc_path.write_text(updated, encoding="utf-8")


def _next_image_index(workspace: Path) -> int:
    media_dir = workspace / "word" / "media"
    if not media_dir.exists():
        return 1
    max_idx = 0
    for entry in media_dir.iterdir():
        if entry.is_file():
            m = _IMAGE_RE.match(entry.name)
            if m:
                max_idx = max(max_idx, int(m.group(1)))
    return max_idx + 1


def _generate_image_drawing_xml(
    workspace: Path, image_index: int, rel_id: str, width_px: int, height_px: int, alt_text: str
) -> str:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    doc_pr_id = _next_docpr_id(doc_xml)
    width_emu = _pixels_to_emu(width_px)
    height_emu = _pixels_to_emu(height_px)
    if not alt_text:
        alt_text = f"Picture {image_index}"

    return (
        "<w:p><w:r><w:drawing>"
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{width_emu}" cy="{height_emu}"/>'
        f'<wp:docPr id="{doc_pr_id}" name="Picture {image_index}" descr="{xml_escape(alt_text)}"/>'
        "<wp:cNvGraphicFramePr>"
        '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        "</wp:cNvGraphicFramePr>"
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<pic:nvPicPr>"
        f'<pic:cNvPr id="{doc_pr_id}" name="Picture {image_index}" descr="{xml_escape(alt_text)}"/>'
        '<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr>'
        "</pic:nvPicPr>"
        "<pic:blipFill>"
        f'<a:blip r:embed="{rel_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        "<a:srcRect/><a:stretch><a:fillRect/></a:stretch>"
        "</pic:blipFill>"
        '<pic:spPr bwMode="auto">'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "<a:noFill/><a:ln><a:noFill/></a:ln>"
        "</pic:spPr>"
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )


def _next_docpr_id(doc_xml: str) -> int:
    max_id = 0
    for m in _DOCPR_RE.finditer(doc_xml):
        max_id = max(max_id, int(m.group(1)))
    return max_id + 1


def _pixels_to_emu(px: int) -> int:
    return int(px * EMUS_PER_INCH / DEFAULT_DPI)


def _calculate_dimensions(actual: tuple[int, int], req_w: int, req_h: int) -> tuple[int, int]:
    aw, ah = actual
    if req_w > 0 and req_h > 0:
        return req_w, req_h
    if req_w > 0:
        return req_w, int(req_w * ah / aw)
    if req_h > 0:
        return int(req_h * aw / ah), req_h
    return aw, ah


def _get_image_dimensions(path: Path) -> tuple[int, int]:
    ext = path.suffix.lower()
    if ext in (".png",):
        return _png_dimensions(path)
    if ext in (".jpg", ".jpeg"):
        return _jpeg_dimensions(path)
    if ext == ".gif":
        return _gif_dimensions(path)
    if ext == ".bmp":
        return _bmp_dimensions(path)
    if ext in (".tif", ".tiff"):
        raise ValueError("tiff dimensions not supported without width/height")
    raise ValueError("unsupported image format without width/height")


def _png_dimensions(path: Path) -> tuple[int, int]:
    with path.open("rb") as f:
        sig = f.read(8)
        if sig != b"\x89PNG\r\n\x1a\n":
            raise ValueError("invalid PNG signature")
        f.read(8)  # length + 'IHDR'
        data = f.read(8)
        w = int.from_bytes(data[0:4], "big")
        h = int.from_bytes(data[4:8], "big")
        return w, h


def _gif_dimensions(path: Path) -> tuple[int, int]:
    with path.open("rb") as f:
        header = f.read(10)
        if header[:3] != b"GIF":
            raise ValueError("invalid GIF header")
        w = int.from_bytes(header[6:8], "little")
        h = int.from_bytes(header[8:10], "little")
        return w, h


def _bmp_dimensions(path: Path) -> tuple[int, int]:
    with path.open("rb") as f:
        header = f.read(26)
        if header[:2] != b"BM":
            raise ValueError("invalid BMP header")
        w = int.from_bytes(header[18:22], "little")
        h = int.from_bytes(header[22:26], "little")
        return w, h


def _jpeg_dimensions(path: Path) -> tuple[int, int]:
    with path.open("rb") as f:
        if f.read(2) != b"\xff\xd8":
            raise ValueError("invalid JPEG header")
        while True:
            marker = f.read(2)
            if len(marker) < 2:
                break
            while marker[0] != 0xFF:
                marker = marker[1:] + f.read(1)
            if marker[1] in (0xC0, 0xC2):
                length = int.from_bytes(f.read(2), "big")
                data = f.read(length - 2)
                h = int.from_bytes(data[1:3], "big")
                w = int.from_bytes(data[3:5], "big")
                return w, h
            else:
                length = int.from_bytes(f.read(2), "big")
                f.read(length - 2)
    raise ValueError("could not determine JPEG dimensions")
