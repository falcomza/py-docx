from __future__ import annotations

import re
from pathlib import Path

from .options import CaptionOptions, CaptionPosition, CaptionType, CellAlignment
from .xmlutils import xml_escape


def validate_caption(opts: CaptionOptions) -> CaptionOptions:
    if opts.type not in (CaptionType.FIGURE, CaptionType.TABLE):
        raise ValueError("caption type must be Figure or Table")
    if opts.position is None:
        opts.position = CaptionPosition.BEFORE if opts.type == CaptionType.TABLE else CaptionPosition.AFTER
    if opts.position not in (CaptionPosition.BEFORE, CaptionPosition.AFTER):
        raise ValueError("caption position must be before or after")
    if len(opts.description) > 500:
        raise ValueError("caption description too long")
    return opts


def generate_caption_xml(opts: CaptionOptions) -> str:
    opts = validate_caption(opts)
    style = opts.style or "Caption"

    align = "left"
    if opts.alignment == CellAlignment.CENTER:
        align = "center"
    elif opts.alignment == CellAlignment.RIGHT:
        align = "right"

    parts = ["<w:p>"]
    parts.append("<w:pPr>")
    parts.append(f'<w:pStyle w:val="{xml_escape(style)}"/>')
    parts.append(f'<w:jc w:val="{align}"/>')
    parts.append("</w:pPr>")

    parts.append('<w:r><w:t xml:space="preserve">')
    parts.append(xml_escape(opts.type.value))
    parts.append(" </w:t></w:r>")

    if opts.auto_number:
        parts.append(_seq_field_xml(opts.type))
    elif opts.manual_number > 0:
        parts.append(f"<w:r><w:t>{opts.manual_number}</w:t></w:r>")

    if opts.description:
        parts.append('<w:r><w:t xml:space="preserve">: </w:t></w:r>')
        parts.append(f"<w:r><w:t>{xml_escape(opts.description)}</w:t></w:r>")

    parts.append("</w:p>")
    return "".join(parts)


def _seq_field_xml(caption_type: CaptionType) -> str:
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> SEQ {caption_type.value} \\* ARABIC </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>1</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def insert_caption_with_element(caption_xml: str, element_xml: str, position: CaptionPosition) -> str:
    if position == CaptionPosition.BEFORE:
        return caption_xml + element_xml
    return element_xml + caption_xml


def update_caption_in_document(doc_xml: str, caption_type: CaptionType, index: int, new_description: str) -> str:
    if index < 1:
        raise ValueError("caption index must be >= 1")
    paragraphs = list(re.finditer(r"<w:p>.*?</w:p>", doc_xml, flags=re.DOTALL))
    count = 0
    for match in paragraphs:
        para = match.group(0)
        if '<w:pStyle w:val="Caption"' not in para:
            continue
        if f">{caption_type.value} </w:t>" not in para:
            continue
        count += 1
        if count == index:
            replacement = generate_caption_xml(CaptionOptions(type=caption_type, description=new_description))
            return doc_xml[: match.start()] + replacement + doc_xml[match.end() :]
    raise ValueError(f"caption {caption_type.value} #{index} not found")


def update_caption_by_anchor(
    doc_xml: str,
    anchor_text: str,
    caption_type: CaptionType,
    opts: CaptionOptions,
    direction: CaptionPosition = CaptionPosition.AFTER,
) -> str:
    if not anchor_text:
        raise ValueError("anchor_text cannot be empty")
    if opts.type is None:
        opts.type = caption_type

    paragraphs = list(re.finditer(r"<w:p>.*?</w:p>", doc_xml, flags=re.DOTALL))
    anchor_idx = -1
    for i, match in enumerate(paragraphs):
        para = match.group(0)
        if anchor_text in para:
            anchor_idx = i
            break
    if anchor_idx == -1:
        raise ValueError(f"anchor text {anchor_text!r} not found in document")

    if direction == CaptionPosition.AFTER:
        search_range = range(anchor_idx + 1, len(paragraphs))
    else:
        search_range = range(anchor_idx - 1, -1, -1)

    for i in search_range:
        para = paragraphs[i].group(0)
        if '<w:pStyle w:val="Caption"' not in para:
            continue
        if f">{caption_type.value} </w:t>" not in para:
            continue
        replacement = generate_caption_xml(opts)
        return doc_xml[: paragraphs[i].start()] + replacement + doc_xml[paragraphs[i].end() :]

    raise ValueError(f"no {caption_type.value} caption found {direction} anchor")


def update_caption_in_document_with_options(
    doc_xml: str, caption_type: CaptionType, index: int, opts: CaptionOptions
) -> str:
    if index < 1:
        raise ValueError("caption index must be >= 1")
    if opts.type is None:
        opts.type = caption_type
    paragraphs = list(re.finditer(r"<w:p>.*?</w:p>", doc_xml, flags=re.DOTALL))
    count = 0
    for match in paragraphs:
        para = match.group(0)
        if '<w:pStyle w:val="Caption"' not in para:
            continue
        if f">{caption_type.value} </w:t>" not in para:
            continue
        count += 1
        if count == index:
            replacement = generate_caption_xml(opts)
            return doc_xml[: match.start()] + replacement + doc_xml[match.end() :]
    raise ValueError(f"caption {caption_type.value} #{index} not found")


def update_caption(workspace: Path, caption_type: CaptionType, index: int, description: str) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated = update_caption_in_document(doc_xml, caption_type, index, description)
    doc_path.write_text(updated, encoding="utf-8")


def update_caption_opts(workspace: Path, caption_type: CaptionType, index: int, opts: CaptionOptions) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated = update_caption_in_document_with_options(doc_xml, caption_type, index, opts)
    doc_path.write_text(updated, encoding="utf-8")


def update_caption_anchor(
    workspace: Path,
    anchor_text: str,
    caption_type: CaptionType,
    opts: CaptionOptions,
    direction: CaptionPosition = CaptionPosition.AFTER,
) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    updated = update_caption_by_anchor(doc_xml, anchor_text, caption_type, opts, direction)
    doc_path.write_text(updated, encoding="utf-8")
