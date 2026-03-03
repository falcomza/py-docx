from __future__ import annotations

from pathlib import Path
from urllib.parse import urlparse

from .document import insert_at_body_end, insert_at_body_start
from .options import HyperlinkOptions, InsertPosition
from .rels import insert_relationship, next_relationship_id
from .xmlops import insert_after_anchor, insert_before_anchor, write_run_text
from .xmlutils import xml_escape


def insert_hyperlink(workspace: Path, text: str, url: str, opts: HyperlinkOptions) -> None:
    if not text:
        raise ValueError("hyperlink text cannot be empty")
    if not url:
        raise ValueError("hyperlink URL cannot be empty")
    _validate_url(url)

    rel_id = _add_hyperlink_relationship(workspace, url)
    hyperlink_xml = _build_external_hyperlink_xml(text, rel_id, opts)
    _insert_hyperlink(workspace, hyperlink_xml, opts)


def insert_internal_link(workspace: Path, text: str, bookmark_name: str, opts: HyperlinkOptions) -> None:
    if not text:
        raise ValueError("link text cannot be empty")
    if not bookmark_name:
        raise ValueError("bookmark name cannot be empty")

    hyperlink_xml = _build_internal_hyperlink_xml(text, bookmark_name, opts)
    _insert_hyperlink(workspace, hyperlink_xml, opts)


def _add_hyperlink_relationship(workspace: Path, url: str) -> str:
    rels_path = workspace / "word" / "_rels" / "document.xml.rels"
    rels_xml = rels_path.read_text(encoding="utf-8")
    rel_id = next_relationship_id(rels_xml)
    rel = (
        f'<Relationship Id="{rel_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        f'Target="{xml_escape(url)}" TargetMode="External"/>'
    )
    rels_path.write_text(insert_relationship(rels_xml, "\n  " + rel), encoding="utf-8")
    return rel_id


def _build_external_hyperlink_xml(text: str, rel_id: str, opts: HyperlinkOptions) -> str:
    return _build_hyperlink_xml(text, f'r:id="{rel_id}"', opts)


def _build_internal_hyperlink_xml(text: str, bookmark_name: str, opts: HyperlinkOptions) -> str:
    return _build_hyperlink_xml(text, f'w:anchor="{xml_escape(bookmark_name)}"', opts)


def _build_hyperlink_xml(text: str, link_attr: str, opts: HyperlinkOptions) -> str:
    p_pr = ""
    if opts.style and opts.style != "Normal":
        p_pr = f'<w:pPr><w:pStyle w:val="{xml_escape(opts.style)}"/></w:pPr>'

    tooltip = opts.tooltip or opts.screen_tip
    tooltip_attr = f' w:tooltip="{xml_escape(tooltip)}"' if tooltip else ""

    r_pr = ['<w:rStyle w:val="Hyperlink"/>']
    if opts.color:
        r_pr.append(f'<w:color w:val="{opts.color}"/>')
    if opts.underline:
        r_pr.append('<w:u w:val="single"/>')
    r_pr_xml = f"<w:rPr>{''.join(r_pr)}</w:rPr>"

    run_xml = f"<w:r>{r_pr_xml}{write_run_text(text)}</w:r>"
    return f"<w:p>{p_pr}<w:hyperlink {link_attr}{tooltip_attr}>{run_xml}</w:hyperlink></w:p>"


def _insert_hyperlink(workspace: Path, hyperlink_xml: str, opts: HyperlinkOptions) -> None:
    doc_path = workspace / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")

    if opts.position == InsertPosition.BEGINNING:
        updated = insert_at_body_start(doc_xml, hyperlink_xml)
    elif opts.position == InsertPosition.END:
        updated = insert_at_body_end(doc_xml, hyperlink_xml)
    elif opts.position == InsertPosition.AFTER_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for after_text insertion")
        updated = insert_after_anchor(doc_xml, hyperlink_xml, opts.anchor)
    elif opts.position == InsertPosition.BEFORE_TEXT:
        if not opts.anchor:
            raise ValueError("anchor text required for before_text insertion")
        updated = insert_before_anchor(doc_xml, hyperlink_xml, opts.anchor)
    else:
        raise ValueError(f"unsupported insert position: {opts.position}")

    doc_path.write_text(updated, encoding="utf-8")


def _validate_url(url: str) -> None:
    parsed = urlparse(url)
    if not parsed.scheme:
        if not (url.startswith("mailto:") or url.startswith("file:") or url.startswith("ftp:")):
            raise ValueError(f"invalid URL: {url}")


