from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path
from typing import BinaryIO

from .blank import create_blank_docx
from .bookmark import create_bookmark, create_bookmark_with_text, wrap_text_in_bookmark
from .breaks import insert_page_break, insert_section_break, set_page_layout
from .captions import update_caption_by_anchor, update_caption_in_document, update_caption_in_document_with_options
from .chart import insert_chart
from .chart_read import get_chart_data
from .chart_update import update_chart
from .comments import get_comments, insert_comment
from .count import get_chart_count, get_image_count, get_paragraph_count, get_table_count
from .delete import delete_chart, delete_image, delete_paragraphs, delete_table
from .errors import DocumentClosedError
from .header_footer import set_footer, set_header
from .hyperlink import insert_hyperlink, insert_internal_link
from .image import insert_image
from .notes import insert_endnote, insert_footnote
from .options import (
    AppProperties,
    BookmarkOptions,
    BreakOptions,
    CaptionOptions,
    CaptionType,
    ChartData,
    ChartOptions,
    Comment,
    CommentOptions,
    CoreProperties,
    CustomProperty,
    DeleteOptions,
    EndnoteOptions,
    FindOptions,
    FooterOptions,
    FootnoteOptions,
    HeaderFooterContent,
    HeaderOptions,
    HyperlinkOptions,
    ImageOptions,
    InsertPosition,
    ListType,
    PageLayoutOptions,
    PageNumberOptions,
    ParagraphOptions,
    ReplaceOptions,
    TableOptions,
    TextMatch,
    TOCOptions,
    TrackedDeleteOptions,
    TrackedInsertOptions,
    WatermarkOptions,
)
from .package_validator import validate_workspace
from .page_numbers import set_page_number
from .paragraph import add_heading, add_text, insert_paragraph, insert_paragraphs
from .properties import (
    get_app_properties,
    get_core_properties,
    get_custom_properties,
    set_app_properties,
    set_core_properties,
    set_custom_properties,
)
from .read import find_text, get_paragraph_text, get_table_text, get_text
from .replace import replace_text, replace_text_regex
from .table import insert_table
from .table_ops import merge_table_cells_horizontal, merge_table_cells_vertical, update_table_cell
from .toc import insert_toc
from .track_changes import delete_tracked_text, insert_tracked_text
from .watermark import set_text_watermark
from .ziputils import create_zip_from_dir, extract_zip


class Updater:
    def __init__(self, workspace: Path) -> None:
        self._workspace = workspace
        self._closed = False

    @property
    def workspace(self) -> Path:
        self._ensure_open()
        return self._workspace

    def save(self, output_path: str | Path) -> None:
        self._ensure_open()
        validate_workspace(self._workspace)
        create_zip_from_dir(self._workspace, output_path)

    def save_to_writer(self, writer: BinaryIO) -> None:
        self._ensure_open()
        validate_workspace(self._workspace)
        fd, tmp_str = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        tmp = Path(tmp_str)
        try:
            create_zip_from_dir(self._workspace, tmp)
            writer.write(tmp.read_bytes())
        finally:
            tmp.unlink(missing_ok=True)

    def __enter__(self) -> Updater:
        return self

    def __exit__(self, *_: object) -> None:
        self.cleanup()

    def cleanup(self) -> None:
        if self._closed:
            return
        shutil.rmtree(self._workspace, ignore_errors=True)
        self._closed = True

    def _ensure_open(self) -> None:
        if self._closed:
            raise DocumentClosedError("Document has been cleaned up")

    def insert_chart(self, opts: ChartOptions) -> None:
        self._ensure_open()
        insert_chart(self._workspace, opts)

    def update_chart(self, chart_index: int, data: ChartData) -> None:
        self._ensure_open()
        update_chart(self._workspace, chart_index, data)

    def get_chart_data(self, chart_index: int) -> ChartData:
        self._ensure_open()
        return get_chart_data(self._workspace, chart_index)

    def insert_paragraph(self, opts: ParagraphOptions) -> None:
        self._ensure_open()
        insert_paragraph(self._workspace, opts)

    def insert_paragraphs(self, paragraphs: list[ParagraphOptions]) -> None:
        self._ensure_open()
        insert_paragraphs(self._workspace, paragraphs)

    def add_heading(self, level: int, text: str, position: InsertPosition) -> None:
        self._ensure_open()
        add_heading(self._workspace, level, text, position)

    def add_text(self, text: str, position: InsertPosition) -> None:
        self._ensure_open()
        add_text(self._workspace, text, position)

    def add_bullet_item(self, text: str, level: int, position: InsertPosition) -> None:
        self._ensure_open()
        insert_paragraph(
            self._workspace,
            ParagraphOptions(text=text, list_type=ListType.BULLET,
                             list_level=level, position=position),
        )

    def add_numbered_item(self, text: str, level: int, position: InsertPosition) -> None:
        self._ensure_open()
        insert_paragraph(
            self._workspace,
            ParagraphOptions(text=text, list_type=ListType.NUMBERED,
                             list_level=level, position=position),
        )

    def add_bullet_list(self, items: list[str], level: int, position: InsertPosition) -> None:
        self._ensure_open()
        for item in items:
            insert_paragraph(
                self._workspace,
                ParagraphOptions(text=item, list_type=ListType.BULLET,
                                 list_level=level, position=position),
            )

    def add_numbered_list(self, items: list[str], level: int, position: InsertPosition) -> None:
        self._ensure_open()
        for item in items:
            insert_paragraph(
                self._workspace,
                ParagraphOptions(text=item, list_type=ListType.NUMBERED,
                                 list_level=level, position=position),
            )

    def insert_image(self, opts: ImageOptions) -> None:
        self._ensure_open()
        insert_image(self._workspace, opts)

    def set_header(self, content: HeaderFooterContent, opts: HeaderOptions, section_index: int = -1) -> None:
        self._ensure_open()
        set_header(self._workspace, content, opts, section_index=section_index)

    def set_footer(self, content: HeaderFooterContent, opts: FooterOptions, section_index: int = -1) -> None:
        self._ensure_open()
        set_footer(self._workspace, content, opts, section_index=section_index)

    def set_page_number(self, opts: PageNumberOptions) -> None:
        self._ensure_open()
        set_page_number(self._workspace, opts)

    def insert_toc(self, opts: TOCOptions) -> None:
        self._ensure_open()
        insert_toc(self._workspace, opts)

    def insert_footnote(self, opts: FootnoteOptions) -> None:
        self._ensure_open()
        insert_footnote(self._workspace, opts)

    def insert_endnote(self, opts: EndnoteOptions) -> None:
        self._ensure_open()
        insert_endnote(self._workspace, opts)

    def insert_comment(self, opts: CommentOptions) -> None:
        self._ensure_open()
        insert_comment(self._workspace, opts)

    def get_comments(self) -> list[Comment]:
        self._ensure_open()
        return get_comments(self._workspace)

    def insert_tracked_text(self, opts: TrackedInsertOptions) -> None:
        self._ensure_open()
        insert_tracked_text(self._workspace, opts)

    def delete_tracked_text(self, opts: TrackedDeleteOptions) -> None:
        self._ensure_open()
        delete_tracked_text(self._workspace, opts)

    def get_text(self) -> str:
        self._ensure_open()
        return get_text(self._workspace)

    def get_paragraph_text(self) -> list[str]:
        self._ensure_open()
        return get_paragraph_text(self._workspace)

    def get_table_text(self) -> list[list[list[str]]]:
        self._ensure_open()
        return get_table_text(self._workspace)

    def find_text(self, pattern: str, opts: FindOptions) -> list[TextMatch]:
        self._ensure_open()
        return find_text(self._workspace, pattern, opts)

    def replace_text(self, old: str, new: str, opts: ReplaceOptions) -> int:
        self._ensure_open()
        return replace_text(self._workspace, old, new, opts)

    def replace_text_regex(self, pattern: str, replacement: str, opts: ReplaceOptions) -> int:
        self._ensure_open()
        return replace_text_regex(self._workspace, pattern, replacement, opts)

    def delete_paragraphs(self, text: str, opts: DeleteOptions) -> int:
        self._ensure_open()
        return delete_paragraphs(self._workspace, text, opts)

    def delete_table(self, table_index: int) -> None:
        self._ensure_open()
        delete_table(self._workspace, table_index)

    def delete_image(self, image_index: int) -> None:
        self._ensure_open()
        delete_image(self._workspace, image_index)

    def delete_chart(self, chart_index: int) -> None:
        self._ensure_open()
        delete_chart(self._workspace, chart_index)

    def get_table_count(self) -> int:
        self._ensure_open()
        return get_table_count(self._workspace)

    def get_paragraph_count(self) -> int:
        self._ensure_open()
        return get_paragraph_count(self._workspace)

    def get_image_count(self) -> int:
        self._ensure_open()
        return get_image_count(self._workspace)

    def get_chart_count(self) -> int:
        self._ensure_open()
        return get_chart_count(self._workspace)

    def set_core_properties(self, props: CoreProperties) -> None:
        self._ensure_open()
        set_core_properties(self._workspace, props)

    def set_app_properties(self, props: AppProperties) -> None:
        self._ensure_open()
        set_app_properties(self._workspace, props)

    def set_custom_properties(self, properties: list[CustomProperty]) -> None:
        self._ensure_open()
        set_custom_properties(self._workspace, properties)

    def get_core_properties(self) -> CoreProperties:
        self._ensure_open()
        return get_core_properties(self._workspace)

    def get_app_properties(self) -> AppProperties:
        self._ensure_open()
        return get_app_properties(self._workspace)

    def get_custom_properties(self) -> list[CustomProperty]:
        self._ensure_open()
        return get_custom_properties(self._workspace)

    def insert_page_break(self, opts: BreakOptions) -> None:
        self._ensure_open()
        insert_page_break(self._workspace, opts)

    def insert_section_break(self, opts: BreakOptions) -> None:
        self._ensure_open()
        insert_section_break(self._workspace, opts)

    def set_page_layout(self, layout: PageLayoutOptions, section_index: int = -1) -> None:
        self._ensure_open()
        set_page_layout(self._workspace, layout, section_index=section_index)

    def set_text_watermark(self, opts: WatermarkOptions) -> None:
        self._ensure_open()
        set_text_watermark(self._workspace, opts)

    def insert_hyperlink(self, text: str, url: str, opts: HyperlinkOptions) -> None:
        self._ensure_open()
        insert_hyperlink(self._workspace, text, url, opts)

    def insert_internal_link(self, text: str, bookmark_name: str, opts: HyperlinkOptions) -> None:
        self._ensure_open()
        insert_internal_link(self._workspace, text, bookmark_name, opts)

    def create_bookmark(self, name: str, opts: BookmarkOptions) -> None:
        self._ensure_open()
        create_bookmark(self._workspace, name, opts)

    def create_bookmark_with_text(self, name: str, text: str, opts: BookmarkOptions) -> None:
        self._ensure_open()
        create_bookmark_with_text(self._workspace, name, text, opts)

    def wrap_text_in_bookmark(self, name: str, anchor_text: str) -> None:
        self._ensure_open()
        wrap_text_in_bookmark(self._workspace, name, anchor_text)

    def update_caption(self, caption_type: CaptionType, index: int, description: str) -> None:
        self._ensure_open()
        doc_path = self._workspace / "word" / "document.xml"
        doc_xml = doc_path.read_text(encoding="utf-8")
        updated = update_caption_in_document(
            doc_xml, caption_type, index, description)
        doc_path.write_text(updated, encoding="utf-8")

    def update_caption_with_options(self, caption_type: CaptionType, index: int, opts: CaptionOptions) -> None:
        self._ensure_open()
        doc_path = self._workspace / "word" / "document.xml"
        doc_xml = doc_path.read_text(encoding="utf-8")
        updated = update_caption_in_document_with_options(
            doc_xml, caption_type, index, opts)
        doc_path.write_text(updated, encoding="utf-8")

    def update_caption_by_anchor(
        self,
        anchor_text: str,
        caption_type: CaptionType,
        opts: CaptionOptions,
        direction: str = "after",
    ) -> None:
        self._ensure_open()
        doc_path = self._workspace / "word" / "document.xml"
        doc_xml = doc_path.read_text(encoding="utf-8")
        updated = update_caption_by_anchor(
            doc_xml, anchor_text, caption_type, opts, direction)
        doc_path.write_text(updated, encoding="utf-8")

    def insert_table(self, opts: TableOptions) -> None:
        self._ensure_open()
        insert_table(self._workspace, opts)

    def update_table_cell(self, table_index: int, row: int, col: int, value: str) -> None:
        self._ensure_open()
        update_table_cell(self._workspace, table_index, row, col, value)

    def merge_table_cells_horizontal(self, table_index: int, row: int, start_col: int, end_col: int) -> None:
        self._ensure_open()
        merge_table_cells_horizontal(
            self._workspace, table_index, row, start_col, end_col)

    def merge_table_cells_vertical(self, table_index: int, start_row: int, end_row: int, col: int) -> None:
        self._ensure_open()
        merge_table_cells_vertical(
            self._workspace, table_index, start_row, end_row, col)


def new(path: str | Path) -> Updater:
    workspace = Path(tempfile.mkdtemp(prefix="pydocx_"))
    extract_zip(path, workspace)
    return Updater(workspace)


def new_blank() -> Updater:
    workspace = Path(tempfile.mkdtemp(prefix="pydocx_blank_"))
    create_blank_docx(workspace)
    return Updater(workspace)


def new_from_bytes(data: bytes) -> Updater:
    workspace = Path(tempfile.mkdtemp(prefix="pydocx_bytes_"))
    tmp_docx = workspace / "input.docx"
    tmp_docx.write_bytes(data)
    try:
        extract_zip(tmp_docx, workspace)
    finally:
        tmp_docx.unlink(missing_ok=True)
    return Updater(workspace)


def new_from_reader(reader: BinaryIO) -> Updater:
    data = reader.read()
    return new_from_bytes(data)
