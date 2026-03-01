# API Reference

Complete method reference for the `Updater` class, organized by feature.

All methods raise `DocumentClosedError` if called after `cleanup()`.

---

## Entry Points

These module-level functions create an `Updater` instance:

### `new(path: str | Path) -> Updater`

Open an existing DOCX file from disk.

```python
u = pydocx.new("template.docx")
```

### `new_blank() -> Updater`

Create a new blank document (no template needed).

```python
u = pydocx.new_blank()
```

### `new_from_bytes(data: bytes) -> Updater`

Load a document from raw bytes.

```python
u = pydocx.new_from_bytes(docx_bytes)
```

### `new_from_reader(reader: BinaryIO) -> Updater`

Load a document from any binary stream.

```python
with open("doc.docx", "rb") as f:
    u = pydocx.new_from_reader(f)
```

---

## Save & Cleanup

### `save(output_path: str | Path) -> None`

Write the modified document to a file.

### `save_to_writer(writer: BinaryIO) -> None`

Write the modified document to a binary stream (e.g. `io.BytesIO`).

### `cleanup() -> None`

Delete the temporary workspace. Safe to call multiple times. **Must be called** when done.

---

## Paragraphs & Text

### `insert_paragraph(opts: ParagraphOptions) -> None`

Insert a single paragraph with full control over style, alignment, formatting, position, and list type.

```python
u.insert_paragraph(ParagraphOptions(
    text="Important note",
    bold=True,
    alignment=ParagraphAlignment.CENTER,
    position=InsertPosition.END,
))
```

### `insert_paragraphs(paragraphs: list[ParagraphOptions]) -> None`

Insert multiple paragraphs in order.

### `add_heading(level: int, text: str, position: InsertPosition) -> None`

Insert a heading paragraph. `level` must be 1, 2, or 3.

```python
u.add_heading(1, "Executive Summary", InsertPosition.END)
```

### `add_text(text: str, position: InsertPosition) -> None`

Insert a plain "Normal" style paragraph.

### `add_bullet_item(text: str, level: int, position: InsertPosition) -> None`

Insert a single bullet list item. `level` is 0-based nesting depth (0–8).

### `add_numbered_item(text: str, level: int, position: InsertPosition) -> None`

Insert a single numbered list item.

### `add_bullet_list(items: list[str], level: int, position: InsertPosition) -> None`

Insert multiple bullet items at once.

### `add_numbered_list(items: list[str], level: int, position: InsertPosition) -> None`

Insert multiple numbered items at once.

---

## Tables

### `insert_table(opts: TableOptions) -> None`

Insert a table with full styling support: column definitions, borders, colors, alignment, cell text styles, conditional formatting, row heights, captions, and more.

```python
u.insert_table(TableOptions(
    columns=[ColumnDefinition("Product"), ColumnDefinition("Sales")],
    rows=[["Widget A", "120"], ["Widget B", "98"]],
    header_background="4472C4",
    header_text_style=CellTextStyle(bold=True, font_color="FFFFFF"),
))
```

### `update_table_cell(table_index: int, row: int, col: int, value: str) -> None`

Update the text content of a specific cell. All indices are 1-based.

### `merge_table_cells_horizontal(table_index: int, row: int, start_col: int, end_col: int) -> None`

Merge cells horizontally within a row. All indices are 1-based.

### `merge_table_cells_vertical(table_index: int, start_row: int, end_row: int, col: int) -> None`

Merge cells vertically within a column. All indices are 1-based.

---

## Charts

### `insert_chart(opts: ChartOptions) -> None`

Insert a chart with an embedded Excel workbook. Supports column, bar, line, pie, area, and scatter chart types.

```python
u.insert_chart(ChartOptions(
    title="Quarterly Revenue",
    chart_kind=ChartKind.LINE,
    categories=["Q1", "Q2", "Q3", "Q4"],
    series=[
        SeriesOptions("2024", [100, 120, 110, 130]),
        SeriesOptions("2025", [130, 140, 135, 150], color="FF6600"),
    ],
))
```

### `update_chart(chart_index: int, data: ChartData) -> None`

Update the data of an existing chart by index (1-based). Updates both the chart XML cache and the embedded Excel workbook.

```python
u.update_chart(1, ChartData(
    categories=["Q1", "Q2"],
    series=[SeriesData("Revenue", [200, 250])],
))
```

### `get_chart_data(chart_index: int) -> ChartData`

Read the data from an existing chart. Returns categories, series names, and values. Reads from the embedded workbook first, falling back to the XML cache.

---

## Images

### `insert_image(opts: ImageOptions) -> None`

Insert an image into the document. Supported formats: PNG, JPEG, GIF, BMP.

- If both `width` and `height` are provided, the image is sized exactly.
- If only one dimension is provided, the other is calculated to maintain aspect ratio.
- If neither is provided, the image's native dimensions are used.

```python
u.insert_image(ImageOptions(
    path="logo.png",
    width=300,
    alt_text="Company Logo",
    position=InsertPosition.END,
))
```

---

## Headers & Footers

### `set_header(content: HeaderFooterContent, opts: HeaderOptions) -> None`

Set the document header. Content is laid out in a three-column table (left, center, right).

```python
u.set_header(
    HeaderFooterContent(left_text="Company", center_text="Confidential", right_text="2026"),
    HeaderOptions(type=HeaderType.DEFAULT),
)
```

### `set_footer(content: HeaderFooterContent, opts: FooterOptions) -> None`

Set the document footer with the same three-column layout.

---

## Page Numbers

### `set_page_number(opts: PageNumberOptions) -> None`

Set page numbering start value and format in the document's section properties.

```python
u.set_page_number(PageNumberOptions(start=1, format=PageNumberFormat.LOWER_ROMAN))
```

---

## Table of Contents

### `insert_toc(opts: TOCOptions) -> None`

Insert a Table of Contents field. The TOC is rendered by Word on document open (not by py-docx).

```python
u.insert_toc(TOCOptions(title="Contents", outline_levels="1-3"))
```

If `update_on_open` is `True` (default), Word will prompt to update the TOC when the file is opened.

---

## Page Layout & Breaks

### `insert_page_break(opts: BreakOptions) -> None`

Insert a page break at the specified position.

### `insert_section_break(opts: BreakOptions) -> None`

Insert a section break. Use `section_type` to control break behavior (`nextPage`, `continuous`, `evenPage`, `oddPage`). Optionally pass `page_layout` to change the page size/orientation for the new section.

```python
u.insert_section_break(BreakOptions(
    position=InsertPosition.END,
    section_type=SectionBreakType.NEXT_PAGE,
    page_layout=page_layout_letter_landscape(),
))
```

### `set_page_layout(layout: PageLayoutOptions) -> None`

Set the page size, orientation, and margins for the document's last section.

**Preset layout functions:**

| Function | Size |
|---|---|
| `page_layout_letter_portrait()` | 8.5 × 11 in |
| `page_layout_letter_landscape()` | 11 × 8.5 in |
| `page_layout_a4_portrait()` | 210 × 297 mm |
| `page_layout_a4_landscape()` | 297 × 210 mm |
| `page_layout_a3_portrait()` | 297 × 420 mm |
| `page_layout_a3_landscape()` | 420 × 297 mm |
| `page_layout_legal_portrait()` | 8.5 × 14 in |

---

## Watermarks

### `set_text_watermark(opts: WatermarkOptions) -> None`

Add a text watermark to the document (rendered via VML in the default header).

```python
u.set_text_watermark(WatermarkOptions(text="DRAFT", color="C0C0C0", opacity=0.5))
```

---

## Hyperlinks & Bookmarks

### `insert_hyperlink(text: str, url: str, opts: HyperlinkOptions) -> None`

Insert an external hyperlink.

```python
u.insert_hyperlink("Visit Python.org", "https://www.python.org", HyperlinkOptions())
```

### `insert_internal_link(text: str, bookmark_name: str, opts: HyperlinkOptions) -> None`

Insert an internal link that jumps to a named bookmark within the document.

### `create_bookmark(name: str, opts: BookmarkOptions) -> None`

Create an empty bookmark (invisible anchor point). Bookmark names must start with a letter, contain only letters/digits/underscores, and be ≤ 40 characters.

### `create_bookmark_with_text(name: str, text: str, opts: BookmarkOptions) -> None`

Create a bookmark that wraps visible text.

### `wrap_text_in_bookmark(name: str, anchor_text: str) -> None`

Wrap existing text in the document with a bookmark by searching for `anchor_text`.

---

## Footnotes & Endnotes

### `insert_footnote(opts: FootnoteOptions) -> None`

Insert a footnote anchored to text found by `opts.anchor`.

```python
u.insert_footnote(FootnoteOptions(text="See Appendix A.", anchor="significant results"))
```

### `insert_endnote(opts: EndnoteOptions) -> None`

Insert an endnote anchored to text found by `opts.anchor`.

---

## Comments

### `insert_comment(opts: CommentOptions) -> None`

Insert a comment anchored to text found by `opts.anchor`.

```python
u.insert_comment(CommentOptions(
    text="Please verify this number.",
    author="Reviewer",
    anchor="42% improvement",
))
```

### `get_comments() -> list[Comment]`

Read all comments from the document. Returns a list of `Comment` dataclasses with `id`, `author`, `initials`, `date`, and `text`.

---

## Track Changes

### `insert_tracked_text(opts: TrackedInsertOptions) -> None`

Insert text marked as a tracked insertion (revision mark).

```python
u.insert_tracked_text(TrackedInsertOptions(
    text="New content added by editor.",
    author="Editor",
    position=InsertPosition.END,
))
```

### `delete_tracked_text(opts: TrackedDeleteOptions) -> None`

Mark existing text as a tracked deletion by searching for `opts.anchor`.

---

## Captions

### `update_caption(caption_type: CaptionType, index: int, description: str) -> None`

Update the description text of an existing caption by type and index (1-based).

### `update_caption_with_options(caption_type: CaptionType, index: int, opts: CaptionOptions) -> None`

Update a caption with full control over style, alignment, and numbering.

### `update_caption_by_anchor(anchor_text: str, caption_type: CaptionType, opts: CaptionOptions, direction: str = "after") -> None`

Find a caption near the paragraph containing `anchor_text` and update it. `direction` is `"after"` or `"before"`.

> **Note:** Captions are inserted automatically when you pass a `CaptionOptions` to `insert_table()`, `insert_image()`, or `insert_chart()`.

---

## Text Search & Replace

### `get_text() -> str`

Read all text content from the document body as a single string.

### `get_paragraph_text() -> list[str]`

Read text from each paragraph as a list of strings.

### `get_table_text() -> list[list[list[str]]]`

Read table data as a 3D list: `tables[table_index][row_index][cell_index]`.

### `find_text(pattern: str, opts: FindOptions) -> list[TextMatch]`

Search for text in the document. Supports case sensitivity, whole word, regex, and searching in paragraphs, tables, headers, and footers.

```python
matches = u.find_text("revenue", FindOptions(match_case=False, whole_word=True))
for m in matches:
    print(f"Found '{m.text}' in paragraph {m.paragraph} at position {m.position}")
```

### `replace_text(old: str, new: str, opts: ReplaceOptions) -> int`

Replace text occurrences. Returns the number of replacements made. Supports case sensitivity, whole word matching, and scope control (paragraphs, tables, headers, footers).

### `replace_text_regex(pattern: str, replacement: str, opts: ReplaceOptions) -> int`

Replace text using a regular expression pattern. Returns the number of replacements made.

---

## Delete Operations

### `delete_paragraphs(text: str, opts: DeleteOptions) -> int`

Delete all paragraphs containing the specified text. Returns the count of deleted paragraphs.

### `delete_table(table_index: int) -> None`

Delete a table by index (1-based).

### `delete_image(image_index: int) -> None`

Delete an image by index (1-based).

### `delete_chart(chart_index: int) -> None`

Delete a chart by index (1-based).

---

## Count Operations

### `get_paragraph_count() -> int`

Count all paragraphs in the document.

### `get_table_count() -> int`

Count all tables in the document.

### `get_image_count() -> int`

Count all images in the document.

### `get_chart_count() -> int`

Count all charts in the document.

---

## Document Properties

### `set_core_properties(props: CoreProperties) -> None`

Set core metadata (title, creator, dates, revision, etc.).

### `get_core_properties() -> CoreProperties`

Read core metadata.

### `set_app_properties(props: AppProperties) -> None`

Set application metadata (company, manager, word count, etc.).

### `get_app_properties() -> AppProperties`

Read application metadata.

### `set_custom_properties(properties: list[CustomProperty]) -> None`

Set custom properties (replaces all existing custom properties).

### `get_custom_properties() -> list[CustomProperty]`

Read all custom properties.

```python
u.set_core_properties(CoreProperties(title="Q4 Report", creator="Finance Team"))
u.set_custom_properties([CustomProperty(name="Department", value="Finance")])

core = u.get_core_properties()
print(core.title)  # "Q4 Report"
```
