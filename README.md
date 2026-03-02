# py-docx

[![CI](https://github.com/falcomza/py-docx/actions/workflows/ci.yml/badge.svg)](https://github.com/falcomza/py-docx/actions/workflows/ci.yml)
[![Python Version](https://img.shields.io/badge/Python-3.12%2B-3776AB?style=flat&logo=python)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A Python 3.12+ library for programmatically manipulating Microsoft Word (DOCX) documents using **stdlib only**. This project is a Python rework of the Go [`go-docx`](https://github.com/falcomza/go-docx) feature set, delivered incrementally.

## Status

All core features are implemented and stable. The library follows a **stdlib-only** design — no runtime dependencies. The public API is typed (`py.typed`, PEP 561 compliant) and every release is linted with `ruff` and type-checked with `pyright`.

## Installation

```bash
pip install py-docx
```

For development (includes `pytest`, `ruff`, and `pyright`):

```bash
pip install -e ".[dev]"
```

## Documentation

📖 **[Read the docs](https://falcomza.github.io/py-docx/)** — full documentation site powered by MkDocs

- [Getting Started](docs/getting-started.md) — Core concepts, lifecycle, and quick-start examples
- [API Reference](docs/api-reference.md) — Full method reference for the `Updater` class
- [Options Reference](docs/options-reference.md) — All dataclasses, enums, and constants

To serve the docs locally:

```bash
make docs          # or: uvx --with mkdocs-material mkdocs serve
```

### Lists

`ListType.BULLET` and `ListType.NUMBERED` are typed enum values (both are also accepted as plain strings for backwards compatibility).

```python
from pydocx import InsertPosition, ListType, ParagraphOptions, new

with new("template.docx") as u:
    # Convenience helpers
    u.add_bullet_list(["First item", "Second item", "Third item"], 0, InsertPosition.END)
    u.add_numbered_list(["Step 1", "Step 2", "Step 3"], 0, InsertPosition.END)

    # Full control via ParagraphOptions
    u.insert_paragraph(
        ParagraphOptions(
            text="Custom bullet",
            list_type=ListType.BULLET,
            list_level=1,
            position=InsertPosition.END,
        )
    )
    u.save("output.docx")
```

### Tables (read data)

```python
from pydocx import ColumnDefinition, TableOptions, new

with new("template.docx") as u:
    u.insert_table(
        TableOptions(
            columns=[ColumnDefinition("Product"), ColumnDefinition("Sales")],
            rows=[["Product A", "120"], ["Product B", "98"]],
        )
    )
    tables = u.get_table_text()
    print(tables[0])
    u.save("output.docx")
```

### Charts (read data)

```python
from pydocx import ChartOptions, SeriesOptions, new

with new("template.docx") as u:
    u.insert_chart(
        ChartOptions(
            title="Quarterly Revenue",
            categories=["Q1", "Q2", "Q3", "Q4"],
            series=[SeriesOptions("2025", [100, 120, 110, 130])],
        )
    )
    data = u.get_chart_data(1)
    print(data.categories, data.series[0].values)
    u.save("output.docx")
```

### Charts (extended)

```python
from pydocx import ChartKind, ChartOptions, DataLabelOptions, SeriesOptions, new

with new("template.docx") as u:
    u.insert_chart(
        ChartOptions(
            title="Market Share",
            chart_kind=ChartKind.PIE,
            categories=["Product A", "Product B", "Product C"],
            series=[SeriesOptions("Share", [45, 30, 25])],
            data_labels=DataLabelOptions(show_value=True, show_percent=True),
        )
    )
    u.save("output.docx")
```

### Watermarks

```python
from pydocx import WatermarkOptions, new

with new("template.docx") as u:
    u.set_text_watermark(WatermarkOptions(text="CONFIDENTIAL", opacity=0.4))
    u.save("output.docx")
```

### Page/section breaks and page layout

```python
from pydocx import BreakOptions, InsertPosition, SectionBreakType, new, page_layout_letter_landscape

with new("template.docx") as u:
    u.add_text("Page 1 content", InsertPosition.END)
    u.insert_page_break(BreakOptions(position=InsertPosition.END))
    u.add_text("Page 2 content", InsertPosition.END)
    u.insert_section_break(
        BreakOptions(
            position=InsertPosition.END,
            section_type=SectionBreakType.NEXT_PAGE,
            page_layout=page_layout_letter_landscape(),
        )
    )
    u.add_text("Landscape section content", InsertPosition.END)
    u.save("output.docx")
```

### Blank document creation

```python
from pydocx import InsertPosition, new_blank

with new_blank() as u:
    u.add_text("Hello from a blank document.", InsertPosition.END)
    u.save("output.docx")
```

### Document properties

```python
from datetime import datetime, timezone

from pydocx import AppProperties, CoreProperties, CustomProperty, new

with new("template.docx") as u:
    u.set_core_properties(
        CoreProperties(
            title="Q4 Report",
            creator="Finance Team",
            created=datetime(2026, 2, 1, tzinfo=timezone.utc),
            revision="2",
        )
    )
    u.set_app_properties(AppProperties(company="Acme Corp", manager="J. Smith"))
    u.set_custom_properties([CustomProperty(name="Department", value="Finance")])
    core = u.get_core_properties()
    print(core.title, core.creator, core.revision)
    u.save("output.docx")
```

### Count operations

```python
from pydocx import new

with new("template.docx") as u:
    print("Paragraphs:", u.get_paragraph_count())
    print("Tables:", u.get_table_count())
    print("Images:", u.get_image_count())
    print("Charts:", u.get_chart_count())
```

### Delete operations

```python
from pydocx import DeleteOptions, InsertPosition, new

with new("template.docx") as u:
    u.add_text("Delete me paragraph.", InsertPosition.END)
    u.add_text("Keep me paragraph.", InsertPosition.END)
    deleted = u.delete_paragraphs("Delete me", DeleteOptions(match_case=False, whole_word=False))
    print("Deleted", deleted)
    # u.delete_table(1)
    # u.delete_image(1)
    # u.delete_chart(1)
    u.save("output.docx")
```

### Find/replace and read

```python
from pydocx import FindOptions, InsertPosition, ReplaceOptions, new

with new("template.docx") as u:
    u.add_text("The quick brown fox jumps over the lazy dog.", InsertPosition.END)
    matches = u.find_text("quick", FindOptions(match_case=False, whole_word=True))
    for m in matches:
        print(m.text, m.paragraph, m.position)
    replaced = u.replace_text("brown", "red", ReplaceOptions(match_case=False))
    print("Replaced", replaced)
    all_text = u.get_text()
    u.save("output.docx")
```

### Track changes

```python
from pydocx import InsertPosition, TrackedDeleteOptions, TrackedInsertOptions, new

with new("template.docx") as u:
    u.insert_tracked_text(
        TrackedInsertOptions(
            text="This paragraph was added with track changes.",
            author="Editor",
            position=InsertPosition.END,
            bold=True,
        )
    )
    u.add_text("This sentence will be marked as deleted.", InsertPosition.END)
    u.delete_tracked_text(TrackedDeleteOptions(anchor="marked as deleted", author="Reviewer"))
    u.save("output.docx")
```

### Comments

```python
from pydocx import CommentOptions, InsertPosition, new

with new("template.docx") as u:
    u.add_text("This section needs review.", InsertPosition.END)
    u.insert_comment(
        CommentOptions(
            text="Please verify the numbers and sources.",
            author="Reviewer",
            anchor="needs review",
        )
    )
    comments = u.get_comments()
    for comment in comments:
        print(comment.id, comment.author, comment.text)
    u.save("output.docx")
```

### Captions (tables, charts, images)

```python
from pydocx import CaptionOptions, CaptionType, CellAlignment, ColumnDefinition, ImageOptions, TableOptions, new

with new("template.docx") as u:
    u.insert_table(
        TableOptions(
            columns=[ColumnDefinition("A"), ColumnDefinition("B")],
            rows=[["1", "2"]],
            caption=CaptionOptions(type=CaptionType.TABLE, description="Sample table"),
        )
    )
    u.insert_image(
        ImageOptions(
            path="examples/assets/tiny.png",
            width=100,
            height=100,
            caption=CaptionOptions(type=CaptionType.FIGURE, description="Sample image"),
        )
    )
    u.update_caption(CaptionType.TABLE, 1, "Updated table caption")
    u.update_caption_with_options(
        CaptionType.TABLE,
        1,
        CaptionOptions(
            type=CaptionType.TABLE,
            description="Styled caption",
            style="Caption",
            alignment=CellAlignment.CENTER,
            auto_number=True,
        ),
    )
    u.update_caption_by_anchor(
        anchor_text="Executive Summary",
        caption_type=CaptionType.TABLE,
        opts=CaptionOptions(type=CaptionType.TABLE, description="Caption near anchor"),
        direction="after",
    )
    u.save("output.docx")
```

### Footnotes / Endnotes

```python
from pydocx import EndnoteOptions, FootnoteOptions, InsertPosition, new

with new("template.docx") as u:
    u.add_text("The experiment showed significant results.", InsertPosition.END)
    u.insert_footnote(FootnoteOptions(text="Based on data collected in Q3 2026.", anchor="significant results"))
    u.insert_endnote(EndnoteOptions(text="See full methodology in Appendix A.", anchor="experiment"))
    u.save("output.docx")
```

### Hyperlinks & Bookmarks

```python
from pydocx import BookmarkOptions, HyperlinkOptions, InsertPosition, new

with new("template.docx") as u:
    u.create_bookmark_with_text(
        "executive_summary",
        "Executive Summary",
        BookmarkOptions(position=InsertPosition.END, style="Heading1"),
    )
    u.insert_internal_link(
        "Jump to Executive Summary",
        "executive_summary",
        HyperlinkOptions(position=InsertPosition.BEGINNING),
    )
    u.insert_hyperlink(
        "Open Python.org",
        "https://www.python.org",
        HyperlinkOptions(position=InsertPosition.END),
    )
    u.add_text("Key finding: 42% improvement in throughput.", InsertPosition.END)
    u.wrap_text_in_bookmark("key_finding", "Key finding")
    u.insert_internal_link(
        "See key finding",
        "key_finding",
        HyperlinkOptions(position=InsertPosition.END),
    )
    u.save("output.docx")
```

### Table of Contents

```python
from pydocx import TOCOptions, new

with new("template.docx") as u:
    u.insert_toc(TOCOptions(title="Table of Contents", outline_levels="1-3"))
    u.save("output.docx")
```

### Page numbers

```python
from pydocx import PageNumberFormat, PageNumberOptions, new

with new("template.docx") as u:
    u.set_page_number(PageNumberOptions(start=1, format=PageNumberFormat.DECIMAL))
    u.save("output.docx")
```

### Images

```python
from pydocx import ImageOptions, InsertPosition, new

with new("template.docx") as u:
    u.insert_image(ImageOptions(path="images/logo.png", width=400, position=InsertPosition.END))
    u.save("output.docx")
```

Supported formats: PNG, JPEG, GIF, BMP. For TIFF or unknown formats, pass both `width` and `height`.

### Headers / Footers

```python
from pydocx import FooterOptions, HeaderFooterContent, HeaderOptions, new

with new("template.docx") as u:
    u.set_header(
        HeaderFooterContent(left_text="Company Name", center_text="Confidential", right_text="Feb 2026"),
        HeaderOptions(),
    )
    u.set_footer(
        HeaderFooterContent(center_text="Page "),
        FooterOptions(),
    )
    u.save("output.docx")
```

### Paragraphs

Headings support levels 1–9 (`Heading1` through `Heading9`).

```python
from pydocx import InsertPosition, ParagraphAlignment, ParagraphOptions, new

with new("template.docx") as u:
    u.add_heading(1, "Executive Summary", InsertPosition.END)
    u.add_heading(3, "Sub-section", InsertPosition.END)  # levels 1-9 supported
    u.add_text("This quarter showed strong growth.", InsertPosition.END)
    u.insert_paragraph(
        ParagraphOptions(
            text="Important: Review required",
            alignment=ParagraphAlignment.CENTER,
            bold=True,
            italic=True,
            position=InsertPosition.END,
        )
    )
    u.insert_paragraph(
        ParagraphOptions(
            text="Line 1\nLine 2\tTabbed",
            position=InsertPosition.END,
        )
    )
    u.save("output.docx")
```

### Table styling/width/alignment

```python
from pydocx import (
    BorderStyle,
    CellAlignment,
    CellTextStyle,
    ColumnDefinition,
    ConditionalStyle,
    RowHeightRule,
    TableAlignment,
    TableOptions,
    TableWidthType,
    VerticalAlignment,
    new,
)

with new("template.docx") as u:
    u.insert_table(
        TableOptions(
            columns=[
                ColumnDefinition("Product", width=2400, alignment=CellAlignment.LEFT),
                ColumnDefinition("Sales", width=1600, alignment=CellAlignment.RIGHT),
            ],
            rows=[["Product A", "120"], ["Product B", "98"]],
            table_alignment=TableAlignment.CENTER,
            table_width_type=TableWidthType.FIXED,
            table_width=4000,
            header_background="4472C4",
            row_background="FFFFFF",
            alternate_row_background="F2F2F2",
            border_style=BorderStyle.SINGLE,
            border_color="000000",
            vertical_alignment=VerticalAlignment.CENTER,
            header_repeat=True,
            header_row_height=360,
            header_height_rule=RowHeightRule.EXACT,
            row_height=280,
            row_height_rule=RowHeightRule.AT_LEAST,
            header_text_style=CellTextStyle(bold=True, font_color="FFFFFF", font_size=22),
            row_text_style=CellTextStyle(font_size=20),
            cell_text_styles={
                (2, 2): CellTextStyle(bold=True, italic=True, font_color="C00000"),
            },
            conditional_styles={
                "overdue": ConditionalStyle(
                    text_style=CellTextStyle(bold=True, font_color="C00000"),
                    background="FFF2CC",
                )
            },
            paragraph_style="Normal",
            proportional_column_widths=True,
            available_width=9360,
        )
    )
    u.save("output.docx")
```

## Quick Start

`Updater` implements the context manager protocol — use `with` to ensure the temp directory is always cleaned up:

```python
from pydocx import ChartOptions, ColumnDefinition, SeriesOptions, TableOptions, new

with new("template.docx") as u:
    u.insert_table(
        TableOptions(
            columns=[ColumnDefinition("Product"), ColumnDefinition("Sales")],
            rows=[["Product A", "120"], ["Product B", "98"]],
        )
    )
    u.insert_chart(
        ChartOptions(
            title="Quarterly Revenue",
            categories=["Q1", "Q2", "Q3", "Q4"],
            series=[SeriesOptions("2025", [100, 120, 110, 130])],
        )
    )
    u.save("output.docx")
```

The `try/finally` pattern also works and remains supported for backwards compatibility:

```python
u = new("template.docx")
try:
    u.add_text("Hello.", InsertPosition.END)
    u.save("output.docx")
finally:
    u.cleanup()
```

## Roadmap (tracked)

- [x] Core IO (open/extract/save/cleanup)
- [x] Charts: insert (column/bar), update existing charts
- [x] Charts: extended types (line/pie/area/scatter), advanced options
- [x] Charts: read chart data
- [x] Tables: insert, update cells, merge cells
- [x] Tables: delete/count/read operations
- [x] Paragraphs: insert, headings (levels 1–9), alignment, anchor-based positioning
- [x] Paragraphs: lists (bullet/numbered) with typed `ListType` enum
- [x] Images: insert with sizing and captions
- [x] TOC, headers/footers, page numbers
- [x] Captions: insert/edit with alignment/style and anchor-based edit
- [x] Footnotes/endnotes
- [x] Hyperlinks & bookmarks
- [x] Comments
- [x] Track changes
- [x] Text find/replace and read operations
- [x] Delete operations (paragraphs/tables/images/charts)
- [x] Count operations (paragraphs/tables/images/charts)
- [x] Document properties (core/app/custom CRUD)
- [x] Blank document creation (true NewBlank)
- [x] Page/section breaks and page layout
- [x] Watermarks
- [x] Context manager protocol on `Updater` (`with new(...) as u:`)
- [x] Typed public API with `py.typed` / PEP 561 compliance

## Requirements

- Python 3.12+
- Standard library only (no third-party runtime dependencies)
- Fully typed: ships a `py.typed` marker (PEP 561) — type checkers (`pyright`, `mypy`) understand all public APIs out of the box

## How It Works

DOCX files are ZIP archives containing XML files. This library:
1. Extracts the DOCX archive to a temporary directory
2. Parses and modifies XML files using `xml.etree.ElementTree`
3. Updates relationships (`_rels/*.rels`) and content types (`[Content_Types].xml`)
4. Re-packages everything into a new DOCX file

## License

MIT — see [LICENSE](LICENSE).
