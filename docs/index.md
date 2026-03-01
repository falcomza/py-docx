# py-docx

A **stdlib-only** Python 3.12+ library for programmatically manipulating Microsoft Word (DOCX) documents.

[![CI](https://github.com/falcomza/py-docx/actions/workflows/ci.yml/badge.svg)](https://github.com/falcomza/py-docx/actions/workflows/ci.yml)
[![Python Version](https://img.shields.io/badge/Python-3.12%2B-3776AB?style=flat&logo=python)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## Overview

py-docx lets you create and modify Word documents from Python without any third-party dependencies. It works by extracting the DOCX ZIP archive, manipulating the underlying XML with `xml.etree.ElementTree`, and re-packaging the result.

## Features

- **Tables** — insert with full styling, borders, conditional formatting, merge cells
- **Charts** — column, bar, line, pie, area, scatter with embedded Excel workbooks
- **Paragraphs** — headings, lists (bullet/numbered), alignment, formatting
- **Images** — PNG, JPEG, GIF, BMP with auto-sizing and captions
- **Headers & Footers** — three-column layout with page numbers
- **Page Layout** — page size, orientation, margins, section breaks
- **Find & Replace** — case-sensitive, whole word, regex support
- **Track Changes** — tracked insertions and deletions
- **Comments, Footnotes, Endnotes**
- **Hyperlinks & Bookmarks**
- **Table of Contents**
- **Watermarks**
- **Document Properties** — core, app, and custom metadata

## Installation

```bash
pip install py-docx
```

Requires **Python 3.12+** — no third-party dependencies.

## Quick Start

```python
from pydocx import (
    ChartOptions,
    ColumnDefinition,
    InsertPosition,
    SeriesOptions,
    TableOptions,
    new,
)

u = new("template.docx")
try:
    u.add_heading(1, "Quarterly Report", InsertPosition.END)
    u.add_text("This quarter showed strong growth.", InsertPosition.END)

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
finally:
    u.cleanup()
```

## Next Steps

- [Getting Started](getting-started.md) — Core concepts, lifecycle, and entry points
- [API Reference](api-reference.md) — Full `Updater` method reference
- [Options Reference](options-reference.md) — All dataclasses, enums, and constants
