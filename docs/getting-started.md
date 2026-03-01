# Getting Started

## Installation

```bash
pip install py-docx
```

**Requirements:** Python 3.12+ — no third-party dependencies (stdlib only).

## Core Concepts

### How It Works

DOCX files are ZIP archives containing XML files. py-docx:

1. Extracts the DOCX archive to a temporary directory (the "workspace")
2. Parses and modifies XML files using `xml.etree.ElementTree`
3. Updates relationships (`_rels/*.rels`) and content types (`[Content_Types].xml`)
4. Re-packages everything into a new DOCX file

### The Updater

All document manipulation goes through the `Updater` class. You never instantiate it directly — instead, use one of the entry point functions:

| Function | Purpose |
|---|---|
| `new(path)` | Open an existing DOCX file |
| `new_blank()` | Create a new empty document |
| `new_from_bytes(data)` | Load from `bytes` |
| `new_from_reader(reader)` | Load from a `BinaryIO` stream |

### Lifecycle

Every `Updater` extracts to a temporary directory and **must be cleaned up** when you're done:

```python
from pydocx import new

u = new("template.docx")
try:
    # ... manipulate the document ...
    u.save("output.docx")
finally:
    u.cleanup()
```

- **`save(path)`** — Write the modified document to a file path.
- **`save_to_writer(writer)`** — Write to any `BinaryIO` object (e.g. `io.BytesIO`).
- **`cleanup()`** — Delete the temporary workspace. Safe to call multiple times.

> **Important:** Calling any method after `cleanup()` raises `DocumentClosedError`.

### Options Pattern

py-docx uses dataclasses for all configuration. Every feature has a corresponding `*Options` dataclass with sensible defaults:

```python
from pydocx import ParagraphOptions, InsertPosition

opts = ParagraphOptions(
    text="Hello, world!",
    bold=True,
    position=InsertPosition.END,
)
u.insert_paragraph(opts)
```

Required fields are positional; optional fields have defaults. See the [Options Reference](options-reference.md) for all available dataclasses.

### Insert Positions

Most insert methods accept a `position` parameter using the `InsertPosition` enum:

| Value | Behavior |
|---|---|
| `InsertPosition.BEGINNING` | Insert at the start of the document body |
| `InsertPosition.END` | Insert at the end of the document body (default) |
| `InsertPosition.AFTER_TEXT` | Insert after the paragraph containing `anchor` text |
| `InsertPosition.BEFORE_TEXT` | Insert before the paragraph containing `anchor` text |

When using `AFTER_TEXT` or `BEFORE_TEXT`, you must also set the `anchor` field on the options dataclass.

## Quick Start

### Create a blank document

```python
from pydocx import InsertPosition, new_blank

u = new_blank()
try:
    u.add_heading(1, "My Report", InsertPosition.END)
    u.add_text("This is the first paragraph.", InsertPosition.END)
    u.save("report.docx")
finally:
    u.cleanup()
```

### Open and modify an existing document

```python
from pydocx import InsertPosition, ReplaceOptions, new

u = new("existing.docx")
try:
    u.add_text("Appended paragraph.", InsertPosition.END)
    replaced = u.replace_text("old phrase", "new phrase", ReplaceOptions())
    u.save("modified.docx")
finally:
    u.cleanup()
```

### Load from bytes

```python
from pydocx import new_from_bytes

data = open("template.docx", "rb").read()
u = new_from_bytes(data)
try:
    # ... modify ...
    u.save("output.docx")
finally:
    u.cleanup()
```

## What's Next

- [API Reference](api-reference.md) — Full method reference for the `Updater` class
- [Options Reference](options-reference.md) — All dataclasses, enums, and constants
