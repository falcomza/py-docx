# Project Knowledge — py-docx

## What This Is

A **stdlib-only** Python 3.12+ library for programmatically manipulating Microsoft Word (DOCX) documents. Python rework of the Go [`go-docx`](https://github.com/falcomza/go-docx) project. DOCX files are ZIP archives of XML; this library extracts, modifies XML via `xml.etree.ElementTree`, and re-packages.

## Project Layout

- `src/pydocx/` — all library source code
  - `updater.py` — core `Updater` class and entry points (`new`, `new_blank`, `new_from_bytes`, `new_from_reader`)
  - `options.py` — all dataclasses/enums for public API options
  - Feature modules: `table.py`, `chart.py`, `chart_read.py`, `chart_update.py`, `paragraph.py`, `image.py`, `captions.py`, `comments.py`, `notes.py`, `hyperlink.py`, `bookmark.py`, `header_footer.py`, `page_numbers.py`, `page_layout.py`, `breaks.py`, `watermark.py`, `toc.py`, `lists.py`, `track_changes.py`, `properties.py`, `read.py`, `replace.py`, `delete.py`, `count.py`, `table_ops.py`, `blank.py`
  - XML utilities: `xmlops.py`, `xmlutils.py`, `ziputils.py`, `rels.py`
  - `__init__.py` — public API re-exports
- `tests/` — test suite (currently minimal: single placeholder test)
- `examples/` — usage examples for each feature
- `docs/` — design docs, plans, and guides

## Commands

```bash
# Install (editable)
pip install -e .

# Run tests
python -m pytest tests/

# Lint
uvx ruff check src/ tests/ examples/

# Lint (auto-fix)
uvx ruff check --fix src/ tests/ examples/

# Format
uvx ruff format src/ tests/ examples/

# Format (check only)
uvx ruff format --check src/ tests/ examples/

# Serve docs locally (MkDocs + Material theme)
make docs
# or: uvx --with mkdocs-material mkdocs serve

# Build static docs site
make docs-build
# or: uvx --with mkdocs-material mkdocs build

# Clean docs build output
make docs-clean
```

## Key Conventions

- **Zero third-party dependencies** — stdlib only (`xml.etree.ElementTree`, `zipfile`, `tempfile`, etc.)
- **Python 3.12+** required
- Source lives under `src/` (setuptools `src`-layout)
- Public API is entirely re-exported from `src/pydocx/__init__.py`
- Options/config use `@dataclass` classes defined in `options.py`
- Entry points: `new(path)`, `new_blank()`, `new_from_bytes()`, `new_from_reader()` → return `Updater` instance
- `Updater` must be cleaned up with `.cleanup()` (extracts to temp dir)
- Build system: `setuptools` + `wheel` (see `pyproject.toml`)
- Documentation: MkDocs with Material theme (`mkdocs.yml`); docs live in `docs/`; build output goes to `site/`
- **Ruff** is used for linting and formatting (config in `pyproject.toml`)
- Line length: 120, double quotes, sorted imports

## Gotchas

- The test suite is a single placeholder — no real test coverage yet
- CI runs via GitHub Actions (`.github/workflows/ci.yml`)
- Package name is `py-docx` (hyphenated) but import is `pydocx` (no hyphen)
- All XML namespace handling is manual (OOXML namespaces)
