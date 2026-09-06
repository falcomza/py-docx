# Changelog

## Unreleased — go-docx parity pass

Brings `pydocx` to feature parity with the reference Go implementation
(`go-docx`), plus fixes to list numbering and find-and-replace.

### New features

| API | Description |
|---|---|
| `Updater.add_style(defn)` / `add_styles(defns)` | Add custom paragraph or character styles (`StyleDefinition`, `StyleType`). Reference the id from `ParagraphOptions.style`. |
| `Updater.force_field_update_on_open()` | Make Word/LibreOffice recalculate every field code (PAGE, NUMPAGES, DATE, TOC, cross-references, SEQ, …) on first open. |
| `Updater.insert_table_of_figures(opts)` / `insert_table_of_tables(opts)` | Insert a `TOC \c "Figure"` / `"Table"` field (`CaptionListOptions`). |
| `Updater.get_toc_entries()` | Parse level + text of an existing table of contents (`TOCEntry`). |
| `Updater.update_toc()` | Mark TOC / caption-list fields dirty so Word rebuilds them on open. |
| `Updater.insert_embedded_object(opts)` | Embed an OLE object (e.g. an `.xlsx` workbook) shown as a clickable icon (`EmbeddedObjectOptions`; a built-in Excel icon is used when none is supplied). |
| `ParagraphOptions` run formatting | New fields: `font_family`, `font_size` (half-points), `font_color`, `strikethrough`, `highlight`, `all_caps`, `small_caps`. |
| `DeleteOptions.mode` / `max_deletions` | `DeleteMatchMode.CONTAINS` (default) / `EXACT` / `REGEX`, and a deletion cap. |
| `CaptionType.EQUATION` | Equation captions (`SEQ Equation`). |
| `.dotx` input | `new()` / `new_from_bytes()` auto-promote a template's main content type to the document type, so opened templates can be edited and re-saved. |

### Fixes

- **Find & replace across runs** — `replace_text` / `replace_text_regex` now merge
  adjacent `<w:r>` runs before matching, so placeholders that Word split across
  runs (a common cause of "the replacement didn't happen") are matched. The merge
  only runs on a retry, when a plain pass finds nothing.
- **Nested list numbering** — `numbering.xml` now defines all nine levels for both
  bullet and numbered lists (previously only level 0, so `list_level > 0` items
  referenced an undefined level).
- **Numbered-list restart** — `ParagraphOptions(restart=True)` now allocates a
  fresh numbering instance with a `<w:startOverride>` (the previous
  `<w:numRestart>` element is not valid OOXML).
- **Tracked deletions** — `delete_tracked_text` referenced an undefined helper and
  raised `NameError`; fixed.
