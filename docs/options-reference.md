# Options Reference

All dataclasses, enums, and constants exported by `pydocx`. Every type listed here is importable directly from the `pydocx` package.

---

## Enums

### `InsertPosition`

Controls where content is placed in the document.

| Value | Description |
|---|---|
| `BEGINNING` | Start of the document body |
| `END` | End of the document body |
| `AFTER_TEXT` | After the paragraph containing the `anchor` text |
| `BEFORE_TEXT` | Before the paragraph containing the `anchor` text |

### `ParagraphAlignment`

| Value | OOXML |
|---|---|
| `LEFT` | `left` |
| `CENTER` | `center` |
| `RIGHT` | `right` |
| `JUSTIFY` | `both` |

### `TableAlignment`

Horizontal alignment of the table within the page.

| Value | Description |
|---|---|
| `LEFT` | Left-aligned |
| `CENTER` | Centered |
| `RIGHT` | Right-aligned |

### `CellAlignment`

Horizontal text alignment within a table cell.

| Value | OOXML |
|---|---|
| `LEFT` | `start` |
| `CENTER` | `center` |
| `RIGHT` | `end` |

### `VerticalAlignment`

Vertical text alignment within a table cell.

| Value | Description |
|---|---|
| `TOP` | Top-aligned |
| `CENTER` | Vertically centered |
| `BOTTOM` | Bottom-aligned |

### `TableWidthType`

How the table width value is interpreted.

| Value | Description |
|---|---|
| `AUTO` | Automatic sizing |
| `PERCENTAGE` | Width as fiftieths of a percent (5000 = 100%) |
| `FIXED` | Width in twips (1/1440 inch) |

### `BorderStyle`

| Value | Description |
|---|---|
| `NONE` | No borders |
| `SINGLE` | Single line |
| `DOUBLE` | Double line |
| `DOTTED` | Dotted line |
| `DASHED` | Dashed line |

### `RowHeightRule`

| Value | Description |
|---|---|
| `AUTO` | Automatic height |
| `AT_LEAST` | Minimum height (can grow) |
| `EXACT` | Fixed height (content clipped) |

### `ChartKind`

| Value | Description |
|---|---|
| `COLUMN` | Vertical bar chart |
| `BAR` | Horizontal bar chart |
| `LINE` | Line chart |
| `PIE` | Pie chart |
| `AREA` | Area chart |
| `SCATTER` | Scatter (XY) chart |

### `HeaderType` / `FooterType`

| Value | Description |
|---|---|
| `FIRST` | First page only |
| `EVEN` | Even pages |
| `DEFAULT` | Default (odd pages, or all pages if no first/even) |

### `PageNumberFormat`

| Value | Example |
|---|---|
| `DECIMAL` | 1, 2, 3 |
| `UPPER_ROMAN` | I, II, III |
| `LOWER_ROMAN` | i, ii, iii |
| `UPPER_LETTER` | A, B, C |
| `LOWER_LETTER` | a, b, c |

### `PageOrientation`

| Value | Description |
|---|---|
| `PORTRAIT` | Tall orientation |
| `LANDSCAPE` | Wide orientation |

### `SectionBreakType`

| Value | Description |
|---|---|
| `NEXT_PAGE` | New section starts on the next page |
| `CONTINUOUS` | New section continues on the same page |
| `EVEN_PAGE` | New section starts on the next even page |
| `ODD_PAGE` | New section starts on the next odd page |

### `CaptionType`

| Value | Description |
|---|---|
| `FIGURE` | Figure caption (for images and charts) |
| `TABLE` | Table caption |

### `CaptionPosition`

| Value | Description |
|---|---|
| `BEFORE` | Caption placed before the element |
| `AFTER` | Caption placed after the element |

---

## Dataclasses — Paragraphs & Text

### `ParagraphOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `text` | `str` | *(required)* | Paragraph text content. Supports `\n` (line break) and `\t` (tab). |
| `style` | `str` | `"Normal"` | Word paragraph style name |
| `alignment` | `ParagraphAlignment \| None` | `None` | Horizontal alignment |
| `position` | `InsertPosition` | `END` | Where to insert |
| `anchor` | `str` | `""` | Anchor text for `AFTER_TEXT`/`BEFORE_TEXT` |
| `bold` | `bool` | `False` | Bold formatting |
| `italic` | `bool` | `False` | Italic formatting |
| `underline` | `bool` | `False` | Underline formatting |
| `list_type` | `str` | `""` | `"bullet"` or `"numbered"` (empty = no list) |
| `list_level` | `int` | `0` | Nesting level (0–8) |

---

## Dataclasses — Tables

### `TableOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `columns` | `list[ColumnDefinition]` | *(required)* | Column headers and widths |
| `rows` | `list[list[str]]` | *(required)* | Row data (list of cell values per row) |
| `position` | `InsertPosition` | `END` | Where to insert |
| `table_style` | `str` | `"TableGrid"` | Word table style |
| `header_bold` | `bool` | `True` | Bold header row text |
| `header_repeat` | `bool` | `False` | Repeat header on page breaks |
| `header_text_style` | `CellTextStyle` | *(default)* | Text style for header cells |
| `row_text_style` | `CellTextStyle` | *(default)* | Text style for data cells |
| `cell_text_styles` | `dict[tuple[int,int], CellTextStyle] \| None` | `None` | Per-cell overrides; keys are `(row, col)` 1-based |
| `conditional_styles` | `dict[str, ConditionalStyle] \| None` | `None` | Styles applied when cell text matches a key |
| `paragraph_style` | `str` | `""` | Paragraph style for cell content |
| `header_background` | `str` | `""` | Hex RGB color for header row |
| `row_background` | `str` | `""` | Hex RGB color for data rows |
| `alternate_row_background` | `str` | `""` | Hex RGB color for odd data rows |
| `table_alignment` | `TableAlignment` | `LEFT` | Table horizontal alignment |
| `header_alignment` | `CellAlignment` | `LEFT` | Default header cell text alignment |
| `row_alignment` | `CellAlignment` | `LEFT` | Default data cell text alignment |
| `vertical_alignment` | `VerticalAlignment` | `CENTER` | Vertical cell alignment |
| `table_width_type` | `TableWidthType` | `PERCENTAGE` | How `table_width` is interpreted |
| `table_width` | `int` | `5000` | Table width (5000 = 100% when `PERCENTAGE`) |
| `column_widths` | `list[int] \| None` | `None` | Explicit column widths in twips |
| `proportional_column_widths` | `bool` | `False` | Auto-calculate widths from content length |
| `available_width` | `int` | `0` | Page content width in twips (for proportional) |
| `border_style` | `BorderStyle` | `SINGLE` | Border line style |
| `border_size` | `int` | `4` | Border thickness (eighth-points) |
| `border_color` | `str` | `"000000"` | Hex RGB border color |
| `cell_padding` | `int` | `108` | Cell padding in twips |
| `header_row_height` | `int` | `0` | Header row height in twips (0 = auto) |
| `header_height_rule` | `RowHeightRule` | `AUTO` | Header height rule |
| `row_height` | `int` | `0` | Data row height in twips (0 = auto) |
| `row_height_rule` | `RowHeightRule` | `AUTO` | Data row height rule |
| `caption` | `CaptionOptions \| None` | `None` | Optional table caption |

### `ColumnDefinition`

| Field | Type | Default | Description |
|---|---|---|---|
| `title` | `str` | *(required)* | Column header text |
| `width` | `int \| None` | `None` | Column width in twips |
| `bold` | `bool` | `False` | Bold the header cell |
| `alignment` | `CellAlignment \| None` | `None` | Override alignment for this column |

### `CellTextStyle`

| Field | Type | Default | Description |
|---|---|---|---|
| `bold` | `bool` | `False` | Bold text |
| `italic` | `bool` | `False` | Italic text |
| `font_size` | `int \| None` | `None` | Font size in half-points (e.g. `20` = 10pt) |
| `font_color` | `str \| None` | `None` | Hex RGB color |

### `ConditionalStyle`

| Field | Type | Default | Description |
|---|---|---|---|
| `text_style` | `CellTextStyle` | *(default)* | Text formatting to apply |
| `background` | `str \| None` | `None` | Cell background color |

---

## Dataclasses — Charts

### `ChartOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `categories` | `list[str]` | *(required)* | Category labels (X-axis) |
| `series` | `list[SeriesOptions]` | *(required)* | Data series |
| `title` | `str` | `""` | Chart title |
| `chart_kind` | `ChartKind \| str` | `COLUMN` | Chart type |
| `title_overlay` | `bool` | `False` | Title overlaps plot area |
| `category_axis_title` | `str` | `""` | Category axis label |
| `value_axis_title` | `str` | `""` | Value axis label |
| `show_legend` | `bool` | `True` | Show the legend |
| `legend_position` | `str` | `"r"` | Legend position (`r`, `l`, `t`, `b`, `tr`) |
| `width_emu` | `int` | `6099523` | Chart width in EMUs |
| `height_emu` | `int` | `3340467` | Chart height in EMUs |
| `position` | `InsertPosition` | `END` | Where to insert |
| `caption` | `CaptionOptions \| None` | `None` | Optional caption |
| `category_axis` | `AxisOptions \| None` | `None` | Category axis options (auto-populated if `None`) |
| `value_axis` | `AxisOptions \| None` | `None` | Value axis options (auto-populated if `None`) |
| `legend` | `LegendOptions \| None` | `None` | Legend options (auto-populated if `None`) |
| `data_labels` | `DataLabelOptions \| None` | `None` | Data label options |
| `properties` | `ChartProperties \| None` | `None` | Global chart properties |
| `bar_chart_options` | `BarChartOptions \| None` | `None` | Bar/column-specific options |
| `scatter_chart_options` | `ScatterChartOptions \| None` | `None` | Scatter-specific options |

### `SeriesOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `name` | `str` | *(required)* | Series name (legend label) |
| `values` | `list[float]` | *(required)* | Data values (must match `categories` length) |
| `x_values` | `list[float] \| None` | `None` | X values for scatter charts |
| `color` | `str` | `""` | Hex RGB fill color |
| `invert_if_negative` | `bool` | `False` | Invert fill for negative values (bar only) |
| `smooth` | `bool` | `False` | Smooth lines (line/scatter only) |
| `show_markers` | `bool` | `True` | Show data point markers (line/scatter) |
| `data_labels` | `DataLabelOptions \| None` | `None` | Per-series data labels |

### `ChartData`

Used for reading and updating chart data.

| Field | Type | Default | Description |
|---|---|---|---|
| `categories` | `list[str]` | *(required)* | Category labels |
| `series` | `list[SeriesData]` | *(required)* | Series data |
| `chart_title` | `str` | `""` | Chart title |
| `category_axis_title` | `str` | `""` | Category axis title |
| `value_axis_title` | `str` | `""` | Value axis title |

### `SeriesData`

| Field | Type | Default | Description |
|---|---|---|---|
| `name` | `str` | *(required)* | Series name |
| `values` | `list[float]` | *(required)* | Data values |
| `x_values` | `list[float] \| None` | `None` | X values (scatter charts) |

### `AxisOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `title` | `str` | `""` | Axis title |
| `title_overlay` | `bool` | `False` | Title overlaps axis |
| `visible` | `bool` | `True` | Show the axis |
| `position` | `str` | `""` | Axis position (`b`, `l`, `t`, `r`) |
| `major_tick_mark` | `str` | `""` | Major tick style (`out`, `in`, `cross`, `none`) |
| `minor_tick_mark` | `str` | `""` | Minor tick style |
| `tick_label_pos` | `str` | `""` | Tick label position (`nextTo`, `low`, `high`, `none`) |
| `number_format` | `str` | `"General"` | Number format code |
| `major_unit` | `float \| None` | `None` | Major gridline interval |
| `minor_unit` | `float \| None` | `None` | Minor gridline interval |
| `min` | `float \| None` | `None` | Axis minimum |
| `max` | `float \| None` | `None` | Axis maximum |
| `crosses_at` | `float \| None` | `None` | Where the cross axis intersects |
| `major_gridlines` | `bool` | `False` | Show major gridlines |
| `minor_gridlines` | `bool` | `False` | Show minor gridlines |

### `LegendOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `show` | `bool` | `True` | Show legend |
| `position` | `str` | `"r"` | Position (`r`, `l`, `t`, `b`, `tr`) |
| `overlay` | `bool` | `False` | Legend overlaps plot area |

### `DataLabelOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `show_legend_key` | `bool` | `False` | Show legend color key |
| `show_value` | `bool` | `False` | Show data values |
| `show_category_name` | `bool` | `False` | Show category labels |
| `show_series_name` | `bool` | `False` | Show series name |
| `show_percent` | `bool` | `False` | Show percentages (pie charts) |
| `position` | `str` | `""` | Label position (`bestFit`, `outEnd`, `ctr`, etc.) |
| `show_leader_lines` | `bool` | `False` | Show leader lines to labels |

### `ChartProperties`

| Field | Type | Default | Description |
|---|---|---|---|
| `style` | `int` | `2` | Built-in chart style index |
| `language` | `str` | `"en-US"` | Chart language |
| `display_blanks_as` | `str` | `"gap"` | How to display blank cells (`gap`, `zero`, `span`) |
| `plot_visible_only` | `bool` | `True` | Plot only visible cells |
| `rounded_corners` | `bool` | `False` | Rounded chart border corners |
| `date1904` | `bool` | `False` | Use 1904 date system |
| `show_data_labels_over_max` | `bool` | `False` | Show labels beyond axis max |

### `BarChartOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `direction` | `str` | `""` | `"col"` (column) or `"bar"` (horizontal) — auto-set from `ChartKind` |
| `grouping` | `str` | `"clustered"` | `"clustered"`, `"stacked"`, `"percentStacked"` |
| `gap_width` | `int` | `150` | Gap between bars (0–500) |
| `overlap` | `int` | `0` | Bar overlap (-100 to 100) |
| `vary_colors` | `bool` | `False` | Different color per data point |

### `ScatterChartOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `scatter_style` | `str` | `"marker"` | `"marker"`, `"lineMarker"`, `"line"`, `"smooth"`, `"smoothMarker"` |
| `vary_colors` | `bool` | `False` | Different color per data point |

---

## Dataclasses — Images

### `ImageOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `path` | `str` | *(required)* | Path to the image file |
| `width` | `int` | `0` | Width in pixels (0 = auto from file) |
| `height` | `int` | `0` | Height in pixels (0 = auto from file) |
| `alt_text` | `str` | `""` | Alt text for accessibility |
| `position` | `InsertPosition` | `END` | Where to insert |
| `caption` | `CaptionOptions \| None` | `None` | Optional image caption |

---

## Dataclasses — Headers, Footers & Page Numbers

### `HeaderFooterContent`

| Field | Type | Default | Description |
|---|---|---|---|
| `left_text` | `str` | `""` | Left-aligned text |
| `center_text` | `str` | `""` | Center-aligned text |
| `right_text` | `str` | `""` | Right-aligned text |

### `HeaderOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `type` | `HeaderType` | `DEFAULT` | Header type |
| `different_first` | `bool` | `False` | Enable different first-page header |
| `different_odd_even` | `bool` | `False` | Enable different odd/even headers |

### `FooterOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `type` | `FooterType` | `DEFAULT` | Footer type |
| `different_first` | `bool` | `False` | Enable different first-page footer |
| `different_odd_even` | `bool` | `False` | Enable different odd/even footers |

### `PageNumberOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `start` | `int` | `1` | Starting page number |
| `format` | `PageNumberFormat` | `DECIMAL` | Number format |

---

## Dataclasses — Page Layout & Breaks

### `PageLayoutOptions`

All dimensions are in twips (1 twip = 1/1440 inch).

| Field | Type | Description |
|---|---|---|
| `page_width` | `int` | Page width |
| `page_height` | `int` | Page height |
| `orientation` | `PageOrientation` | Portrait or landscape |
| `margin_top` | `int` | Top margin |
| `margin_right` | `int` | Right margin |
| `margin_bottom` | `int` | Bottom margin |
| `margin_left` | `int` | Left margin |
| `margin_header` | `int` | Header margin |
| `margin_footer` | `int` | Footer margin |
| `margin_gutter` | `int` | Gutter margin |

### `BreakOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `position` | `InsertPosition` | `END` | Where to insert the break |
| `anchor` | `str` | `""` | Anchor text for positional insertion |
| `section_type` | `SectionBreakType` | `NEXT_PAGE` | Section break type (section breaks only) |
| `page_layout` | `PageLayoutOptions \| None` | `None` | Page layout for the new section |

---

## Dataclasses — Watermarks

### `WatermarkOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `text` | `str` | `"DRAFT"` | Watermark text |
| `font_family` | `str` | `"Calibri"` | Font family |
| `color` | `str` | `"C0C0C0"` | Hex RGB color (without `#`) |
| `opacity` | `float` | `0.5` | Opacity (0.0–1.0) |
| `diagonal` | `bool` | `True` | Rotate text diagonally |

---

## Dataclasses — Hyperlinks & Bookmarks

### `HyperlinkOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `position` | `InsertPosition` | `END` | Where to insert |
| `anchor` | `str` | `""` | Anchor text for positional insertion |
| `tooltip` | `str` | `""` | Hover tooltip |
| `style` | `str` | `"Normal"` | Paragraph style |
| `color` | `str` | `"0563C1"` | Hex RGB link color |
| `underline` | `bool` | `True` | Underline the link text |
| `screen_tip` | `str` | `""` | Alternative to `tooltip` |

### `BookmarkOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `position` | `InsertPosition` | `END` | Where to insert |
| `anchor` | `str` | `""` | Anchor text for positional insertion |
| `style` | `str` | `"Normal"` | Paragraph style |
| `hidden` | `bool` | `True` | Whether bookmark is hidden |

---

## Dataclasses — Footnotes & Endnotes

### `FootnoteOptions`

| Field | Type | Description |
|---|---|---|
| `text` | `str` | Footnote text *(required)* |
| `anchor` | `str` | Text to attach the footnote to *(required)* |

### `EndnoteOptions`

| Field | Type | Description |
|---|---|---|
| `text` | `str` | Endnote text *(required)* |
| `anchor` | `str` | Text to attach the endnote to *(required)* |

---

## Dataclasses — Comments

### `CommentOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `text` | `str` | *(required)* | Comment text |
| `author` | `str` | `""` | Author name (defaults to `"Author"`) |
| `initials` | `str` | `""` | Author initials (defaults to first letter of author) |
| `anchor` | `str` | `""` | Text to attach the comment to |

### `Comment`

Returned by `get_comments()`.

| Field | Type | Description |
|---|---|---|
| `id` | `int` | Comment ID |
| `author` | `str` | Author name |
| `initials` | `str` | Author initials |
| `date` | `str` | ISO 8601 date string |
| `text` | `str` | Comment text |

---

## Dataclasses — Track Changes

### `TrackedInsertOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `text` | `str` | *(required)* | Text to insert |
| `author` | `str` | `""` | Author name |
| `date` | `datetime \| None` | `None` | Revision date (defaults to now) |
| `position` | `InsertPosition` | `END` | Where to insert |
| `anchor` | `str` | `""` | Anchor text for positional insertion |
| `style` | `str` | `"Normal"` | Paragraph style |
| `bold` | `bool` | `False` | Bold formatting |
| `italic` | `bool` | `False` | Italic formatting |
| `underline` | `bool` | `False` | Underline formatting |

### `TrackedDeleteOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `anchor` | `str` | *(required)* | Text to mark as deleted |
| `author` | `str` | `""` | Author name |
| `date` | `datetime \| None` | `None` | Revision date (defaults to now) |

---

## Dataclasses — Captions

### `CaptionOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `type` | `CaptionType` | *(required)* | `FIGURE` or `TABLE` |
| `position` | `CaptionPosition \| None` | `None` | Before or after the element |
| `description` | `str` | `""` | Caption text |
| `style` | `str` | `"Caption"` | Paragraph style |
| `auto_number` | `bool` | `True` | Auto-number captions |
| `alignment` | `CellAlignment` | `LEFT` | Caption alignment |
| `manual_number` | `int` | `0` | Manual caption number (when `auto_number` is `False`) |

---

## Dataclasses — Table of Contents

### `TOCOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `title` | `str` | `"Table of Contents"` | TOC heading text |
| `outline_levels` | `str` | `"1-3"` | Heading levels to include |
| `position` | `InsertPosition` | `BEGINNING` | Where to insert |
| `update_on_open` | `bool` | `True` | Prompt Word to update the TOC on open |

---

## Dataclasses — Document Properties

### `CoreProperties`

| Field | Type | Default | Description |
|---|---|---|---|
| `title` | `str` | `""` | Document title |
| `subject` | `str` | `""` | Subject |
| `creator` | `str` | `""` | Author |
| `keywords` | `str` | `""` | Keywords |
| `description` | `str` | `""` | Description/comments |
| `category` | `str` | `""` | Category |
| `content_status` | `str` | `""` | Status (e.g. "Draft", "Final") |
| `created` | `datetime \| None` | `None` | Creation date |
| `modified` | `datetime \| None` | `None` | Last modified date |
| `last_modified_by` | `str` | `""` | Last author |
| `revision` | `str` | `""` | Revision number |

### `AppProperties`

| Field | Type | Default | Description |
|---|---|---|---|
| `company` | `str` | `""` | Company name |
| `manager` | `str` | `""` | Manager |
| `application` | `str` | `""` | Application name |
| `app_version` | `str` | `""` | Application version |
| `template` | `str` | `""` | Template name |
| `hyperlink_base` | `str` | `""` | Base URL for hyperlinks |
| `total_time` | `int` | `0` | Total editing time (minutes) |
| `pages` | `int` | `0` | Page count |
| `words` | `int` | `0` | Word count |
| `characters` | `int` | `0` | Character count |
| `characters_with_spaces` | `int` | `0` | Character count with spaces |
| `lines` | `int` | `0` | Line count |
| `paragraphs` | `int` | `0` | Paragraph count |
| `doc_security` | `int` | `0` | Security level |

### `CustomProperty`

| Field | Type | Default | Description |
|---|---|---|---|
| `name` | `str` | *(required)* | Property name |
| `value` | `object` | *(required)* | Property value (`str`, `int`, `float`, `bool`, or `datetime`) |
| `type` | `str` | `""` | Value type (auto-inferred: `lpwstr`, `i4`, `r8`, `bool`, `date`) |

---

## Dataclasses — Find & Replace

### `FindOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `match_case` | `bool` | `False` | Case-sensitive search |
| `whole_word` | `bool` | `False` | Match whole words only |
| `use_regex` | `bool` | `False` | Treat pattern as regex |
| `max_results` | `int` | `0` | Maximum results (0 = unlimited) |
| `in_paragraphs` | `bool` | `True` | Search in paragraphs |
| `in_tables` | `bool` | `True` | Search in tables |
| `in_headers` | `bool` | `False` | Search in headers |
| `in_footers` | `bool` | `False` | Search in footers |

### `ReplaceOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `match_case` | `bool` | `False` | Case-sensitive matching |
| `whole_word` | `bool` | `False` | Match whole words only |
| `in_paragraphs` | `bool` | `True` | Replace in paragraphs |
| `in_tables` | `bool` | `True` | Replace in tables |
| `in_headers` | `bool` | `False` | Replace in headers |
| `in_footers` | `bool` | `False` | Replace in footers |
| `max_replacements` | `int` | `0` | Maximum replacements (0 = unlimited) |

### `DeleteOptions`

| Field | Type | Default | Description |
|---|---|---|---|
| `match_case` | `bool` | `False` | Case-sensitive matching |
| `whole_word` | `bool` | `False` | Match whole words only |

### `TextMatch`

Returned by `find_text()`.

| Field | Type | Description |
|---|---|---|
| `text` | `str` | The matched text |
| `paragraph` | `int` | Paragraph index (0-based) |
| `position` | `int` | Character offset in full text |
| `before` | `str` | Up to 50 characters before the match |
| `after` | `str` | Up to 50 characters after the match |

---

## Constants

All page size and margin constants are in **twips** (1 twip = 1/1440 inch).

### Page Sizes

| Constant | Value | Size |
|---|---|---|
| `PAGE_WIDTH_LETTER` | `12240` | 8.5 inches |
| `PAGE_HEIGHT_LETTER` | `15840` | 11 inches |
| `PAGE_WIDTH_LEGAL` | `12240` | 8.5 inches |
| `PAGE_HEIGHT_LEGAL` | `20160` | 14 inches |
| `PAGE_WIDTH_A4` | `11906` | 210 mm |
| `PAGE_HEIGHT_A4` | `16838` | 297 mm |
| `PAGE_WIDTH_A3` | `16838` | 297 mm |
| `PAGE_HEIGHT_A3` | `23811` | 420 mm |
| `PAGE_WIDTH_TABLOID` | `15840` | 11 inches |
| `PAGE_HEIGHT_TABLOID` | `24480` | 17 inches |

### Margins

| Constant | Value | Size |
|---|---|---|
| `MARGIN_DEFAULT` | `1440` | 1 inch |
| `MARGIN_NARROW` | `720` | 0.5 inch |
| `MARGIN_WIDE` | `2160` | 1.5 inches |
| `MARGIN_HEADER_FOOTER_DEFAULT` | `720` | 0.5 inch |
