from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum


class InsertPosition(StrEnum):
    BEGINNING = "beginning"
    END = "end"
    AFTER_TEXT = "after_text"
    BEFORE_TEXT = "before_text"


class TableAlignment(StrEnum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


class CellAlignment(StrEnum):
    LEFT = "start"
    CENTER = "center"
    RIGHT = "end"


class VerticalAlignment(StrEnum):
    TOP = "top"
    CENTER = "center"
    BOTTOM = "bottom"


class TableWidthType(StrEnum):
    AUTO = "auto"
    PERCENTAGE = "pct"
    FIXED = "dxa"


class BorderStyle(StrEnum):
    NONE = "none"
    SINGLE = "single"
    DOUBLE = "double"
    DOTTED = "dotted"
    DASHED = "dashed"


class RowHeightRule(StrEnum):
    AUTO = "auto"
    AT_LEAST = "atLeast"
    EXACT = "exact"


class ListType(StrEnum):
    BULLET = "bullet"
    NUMBERED = "numbered"


class ParagraphAlignment(StrEnum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "both"


@dataclass(slots=True)
class ParagraphOptions:
    text: str
    style: str = "Normal"
    alignment: ParagraphAlignment | None = None
    position: InsertPosition = InsertPosition.END
    anchor: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    list_type: ListType | None = None
    list_level: int = 0
    restart: bool = False


@dataclass(slots=True)
class CommentOptions:
    text: str
    author: str = ""
    initials: str = ""
    anchor: str = ""


@dataclass(slots=True)
class Comment:
    id: int
    author: str
    initials: str
    date: str
    text: str


@dataclass(slots=True)
class TrackedInsertOptions:
    text: str
    author: str = ""
    date: datetime | None = None
    position: InsertPosition = InsertPosition.END
    anchor: str = ""
    style: str = "Normal"
    bold: bool = False
    italic: bool = False
    underline: bool = False


@dataclass(slots=True)
class TrackedDeleteOptions:
    anchor: str
    author: str = ""
    date: datetime | None = None


@dataclass(slots=True)
class TextMatch:
    text: str
    paragraph: int
    position: int
    before: str
    after: str


@dataclass(slots=True)
class FindOptions:
    match_case: bool = False
    whole_word: bool = False
    use_regex: bool = False
    max_results: int = 0
    in_paragraphs: bool = True
    in_tables: bool = True
    in_headers: bool = False
    in_footers: bool = False


@dataclass(slots=True)
class ReplaceOptions:
    match_case: bool = False
    whole_word: bool = False
    in_paragraphs: bool = True
    in_tables: bool = True
    in_headers: bool = False
    in_footers: bool = False
    max_replacements: int = 0


@dataclass(slots=True)
class DeleteOptions:
    match_case: bool = False
    whole_word: bool = False


@dataclass(slots=True)
class CoreProperties:
    title: str = ""
    subject: str = ""
    creator: str = ""
    keywords: str = ""
    description: str = ""
    category: str = ""
    content_status: str = ""
    created: datetime | None = None
    modified: datetime | None = None
    last_modified_by: str = ""
    revision: str = ""


@dataclass(slots=True)
class AppProperties:
    company: str = ""
    manager: str = ""
    application: str = ""
    app_version: str = ""
    template: str = ""
    hyperlink_base: str = ""
    total_time: int = 0
    pages: int = 0
    words: int = 0
    characters: int = 0
    characters_with_spaces: int = 0
    lines: int = 0
    paragraphs: int = 0
    doc_security: int = 0


@dataclass(slots=True)
class CustomProperty:
    name: str
    value: object
    type: str = ""


class SectionBreakType(StrEnum):
    NEXT_PAGE = "nextPage"
    CONTINUOUS = "continuous"
    EVEN_PAGE = "evenPage"
    ODD_PAGE = "oddPage"


class PageOrientation(StrEnum):
    PORTRAIT = "portrait"
    LANDSCAPE = "landscape"


@dataclass(slots=True)
class PageLayoutOptions:
    page_width: int
    page_height: int
    orientation: PageOrientation
    margin_top: int
    margin_right: int
    margin_bottom: int
    margin_left: int
    margin_header: int
    margin_footer: int
    margin_gutter: int


@dataclass(slots=True)
class BreakOptions:
    position: InsertPosition = InsertPosition.END
    anchor: str = ""
    section_type: SectionBreakType = SectionBreakType.NEXT_PAGE
    page_layout: PageLayoutOptions | None = None


# Page size constants in twips (1/1440 inch)
PAGE_WIDTH_LETTER = 12240
PAGE_HEIGHT_LETTER = 15840
PAGE_WIDTH_LEGAL = 12240
PAGE_HEIGHT_LEGAL = 20160
PAGE_WIDTH_A4 = 11906
PAGE_HEIGHT_A4 = 16838
PAGE_WIDTH_A3 = 16838
PAGE_HEIGHT_A3 = 23811
PAGE_WIDTH_TABLOID = 15840
PAGE_HEIGHT_TABLOID = 24480

# Margin constants in twips
MARGIN_DEFAULT = 1440
MARGIN_NARROW = 720
MARGIN_WIDE = 2160
MARGIN_HEADER_FOOTER_DEFAULT = 720


@dataclass(slots=True)
class WatermarkOptions:
    text: str = "DRAFT"
    font_family: str = "Calibri"
    color: str = "C0C0C0"
    opacity: float = 0.5
    diagonal: bool = True


@dataclass(slots=True)
class HyperlinkOptions:
    position: InsertPosition = InsertPosition.END
    anchor: str = ""
    tooltip: str = ""
    style: str = "Normal"
    color: str = "0563C1"
    underline: bool = True
    screen_tip: str = ""


@dataclass(slots=True)
class BookmarkOptions:
    position: InsertPosition = InsertPosition.END
    anchor: str = ""
    style: str = "Normal"
    hidden: bool = True


@dataclass(slots=True)
class ImageOptions:
    path: str
    width: int = 0  # pixels
    height: int = 0  # pixels
    alt_text: str = ""
    position: InsertPosition = InsertPosition.END
    caption: CaptionOptions | None = None


class HeaderType(StrEnum):
    FIRST = "first"
    EVEN = "even"
    DEFAULT = "default"


class FooterType(StrEnum):
    FIRST = "first"
    EVEN = "even"
    DEFAULT = "default"


@dataclass(slots=True)
class HeaderFooterContent:
    left_text: str = ""
    center_text: str = ""
    right_text: str = ""


@dataclass(slots=True)
class HeaderOptions:
    type: HeaderType = HeaderType.DEFAULT
    different_first: bool = False
    different_odd_even: bool = False


@dataclass(slots=True)
class FooterOptions:
    type: FooterType = FooterType.DEFAULT
    different_first: bool = False
    different_odd_even: bool = False


class PageNumberFormat(StrEnum):
    DECIMAL = "decimal"
    UPPER_ROMAN = "upperRoman"
    LOWER_ROMAN = "lowerRoman"
    UPPER_LETTER = "upperLetter"
    LOWER_LETTER = "lowerLetter"


@dataclass(slots=True)
class PageNumberOptions:
    start: int = 1
    format: PageNumberFormat = PageNumberFormat.DECIMAL


@dataclass(slots=True)
class TOCOptions:
    title: str = "Table of Contents"
    outline_levels: str = "1-3"
    position: InsertPosition = InsertPosition.BEGINNING
    update_on_open: bool = True


class CaptionType(StrEnum):
    FIGURE = "Figure"
    TABLE = "Table"


class CaptionPosition(StrEnum):
    BEFORE = "before"
    AFTER = "after"


@dataclass(slots=True)
class CaptionOptions:
    type: CaptionType
    position: CaptionPosition | None = None
    description: str = ""
    style: str = "Caption"
    auto_number: bool = True
    alignment: CellAlignment = CellAlignment.LEFT
    manual_number: int = 0


@dataclass(slots=True)
class FootnoteOptions:
    text: str
    anchor: str


@dataclass(slots=True)
class EndnoteOptions:
    text: str
    anchor: str


@dataclass(slots=True)
class CellTextStyle:
    bold: bool = False
    italic: bool = False
    font_size: int | None = None  # half-points (e.g., 20 = 10pt)
    font_color: str | None = None  # hex RGB


@dataclass(slots=True)
class ConditionalStyle:
    text_style: CellTextStyle = field(default_factory=CellTextStyle)
    background: str | None = None


class ChartKind(StrEnum):
    COLUMN = "column"
    BAR = "bar"
    LINE = "line"
    PIE = "pie"
    AREA = "area"
    SCATTER = "scatter"


@dataclass(slots=True)
class AxisOptions:
    title: str = ""
    title_overlay: bool = False
    visible: bool = True
    position: str = ""
    major_tick_mark: str = ""
    minor_tick_mark: str = ""
    tick_label_pos: str = ""
    number_format: str = "General"
    major_unit: float | None = None
    minor_unit: float | None = None
    min: float | None = None
    max: float | None = None
    crosses_at: float | None = None
    major_gridlines: bool = False
    minor_gridlines: bool = False


@dataclass(slots=True)
class LegendOptions:
    show: bool = True
    position: str = "r"
    overlay: bool = False


@dataclass(slots=True)
class DataLabelOptions:
    show_legend_key: bool = False
    show_value: bool = False
    show_category_name: bool = False
    show_series_name: bool = False
    show_percent: bool = False
    position: str = ""
    show_leader_lines: bool = False


@dataclass(slots=True)
class ChartProperties:
    style: int = 2
    language: str = "en-US"
    display_blanks_as: str = "gap"
    plot_visible_only: bool = True
    rounded_corners: bool = False
    date1904: bool = False
    show_data_labels_over_max: bool = False


@dataclass(slots=True)
class BarChartOptions:
    direction: str = ""
    grouping: str = "clustered"
    gap_width: int = 150
    overlap: int = 0
    vary_colors: bool = False


@dataclass(slots=True)
class ScatterChartOptions:
    scatter_style: str = "marker"
    vary_colors: bool = False


@dataclass(slots=True)
class SeriesOptions:
    name: str
    values: list[float]
    x_values: list[float] | None = None
    color: str = ""
    invert_if_negative: bool = False
    smooth: bool = False
    show_markers: bool = True
    data_labels: DataLabelOptions | None = None


@dataclass(slots=True)
class SeriesData:
    name: str
    values: list[float]
    x_values: list[float] | None = None


@dataclass(slots=True)
class ChartData:
    categories: list[str]
    series: list[SeriesData]
    chart_title: str = ""
    category_axis_title: str = ""
    value_axis_title: str = ""


@dataclass(slots=True)
class ChartOptions:
    categories: list[str]
    series: list[SeriesOptions]
    title: str = ""
    chart_kind: ChartKind | str = ChartKind.COLUMN
    title_overlay: bool = False
    category_axis_title: str = ""
    value_axis_title: str = ""
    show_legend: bool = True
    legend_position: str = "r"
    width_emu: int = 6099523
    height_emu: int = 3340467
    position: InsertPosition = InsertPosition.END
    caption: CaptionOptions | None = None
    category_axis: AxisOptions | None = None
    value_axis: AxisOptions | None = None
    legend: LegendOptions | None = None
    data_labels: DataLabelOptions | None = None
    properties: ChartProperties | None = None
    bar_chart_options: BarChartOptions | None = None
    scatter_chart_options: ScatterChartOptions | None = None


@dataclass(slots=True)
class ColumnDefinition:
    title: str
    width: int | None = None
    bold: bool = False
    alignment: CellAlignment | None = None


@dataclass(slots=True)
class TableOptions:
    columns: list[ColumnDefinition]
    rows: list[list[str]]
    position: InsertPosition = InsertPosition.END
    table_style: str = "TableGrid"
    header_bold: bool = True
    header_repeat: bool = False
    header_text_style: CellTextStyle = field(default_factory=CellTextStyle)
    row_text_style: CellTextStyle = field(default_factory=CellTextStyle)
    cell_text_styles: dict[tuple[int, int], CellTextStyle] | None = None
    conditional_styles: dict[str, ConditionalStyle] | None = None
    paragraph_style: str = ""
    header_background: str = ""
    row_background: str = ""
    alternate_row_background: str = ""
    table_alignment: TableAlignment = TableAlignment.LEFT
    header_alignment: CellAlignment = CellAlignment.LEFT
    row_alignment: CellAlignment = CellAlignment.LEFT
    vertical_alignment: VerticalAlignment = VerticalAlignment.CENTER
    table_width_type: TableWidthType = TableWidthType.PERCENTAGE
    table_width: int = 5000  # 5000 = 100% when pct
    column_widths: list[int] | None = None
    proportional_column_widths: bool = False
    available_width: int = 0
    border_style: BorderStyle = BorderStyle.SINGLE
    border_size: int = 4
    border_color: str = "000000"
    cell_padding: int = 108
    header_row_height: int = 0
    header_height_rule: RowHeightRule = RowHeightRule.AUTO
    row_height: int = 0
    row_height_rule: RowHeightRule = RowHeightRule.AUTO
    caption: CaptionOptions | None = None
