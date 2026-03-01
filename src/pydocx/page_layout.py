from __future__ import annotations

from .options import (
    MARGIN_DEFAULT,
    MARGIN_HEADER_FOOTER_DEFAULT,
    PAGE_HEIGHT_A3,
    PAGE_HEIGHT_A4,
    PAGE_HEIGHT_LEGAL,
    PAGE_HEIGHT_LETTER,
    PAGE_WIDTH_A3,
    PAGE_WIDTH_A4,
    PAGE_WIDTH_LEGAL,
    PAGE_WIDTH_LETTER,
    PageLayoutOptions,
    PageOrientation,
)


def page_layout_letter_portrait() -> PageLayoutOptions:
    return PageLayoutOptions(
        page_width=PAGE_WIDTH_LETTER,
        page_height=PAGE_HEIGHT_LETTER,
        orientation=PageOrientation.PORTRAIT,
        margin_top=MARGIN_DEFAULT,
        margin_right=MARGIN_DEFAULT,
        margin_bottom=MARGIN_DEFAULT,
        margin_left=MARGIN_DEFAULT,
        margin_header=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_footer=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_gutter=0,
    )


def page_layout_letter_landscape() -> PageLayoutOptions:
    return PageLayoutOptions(
        page_width=PAGE_HEIGHT_LETTER,
        page_height=PAGE_WIDTH_LETTER,
        orientation=PageOrientation.LANDSCAPE,
        margin_top=MARGIN_DEFAULT,
        margin_right=MARGIN_DEFAULT,
        margin_bottom=MARGIN_DEFAULT,
        margin_left=MARGIN_DEFAULT,
        margin_header=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_footer=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_gutter=0,
    )


def page_layout_a4_portrait() -> PageLayoutOptions:
    return PageLayoutOptions(
        page_width=PAGE_WIDTH_A4,
        page_height=PAGE_HEIGHT_A4,
        orientation=PageOrientation.PORTRAIT,
        margin_top=MARGIN_DEFAULT,
        margin_right=MARGIN_DEFAULT,
        margin_bottom=MARGIN_DEFAULT,
        margin_left=MARGIN_DEFAULT,
        margin_header=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_footer=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_gutter=0,
    )


def page_layout_a4_landscape() -> PageLayoutOptions:
    return PageLayoutOptions(
        page_width=PAGE_HEIGHT_A4,
        page_height=PAGE_WIDTH_A4,
        orientation=PageOrientation.LANDSCAPE,
        margin_top=MARGIN_DEFAULT,
        margin_right=MARGIN_DEFAULT,
        margin_bottom=MARGIN_DEFAULT,
        margin_left=MARGIN_DEFAULT,
        margin_header=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_footer=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_gutter=0,
    )


def page_layout_a3_portrait() -> PageLayoutOptions:
    return PageLayoutOptions(
        page_width=PAGE_WIDTH_A3,
        page_height=PAGE_HEIGHT_A3,
        orientation=PageOrientation.PORTRAIT,
        margin_top=MARGIN_DEFAULT,
        margin_right=MARGIN_DEFAULT,
        margin_bottom=MARGIN_DEFAULT,
        margin_left=MARGIN_DEFAULT,
        margin_header=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_footer=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_gutter=0,
    )


def page_layout_a3_landscape() -> PageLayoutOptions:
    return PageLayoutOptions(
        page_width=PAGE_HEIGHT_A3,
        page_height=PAGE_WIDTH_A3,
        orientation=PageOrientation.LANDSCAPE,
        margin_top=MARGIN_DEFAULT,
        margin_right=MARGIN_DEFAULT,
        margin_bottom=MARGIN_DEFAULT,
        margin_left=MARGIN_DEFAULT,
        margin_header=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_footer=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_gutter=0,
    )


def page_layout_legal_portrait() -> PageLayoutOptions:
    return PageLayoutOptions(
        page_width=PAGE_WIDTH_LEGAL,
        page_height=PAGE_HEIGHT_LEGAL,
        orientation=PageOrientation.PORTRAIT,
        margin_top=MARGIN_DEFAULT,
        margin_right=MARGIN_DEFAULT,
        margin_bottom=MARGIN_DEFAULT,
        margin_left=MARGIN_DEFAULT,
        margin_header=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_footer=MARGIN_HEADER_FOOTER_DEFAULT,
        margin_gutter=0,
    )
