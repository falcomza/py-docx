from pydocx import (
    BreakOptions,
    InsertPosition,
    SectionBreakType,
    new,
    page_layout_letter_landscape,
)


def main() -> None:
    u = new("template.docx")
    try:
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

        u.save("output_breaks.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
