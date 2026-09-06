from pydocx import InsertPosition, ParagraphAlignment, ParagraphOptions, StyleDefinition, new_blank


def main() -> None:
    u = new_blank()
    try:
        u.add_style(
            StyleDefinition(
                id="Callout",
                name="Callout",
                based_on="Normal",
                bold=True,
                font_size=28,
                color="1F6FEB",
                alignment=ParagraphAlignment.CENTER,
                space_before=120,
                space_after=120,
            )
        )
        u.insert_paragraph(ParagraphOptions(text="Big idea lives here", style="Callout", position=InsertPosition.END))
        u.save("styles_out.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
