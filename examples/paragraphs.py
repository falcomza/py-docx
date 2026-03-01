from pydocx import InsertPosition, ParagraphAlignment, ParagraphOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.add_heading(1, "Executive Summary", InsertPosition.END)
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
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
