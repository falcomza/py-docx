from pydocx import FooterOptions, HeaderFooterContent, HeaderOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.set_header(
            HeaderFooterContent(left_text="Company Name", center_text="Confidential", right_text="Feb 2026"),
            HeaderOptions(),
        )
        u.set_footer(
            HeaderFooterContent(center_text="Page "),
            FooterOptions(),
        )
        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
