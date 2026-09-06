from pydocx import HeaderFooterContent, HeaderOptions, new_blank


def main() -> None:
    u = new_blank()
    try:
        u.set_header(HeaderFooterContent(center_text="Page {PAGE}"), HeaderOptions())
        # Force Word/LibreOffice to recalculate PAGE, NUMPAGES, TOC, SEQ, ... on open.
        u.force_field_update_on_open()
        u.save("field_update_out.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
