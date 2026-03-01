from pydocx import PageNumberFormat, PageNumberOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.set_page_number(PageNumberOptions(start=1, format=PageNumberFormat.DECIMAL))
        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
