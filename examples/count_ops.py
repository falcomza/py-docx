from pydocx import new


def main() -> None:
    u = new("template.docx")
    try:
        print("Paragraphs:", u.get_paragraph_count())
        print("Tables:", u.get_table_count())
        print("Images:", u.get_image_count())
        print("Charts:", u.get_chart_count())
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
