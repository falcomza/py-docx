from pydocx import TOCOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.insert_toc(TOCOptions(title="Table of Contents", outline_levels="1-3"))
        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
