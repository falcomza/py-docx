from pydocx import new


def main() -> None:
    u = new("template.docx")
    try:
        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
