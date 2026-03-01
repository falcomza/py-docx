from pydocx import ImageOptions, InsertPosition, new


def main() -> None:
    u = new("template.docx")
    try:
        u.insert_image(ImageOptions(path="images/logo.png", width=400, position=InsertPosition.END))
        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
