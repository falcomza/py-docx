from pydocx import CaptionOptions, CaptionType, ColumnDefinition, ImageOptions, TableOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.insert_table(
            TableOptions(
                columns=[ColumnDefinition("A"), ColumnDefinition("B")],
                rows=[["1", "2"]],
                caption=CaptionOptions(type=CaptionType.TABLE, description="Sample table"),
            )
        )
        u.insert_image(
            ImageOptions(
                path="examples/assets/tiny.png",
                width=100,
                height=100,
                caption=CaptionOptions(type=CaptionType.FIGURE, description="Sample image"),
            )
        )
        u.update_caption(CaptionType.TABLE, 1, "Updated table caption")
        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
