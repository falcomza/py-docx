import contextlib

from pydocx import CaptionListOptions, CaptionOptions, CaptionType, ImageOptions, InsertPosition, new_blank


def main() -> None:
    u = new_blank()
    try:
        u.insert_table_of_figures(
            CaptionListOptions(caption_label="Figure", title="Table of Figures", position=InsertPosition.BEGINNING)
        )
        u.add_text("Body text.", InsertPosition.END)
        with contextlib.suppress(FileNotFoundError):  # image asset optional for this demo
            u.insert_image(
                ImageOptions(
                    path="examples/assets/sample.png",
                    position=InsertPosition.END,
                    caption=CaptionOptions(type=CaptionType.FIGURE, description="A sample figure"),
                )
            )
        u.update_toc()
        u.save("tof_out.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
