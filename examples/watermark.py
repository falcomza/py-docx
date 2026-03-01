from pydocx import WatermarkOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.set_text_watermark(
            WatermarkOptions(
                text="CONFIDENTIAL",
                color="C0C0C0",
                opacity=0.4,
                diagonal=True,
            )
        )
        u.save("output_watermark.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
