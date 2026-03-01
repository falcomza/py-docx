from pydocx import BookmarkOptions, HyperlinkOptions, InsertPosition, new


def main() -> None:
    u = new("template.docx")
    try:
        u.add_heading(1, "Hyperlink Demo", InsertPosition.END)
        u.add_text("See the section below for details.", InsertPosition.END)

        u.create_bookmark_with_text(
            "details_section",
            "Details Section",
            BookmarkOptions(position=InsertPosition.END, style="Heading2"),
        )
        u.add_text("This is the target of the internal link.", InsertPosition.END)

        u.insert_internal_link(
            "Jump to Details Section",
            "details_section",
            HyperlinkOptions(position=InsertPosition.BEGINNING),
        )

        u.insert_hyperlink(
            "Open Python.org",
            "https://www.python.org",
            HyperlinkOptions(position=InsertPosition.END),
        )

        u.add_text("Key finding: 42% improvement in throughput.", InsertPosition.END)
        u.wrap_text_in_bookmark("key_finding", "Key finding")
        u.insert_internal_link(
            "See key finding",
            "key_finding",
            HyperlinkOptions(position=InsertPosition.END),
        )

        u.save("output_hyperlinks_bookmarks.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
