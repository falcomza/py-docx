from pydocx import DeleteOptions, InsertPosition, new


def main() -> None:
    u = new("template.docx")
    try:
        u.add_text("Delete me paragraph.", InsertPosition.END)
        u.add_text("Keep me paragraph.", InsertPosition.END)
        deleted = u.delete_paragraphs("Delete me", DeleteOptions(match_case=False, whole_word=False))
        print(f"Deleted {deleted} paragraph(s)")

        # Delete first table/image/chart if present in template
        # u.delete_table(1)
        # u.delete_image(1)
        # u.delete_chart(1)

        u.save("output_delete_ops.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
