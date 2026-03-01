from pydocx import InsertPosition, TrackedDeleteOptions, TrackedInsertOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.insert_tracked_text(
            TrackedInsertOptions(
                text="This paragraph was added with track changes.",
                author="Editor",
                position=InsertPosition.END,
                bold=True,
            )
        )
        u.add_text("This sentence will be marked as deleted.", InsertPosition.END)
        u.delete_tracked_text(TrackedDeleteOptions(anchor="marked as deleted", author="Reviewer"))
        u.save("output_tracked_changes.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
