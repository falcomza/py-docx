from pydocx import CommentOptions, InsertPosition, new


def main() -> None:
    u = new("template.docx")
    try:
        u.add_text("This section needs review.", InsertPosition.END)
        u.insert_comment(
            CommentOptions(
                text="Please verify the numbers and sources.",
                author="Reviewer",
                anchor="needs review",
            )
        )
        comments = u.get_comments()
        for comment in comments:
            print(f"{comment.id}: {comment.author} -> {comment.text}")
        u.save("output_comments.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
