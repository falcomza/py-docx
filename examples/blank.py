from pydocx import InsertPosition, new_blank


def main() -> None:
    u = new_blank()
    try:
        u.add_text("Hello from a blank document.", InsertPosition.END)
        u.save("output_blank.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
