from pydocx import FindOptions, InsertPosition, ReplaceOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.add_text("The quick brown fox jumps over the lazy dog.", InsertPosition.END)
        u.add_text("The QUICK brown fox.", InsertPosition.END)

        matches = u.find_text("quick", FindOptions(match_case=False, whole_word=True))
        for m in matches:
            print(f"Found '{m.text}' in paragraph {m.paragraph} at {m.position}")

        replaced = u.replace_text("brown", "red", ReplaceOptions(match_case=False))
        print(f"Replaced {replaced} occurrence(s)")

        u.save("output_find_replace.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
