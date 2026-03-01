from pydocx import EndnoteOptions, FootnoteOptions, InsertPosition, new


def main() -> None:
    u = new("template.docx")
    try:
        u.add_text("The experiment showed significant results.", position=InsertPosition.END)
        u.insert_footnote(FootnoteOptions(text="Based on data collected in Q3 2026.", anchor="significant results"))
        u.insert_endnote(EndnoteOptions(text="See full methodology in Appendix A.", anchor="experiment"))
        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
