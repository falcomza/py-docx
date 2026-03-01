from pydocx import ColumnDefinition, TableOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.insert_table(
            TableOptions(
                columns=[ColumnDefinition("Product"), ColumnDefinition("Sales")],
                rows=[["Product A", "120"], ["Product B", "98"]],
            )
        )

        tables = u.get_table_text()
        for t_idx, table in enumerate(tables, start=1):
            print(f"Table {t_idx}:")
            for row in table:
                print(row)

        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
