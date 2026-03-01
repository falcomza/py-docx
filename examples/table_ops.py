from pydocx import ColumnDefinition, TableOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.insert_table(
            TableOptions(
                columns=[ColumnDefinition("Col A"), ColumnDefinition("Col B"), ColumnDefinition("Col C")],
                rows=[["A1", "B1", "C1"], ["A2", "B2", "C2"], ["A3", "B3", "C3"]],
            )
        )

        # Update a cell (table 1, row 2, col 3)
        u.update_table_cell(1, 2, 3, "Updated")

        # Merge horizontally (row 1, cols 1-3)
        u.merge_table_cells_horizontal(1, 1, 1, 3)

        # Merge vertically (col 1, rows 2-3)
        u.merge_table_cells_vertical(1, 2, 3, 1)

        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
