from pydocx import ChartOptions, ColumnDefinition, SeriesOptions, TableOptions, new


def main() -> None:
    u = new("template.docx")
    try:
        u.insert_table(
            TableOptions(
                columns=[
                    ColumnDefinition("Product"),
                    ColumnDefinition("Sales"),
                    ColumnDefinition("Growth"),
                ],
                rows=[
                    ["Product A", "120", "15%"],
                    ["Product B", "98", "8%"],
                ],
            )
        )

        u.insert_chart(
            ChartOptions(
                title="Quarterly Revenue",
                categories=["Q1", "Q2", "Q3", "Q4"],
                series=[
                    SeriesOptions("2025", [100, 120, 110, 130]),
                    SeriesOptions("2026", [110, 130, 125, 145]),
                ],
            )
        )

        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
