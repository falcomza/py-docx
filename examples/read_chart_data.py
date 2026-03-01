from pydocx import ChartOptions, SeriesOptions, new


def main() -> None:
    u = new("template.docx")
    try:
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

        data = u.get_chart_data(1)
        print(data.chart_title)
        print(data.categories)
        for s in data.series:
            print(s.name, s.values, s.x_values)

        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
