from pydocx import ChartData, SeriesData, new


def main() -> None:
    u = new("templates/docx_template.docx")
    try:
        u.update_chart(
            1,
            ChartData(
                categories=["Q1", "Q2", "Q3", "Q4"],
                series=[
                    SeriesData("Revenue", [100, 150, 120, 180]),
                    SeriesData("Costs", [80, 90, 85, 95]),
                ],
                chart_title="Updated Chart",
            ),
        )
        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
