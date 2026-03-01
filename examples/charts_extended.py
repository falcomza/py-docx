from pydocx import (
    AxisOptions,
    ChartKind,
    ChartOptions,
    DataLabelOptions,
    LegendOptions,
    ScatterChartOptions,
    SeriesOptions,
    new,
)


def main() -> None:
    u = new("template.docx")
    try:
        categories = ["Q1", "Q2", "Q3", "Q4"]
        series = [
            SeriesOptions("2025", [100, 120, 110, 130], color="4472C4"),
            SeriesOptions("2026", [110, 130, 125, 145], color="ED7D31"),
        ]

        u.insert_chart(
            ChartOptions(
                title="Quarterly Revenue (Line)",
                chart_kind=ChartKind.LINE,
                categories=categories,
                series=series,
                category_axis_title="Quarter",
                value_axis_title="Revenue",
                data_labels=DataLabelOptions(show_value=True),
                legend=LegendOptions(show=True, position="r"),
            )
        )

        u.insert_chart(
            ChartOptions(
                title="Market Share (Pie)",
                chart_kind=ChartKind.PIE,
                categories=["Product A", "Product B", "Product C"],
                series=[SeriesOptions("Share", [45, 30, 25])],
                data_labels=DataLabelOptions(show_value=True, show_percent=True, show_leader_lines=True),
            )
        )

        u.insert_chart(
            ChartOptions(
                title="Growth Trend (Area)",
                chart_kind=ChartKind.AREA,
                categories=categories,
                series=series,
                value_axis=AxisOptions(number_format="0", major_gridlines=True),
            )
        )

        u.insert_chart(
            ChartOptions(
                title="Experiment Results (Scatter)",
                chart_kind=ChartKind.SCATTER,
                categories=["A", "B", "C", "D"],
                series=[
                    SeriesOptions("Series 1", [1.5, 2.2, 2.8, 3.5], x_values=[10, 20, 30, 40]),
                    SeriesOptions("Series 2", [1.2, 2.0, 3.0, 3.8], x_values=[12, 22, 32, 42]),
                ],
                scatter_chart_options=ScatterChartOptions(scatter_style="lineMarker"),
                category_axis_title="X",
                value_axis_title="Y",
            )
        )

        u.save("output.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
