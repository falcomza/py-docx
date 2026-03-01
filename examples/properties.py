from datetime import UTC, datetime

from pydocx import AppProperties, CoreProperties, CustomProperty, new


def main() -> None:
    u = new("template.docx")
    try:
        u.set_core_properties(
            CoreProperties(
                title="Q4 Report",
                subject="Financial Summary",
                creator="Finance Team",
                keywords="q4,finance,report",
                description="Quarterly financial results.",
                category="Reports",
                content_status="Draft",
                created=datetime(2026, 2, 1, tzinfo=UTC),
                last_modified_by="Finance Team",
                revision="2",
            )
        )
        u.set_app_properties(
            AppProperties(
                company="Acme Corp",
                manager="J. Smith",
                application="Microsoft Word",
                app_version="16.0000",
            )
        )
        u.set_custom_properties(
            [
                CustomProperty(name="Department", value="Finance"),
                CustomProperty(name="Reviewed", value=True),
                CustomProperty(name="Budget", value=1250000),
            ]
        )
        core = u.get_core_properties()
        print(core.title, core.creator, core.revision)
        u.save("output_properties.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
