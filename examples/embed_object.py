import zipfile
from io import BytesIO

from pydocx import EmbeddedObjectOptions, InsertPosition, new_blank


def _sample_xlsx() -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
    return buf.getvalue()


def main() -> None:
    u = new_blank()
    try:
        u.insert_embedded_object(
            EmbeddedObjectOptions(
                file_bytes=_sample_xlsx(),
                file_name="budget.xlsx",
                position=InsertPosition.END,
            )
        )
        u.save("embed_out.docx")
    finally:
        u.cleanup()


if __name__ == "__main__":
    main()
