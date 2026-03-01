from __future__ import annotations

import os
import zipfile
from pathlib import Path

from .errors import InvalidDocxError


def extract_zip(docx_path: str | os.PathLike[str], dest_dir: str | os.PathLike[str]) -> None:
    path = Path(docx_path)
    destination = Path(dest_dir).resolve()
    if not path.exists() or not path.is_file():
        raise InvalidDocxError(f"File not found: {path}")
    try:
        with zipfile.ZipFile(path) as zf:
            for member in zf.infolist():
                target = destination / member.filename
                if not _is_within_root(destination, target):
                    raise InvalidDocxError(f"Archive contains unsafe path: {member.filename}")
            zf.extractall(destination)
    except zipfile.BadZipFile as exc:
        raise InvalidDocxError(f"Invalid DOCX archive: {path}") from exc


def create_zip_from_dir(source_dir: str | os.PathLike[str], output_path: str | os.PathLike[str]) -> None:
    source = Path(source_dir)
    output = Path(output_path)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(source):
            root_path = Path(root)
            for name in files:
                file_path = root_path / name
                arcname = file_path.relative_to(source)
                zf.write(file_path, arcname.as_posix())


def _is_within_root(root: Path, target: Path) -> bool:
    target_path = target.resolve()
    return root == target_path or root in target_path.parents
