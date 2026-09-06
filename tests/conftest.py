from __future__ import annotations

import zipfile
from collections.abc import Iterator
from pathlib import Path

import pytest

from pydocx.updater import Updater, new_blank


@pytest.fixture
def blank() -> Iterator[Updater]:
    u = new_blank()
    yield u
    u.cleanup()


def save_and_unzip(updater: Updater, out: Path) -> dict[str, str]:
    """Save the updater and return {archive member name: text content}."""
    updater.save(out)
    with zipfile.ZipFile(out, "r") as zf:
        return {n: zf.read(n).decode("utf-8", "replace") for n in zf.namelist()}


def part(updater: Updater, rel_path: str) -> str:
    return (updater.workspace / rel_path).read_text(encoding="utf-8")
