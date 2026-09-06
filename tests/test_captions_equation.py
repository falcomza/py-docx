from __future__ import annotations

from pydocx.captions import generate_caption_xml
from pydocx.options import CaptionOptions, CaptionType


def test_equation_caption_emits_seq_equation() -> None:
    xml = generate_caption_xml(CaptionOptions(type=CaptionType.EQUATION, description="Euler"))
    assert "SEQ Equation" in xml
    assert ">Equation " in xml
    assert ">Euler<" in xml
