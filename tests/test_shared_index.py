"""Tests for AST-based shared.py symbol index."""

from pathlib import Path

from src.shared_index import build_shared_index


def test_build_shared_index_skips_private_and_lists_public(tmp_path: Path) -> None:
    p = tmp_path / "shared.py"
    p.write_text(
        """
from pptx.dml.color import RGBColor

ELF_GREEN = RGBColor(24, 139, 120)
_HIDDEN = 1

def set_font(para, font_name=None, font_size=None):
    \"\"\"First line of doc.\"\"\"
    pass
""",
        encoding="utf-8",
    )
    idx = build_shared_index(p)
    assert "ELF_GREEN" in idx
    assert "_HIDDEN" not in idx
    assert "def set_font" in idx
    assert "First line of doc" in idx
