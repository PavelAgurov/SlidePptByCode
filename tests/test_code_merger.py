"""Tests for chunked script merge."""

import pytest

from src.code_merger import CHUNK_MERGE_MARKER, merge_chunked_modules, replace_merge_marker


def test_replace_merge_marker_inserts_calls_preserving_indent():
    preamble = f"""
from pptx import Presentation

def create_presentation():
    prs = Presentation()
    add_preamble_slides(prs)
        {CHUNK_MERGE_MARKER}
    prs.save("out.pptx")
"""
    out = replace_merge_marker(preamble, 2)
    assert CHUNK_MERGE_MARKER not in out
    assert "        add_section_001(prs)" in out
    assert "        add_section_002(prs)" in out


def test_replace_merge_marker_zero_sections():
    preamble = f"def x():\n    {CHUNK_MERGE_MARKER}\n"
    out = replace_merge_marker(preamble, 0)
    assert "pass  # no H2 sections" in out


def test_merge_chunked_modules_order():
    preamble = f"""
from pptx import Presentation

def add_preamble_slides(prs):
    pass

def create_presentation():
    prs = Presentation()
    add_preamble_slides(prs)
    {CHUNK_MERGE_MARKER}
    prs.save("x.pptx")
"""
    sec1 = """
def add_section_001(prs):
    pass
"""
    sec2 = """
def add_section_002(prs):
    pass
"""
    merged = merge_chunked_modules(preamble, [sec1, sec2])
    idx_section = merged.find("def add_section_001")
    idx_create = merged.find("def create_presentation")
    assert idx_section < idx_create
    assert "add_section_001(prs)" in merged
    assert "add_section_002(prs)" in merged


def test_replace_merge_marker_raises_without_marker():
    with pytest.raises(ValueError, match="Preamble code must contain"):
        _ = replace_merge_marker("def create_presentation():\n    pass\n", 1)


def test_merge_raises_without_create_presentation():
    preamble = f"""
def add_x():
    pass
    {CHUNK_MERGE_MARKER}
"""
    with pytest.raises(ValueError, match="def create_presentation"):
        _ = merge_chunked_modules(preamble, ["def add_section_001(prs):\n    pass"])
