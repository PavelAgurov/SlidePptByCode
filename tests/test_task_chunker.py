"""Tests for markdown task chunking."""

import pytest

from src.task_chunker import (
    extract_deck_title,
    has_h2_sections,
    section_function_name,
    split_into_chunks,
    validate_chunked_task_markdown,
)


def test_extract_deck_title():
    md = "# My Deck\n\n## First\nBody"
    assert extract_deck_title(md) == "My Deck"
    assert extract_deck_title("no h1") is None


def test_section_function_name():
    assert section_function_name(1) == "add_section_001"
    assert section_function_name(12) == "add_section_012"
    with pytest.raises(ValueError):
        section_function_name(0)


def test_split_no_h2_returns_preamble_only():
    md = "# Only title\n\nSome intro.\n\n### Not h2\nMore"
    chunks = split_into_chunks(md)
    assert len(chunks) == 1
    assert chunks[0].kind == "preamble"
    assert not has_h2_sections(chunks)


def test_validate_rejects_no_h2():
    md = "# Only title\n\nSome intro.\n\n### Not h2\nMore"
    with pytest.raises(ValueError, match="H2"):
        validate_chunked_task_markdown(md)


def test_validate_rejects_no_h1():
    md = "## Slide only\n\nBody"
    with pytest.raises(ValueError, match="H1"):
        validate_chunked_task_markdown(md)


def test_validate_rejects_h2_before_h1():
    md = "## First\n\n# Deck\n\n## Second\n"
    with pytest.raises(ValueError, match="H1 presentation title before"):
        validate_chunked_task_markdown(md)


def test_validate_accepts_valid():
    md = "# Deck\n\nIntro\n\n## A\n\nx\n\n## B\n\ny\n"
    validate_chunked_task_markdown(md)


def test_split_preamble_and_h2():
    md = """# Deck

Intro line.

## Alpha

Content A

## Beta

Content B
"""
    chunks = split_into_chunks(md)
    assert has_h2_sections(chunks)
    assert len(chunks) == 3
    assert chunks[0].kind == "preamble"
    assert "Intro line" in chunks[0].body
    assert "## Alpha" not in chunks[0].body
    assert chunks[1].kind == "h2"
    assert chunks[1].title == "Alpha"
    assert "## Alpha" in chunks[1].body
    assert "Content A" in chunks[1].body
    assert "## Beta" not in chunks[1].body
    assert chunks[2].title == "Beta"
    assert "Content B" in chunks[2].body


def test_h3_not_section_boundary():
    md = "## Real h2\n\n### Sub\n\n## Next h2\n"
    chunks = split_into_chunks(md)
    assert len(chunks) == 3
    assert "### Sub" in chunks[1].body
    assert chunks[2].title == "Next h2"


def test_indexes_sequential():
    md = "## A\n\n## B\n"
    chunks = split_into_chunks(md)
    assert [c.index for c in chunks] == [0, 1, 2]
