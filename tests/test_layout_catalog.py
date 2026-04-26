from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from src.layout_catalog import (
    format_layout_card_compact,
    format_layout_card_full,
    format_layouts_catalog_compact,
    read_layouts,
)


def test_read_layouts_default_deck_has_title_placeholder(tmp_path: Path) -> None:
    p = tmp_path / "default.pptx"
    Presentation().save(str(p))

    layouts = read_layouts(p)
    assert len(layouts) >= 1

    # Default templates may differ; at least one layout should contain a
    # title-like placeholder.
    assert any(
        any(ph.type in ("TITLE", "CENTER_TITLE", "VERTICAL_TITLE") for ph in li.placeholders)
        for li in layouts
    )


def test_layout_formatters_shapes(tmp_path: Path) -> None:
    p = tmp_path / "default.pptx"
    Presentation().save(str(p))
    layouts = read_layouts(p)
    li = layouts[0]

    full = format_layout_card_full(li)
    assert "idx=" in full
    assert "type=" in full
    assert 'name="' in full
    assert 'prompt="' in full
    assert "bbox=" in full

    compact = format_layout_card_compact(li)
    assert "placeholders:" in compact
    assert "bbox=" not in compact
    assert "prompt=" not in compact
    assert "idx=" not in compact

    catalog = format_layouts_catalog_compact(layouts[:3])
    assert catalog.count("placeholders:") == min(3, len(layouts))

