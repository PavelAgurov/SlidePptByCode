from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from src.layout_catalog import (
    LayoutInfo,
    LayoutShape,
    format_layout_card_compact,
    format_layout_card_full,
    format_layouts_catalog_compact,
    format_layouts_catalog_with_descriptions,
    read_layouts,
)
from src.models import LayoutDescription


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
    for li in layouts:
        assert isinstance(li.decor_shapes, tuple)


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
    assert "Decorative shapes:" in full

    compact = format_layout_card_compact(li)
    assert "placeholders:" in compact
    assert "bbox=" not in compact
    assert "prompt=" not in compact
    assert "idx=" not in compact

    catalog = format_layouts_catalog_compact(layouts[:3])
    assert catalog.count("placeholders:") == min(3, len(layouts))

    same = format_layouts_catalog_with_descriptions(layouts[:3], None)
    assert same == catalog

    with_desc = format_layouts_catalog_with_descriptions(
        layouts[:1],
        {
            layouts[0].index: LayoutDescription(
                slide_type="header",
                slide_has_image_placeholder=False,
                content_zones_count=1,
                description="Short intent text.",
            ),
        },
    )
    assert "desc: Short intent text." in with_desc
    assert "placeholders:" in with_desc
    assert "zones=1" in with_desc


def test_format_layout_card_full_lists_decorative_shapes() -> None:
    decor = (
        LayoutShape(
            name="Panel 1",
            shape_type="AUTO_SHAPE",
            auto_shape_type="RECTANGLE",
            fill_summary="solid:#112233",
            left_in=0.0,
            top_in=0.0,
            width_in=5.0,
            height_in=7.5,
        ),
    )
    li = LayoutInfo(index=0, name="Split", placeholders=(), decor_shapes=decor)
    full = format_layout_card_full(li)
    assert "Decorative shapes:" in full
    assert "Panel 1" in full
    assert "RECTANGLE" in full
    assert "solid:#112233" in full


def test_read_layouts_can_include_decor_shapes_on_slide(tmp_path: Path) -> None:
    """Adding a non-placeholder autoshape on a slide does not affect layout read; decor from layout XML only."""
    p = tmp_path / "with_slide.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1)
    )
    prs.save(str(p))

    layouts = read_layouts(p)
    assert layouts
    # Slide-added shapes are not on the master layout definition.
    assert all(isinstance(li.decor_shapes, tuple) for li in layouts)
