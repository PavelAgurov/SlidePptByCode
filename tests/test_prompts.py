"""Prompt shape regression tests."""

from src.layout_catalog import LayoutInfo
from src.models import LayoutDescription
from src.prompts import (
    SYSTEM_PROMPT,
    SYSTEM_PROMPT_INCREMENTAL_SLIDE,
    SYSTEM_PROMPT_LAYOUT_DESCRIBE,
    SYSTEM_PROMPT_LAYOUT_SELECTION,
    SYSTEM_PROMPT_SHARED_EXTEND,
    format_layout_describe_user_message,
    format_layout_selection_user_message,
    format_incremental_slide_user_message,
    format_shared_extend_user_message,
)
from src.shared_default import DEFAULT_SHARED_PY


def test_core_system_prompts_no_inline_python_fence() -> None:
    """Large inline ```python blocks were moved to get_code_snippet."""
    for label, body in (
        ("SYSTEM_PROMPT", SYSTEM_PROMPT),
        ("SYSTEM_PROMPT_INCREMENTAL_SLIDE", SYSTEM_PROMPT_INCREMENTAL_SLIDE),
    ):
        assert "```python" not in body, f"{label} must not embed ```python blocks"


def test_system_prompts_reference_snippet_tool() -> None:
    assert "get_code_snippet" in SYSTEM_PROMPT
    assert "get_code_snippet" in SYSTEM_PROMPT_INCREMENTAL_SLIDE


def test_shared_extend_user_message_includes_default_and_style() -> None:
    msg = format_shared_extend_user_message(
        default_shared_py=DEFAULT_SHARED_PY,
        style_content="Brand color #76CCBE (ELF_GREEN).",
    )
    assert "DEFAULT shared.py" in msg
    assert "STYLE GUIDELINES" in msg
    assert "ELF_GREEN" in msg
    assert "hex_to_rgb" in msg


def test_shared_extend_system_prompt_disallows_other_imports() -> None:
    assert "from pptx.dml.color import RGBColor" in SYSTEM_PROMPT_SHARED_EXTEND
    assert "Do not import anything else" in SYSTEM_PROMPT_SHARED_EXTEND


def test_format_incremental_slide_user_message_no_h1_h2_distinction() -> None:
    msg = format_incremental_slide_user_message(
        section_markdown="## Slide A\n- bullet\n",
        deck_title="Deck",
        style_content=None,
        language=None,
        shared_index="",
        chosen_layout=None,
    )
    assert "<slide_markdown>" in msg
    assert "</slide_markdown>" in msg
    assert "## Slide A" in msg
    assert "<deck_title>Deck</deck_title>" in msg
    assert "Content slide " not in msg


def test_format_incremental_slide_user_message_does_not_embed_raw_style() -> None:
    """Style content must NOT leak into per-slide user message — palette is in shared.*."""
    raw_style = "### White\n- **HEX:** #FFFFFF\n### Aqua Squeeze\n- **HEX:** #76CCBE\n"
    msg = format_incremental_slide_user_message(
        section_markdown="## Slide A\n- bullet\n",
        deck_title="Deck",
        style_content=raw_style,
        language="English",
        shared_index="WHITE\nAQUA_SQUEEZE",
        chosen_layout=None,
    )
    assert "#FFFFFF" not in msg
    assert "Aqua Squeeze" not in msg
    assert "## STYLE GUIDELINES" not in msg
    assert "<style_note>" in msg
    assert "<shared_symbols" in msg
    assert "<language_requirement>" in msg
    assert "English" in msg


def test_layout_selection_prompt_has_no_inline_python_fence() -> None:
    assert "```python" not in SYSTEM_PROMPT_LAYOUT_SELECTION
    assert "```python" not in SYSTEM_PROMPT_LAYOUT_DESCRIBE


def test_layout_selection_user_message_includes_allowed_indices() -> None:
    layouts = [
        LayoutInfo(index=0, name="A", placeholders=()),
        LayoutInfo(index=1, name="B", placeholders=()),
        LayoutInfo(index=2, name="C", placeholders=()),
    ]
    msg = format_layout_selection_user_message(
        slide_markdown="## Slide\n- A\n",
        deck_title="Deck",
        layouts=layouts,
        allowed_indices=[0, 1, 2],
        style_content=None,
        language=None,
    )
    assert "Allowed selected_layout_index values:" in msg
    assert "0, 1, 2" in msg


def test_format_layout_describe_user_message_layout_card_last() -> None:
    li = LayoutInfo(index=3, name="Agenda", placeholders=())
    msg = format_layout_describe_user_message(layout=li)
    assert "Describe ONE master slide" in msg
    assert "Return one sentence" in msg
    assert "### Layout #3" in msg
    assert "Agenda" in msg


def test_format_layout_describe_user_message_with_descriptions_in_selection() -> None:
    layouts = [LayoutInfo(index=0, name="X", placeholders=())]
    msg = format_layout_selection_user_message(
        slide_markdown="x",
        deck_title="D",
        layouts=layouts,
        allowed_indices=[0],
        style_content=None,
        language=None,
        descriptions={
            0: LayoutDescription(
                slide_type="content",
                slide_has_image_placeholder=False,
                content_zones_count=2,
                description="Test description for this layout.",
            ),
        },
    )
    assert "desc: Test description" in msg
    assert "zones=2" in msg
