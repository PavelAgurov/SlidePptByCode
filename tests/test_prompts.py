"""Prompt shape regression tests."""

from src.layout_catalog import LayoutInfo
from src.prompts import (
    SYSTEM_PROMPT,
    SYSTEM_PROMPT_INCREMENTAL_H1,
    SYSTEM_PROMPT_INCREMENTAL_H2,
    SYSTEM_PROMPT_LAYOUT_DESCRIBE,
    SYSTEM_PROMPT_LAYOUT_SELECTION,
    format_layout_describe_user_message,
    format_layout_selection_user_message,
)


def test_core_system_prompts_no_inline_python_fence() -> None:
    """Large inline ```python blocks were moved to get_code_snippet."""
    for label, body in (
        ("SYSTEM_PROMPT", SYSTEM_PROMPT),
        ("SYSTEM_PROMPT_INCREMENTAL_H1", SYSTEM_PROMPT_INCREMENTAL_H1),
        ("SYSTEM_PROMPT_INCREMENTAL_H2", SYSTEM_PROMPT_INCREMENTAL_H2),
    ):
        assert "```python" not in body, f"{label} must not embed ```python blocks"


def test_system_prompts_reference_snippet_tool() -> None:
    assert "get_code_snippet" in SYSTEM_PROMPT
    assert "get_code_snippet" in SYSTEM_PROMPT_INCREMENTAL_H1
    assert "get_code_snippet" in SYSTEM_PROMPT_INCREMENTAL_H2


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
        descriptions={0: "Test description for this layout."},
    )
    assert "desc: Test description" in msg
