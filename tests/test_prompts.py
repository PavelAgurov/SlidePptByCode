"""Prompt shape regression tests."""

from src.prompts import (
    SYSTEM_PROMPT,
    SYSTEM_PROMPT_INCREMENTAL_H1,
    SYSTEM_PROMPT_INCREMENTAL_H2,
    SYSTEM_PROMPT_LAYOUT_SELECTION,
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


def test_layout_selection_user_message_includes_allowed_indices() -> None:
    msg = format_layout_selection_user_message(
        slide_markdown="## Slide\n- A\n",
        deck_title="Deck",
        layouts_catalog_compact='- #0 "A" — placeholders: 2 (TITLE, BODY)',
        allowed_indices=[0, 1, 2],
        style_content=None,
        language=None,
    )
    assert "Allowed selected_layout_index values:" in msg
    assert "0, 1, 2" in msg
