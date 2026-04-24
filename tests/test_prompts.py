"""Prompt shape regression tests."""

from src.prompts import (
    SYSTEM_PROMPT,
    SYSTEM_PROMPT_INCREMENTAL_H1,
    SYSTEM_PROMPT_INCREMENTAL_H2,
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
