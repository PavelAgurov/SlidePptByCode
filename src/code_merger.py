"""Merge preamble script and per-section functions into one executable module."""

from __future__ import annotations

import re

from .task_chunker import section_function_name

# Preamble LLM must emit this line inside create_presentation before save logic.
CHUNK_MERGE_MARKER = "# <CHUNKED_SECTION_CALLS>"


def replace_merge_marker(preamble_code: str, num_sections: int) -> str:
    """
    Replace CHUNK_MERGE_MARKER line with add_section_001(prs) ... calls.

    Preserves indentation of the marker line for the injected calls.
    """
    pattern = re.compile(r"^(\s*)" + re.escape(CHUNK_MERGE_MARKER) + r"\s*$", re.MULTILINE)
    m = pattern.search(preamble_code)
    if not m:
        raise ValueError(
            f"Preamble code must contain a line with exactly {CHUNK_MERGE_MARKER!r} "
            "(inside create_presentation, after add_preamble_slides)."
        )
    indent = m.group(1)
    if num_sections == 0:
        replacement = f"{indent}pass  # no H2 sections"
    else:
        lines = [f"{indent}{section_function_name(i)}(prs)" for i in range(1, num_sections + 1)]
        replacement = "\n".join(lines)
    return pattern.sub(replacement, preamble_code, count=1)


def merge_chunked_modules(preamble_code: str, section_codes: list[str]) -> str:
    """
    Insert section function definitions before ``def create_presentation`` and
    expand the merge marker into section calls.

    Args:
        preamble_code: Part 0 LLM output (imports, helpers, add_preamble_slides, create_presentation).
        section_codes: Part 1..N outputs, each typically one ``def add_section_NNN``.

    Returns:
        Single combined Python source string.
    """
    merged = replace_merge_marker(preamble_code, len(section_codes))
    needle = "def create_presentation"
    idx = merged.find(needle)
    if idx == -1:
        raise ValueError("Preamble code must define def create_presentation()")

    block = "\n\n".join(s.strip() for s in section_codes if s.strip())
    if not block:
        return merged

    return merged[:idx] + block + "\n\n" + merged[idx:]
