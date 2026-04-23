"""Split task markdown into preamble (before first H2) and one chunk per H2 section."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal


@dataclass(frozen=True)
class TaskChunk:
    """One slice of the task file for chunked LLM generation."""

    kind: Literal["preamble", "h2"]
    """preamble = content before first ## ; h2 = one ## section (one slide)."""

    title: str
    """Display title: empty for preamble; H2 heading text without ## for h2."""

    body: str
    """Markdown source for this chunk (includes heading line for h2 chunks)."""

    index: int
    """0 = preamble; 1..N = H2 sections in document order."""


# H1 line: exactly one # (not ##); optional leading whitespace on the line
_H1_LINE = re.compile(r"^\s*#(?!#)\s+(.+)$", re.MULTILINE)
# H2 line: exactly two # at line start, not ###
_H2_LINE = re.compile(r"^(##)(?!#)\s+(.+)$", re.MULTILINE)


def extract_deck_title(task_content: str) -> str | None:
    """Return first H1 title text if present."""
    m = re.search(r"^\s*#\s+(.+)$", task_content, re.MULTILINE)
    if not m:
        return None
    return m.group(1).strip()


def split_into_chunks(task_content: str) -> list[TaskChunk]:
    """
    Split task markdown into preamble + one chunk per top-level H2 section.

    Preamble is everything before the first line matching ``## `` (not ``###``).
    Each H2 chunk runs from its ``## Title`` line through the line before the next H2.

    Returns:
        [preamble, h2_1, ..., h2_N]. If there is no H2 in the document, returns
        a single preamble chunk only. For CLI generation, call
        ``validate_chunked_task_markdown`` first so this case is rejected.
    """
    text = task_content.replace("\r\n", "\n")

    matches = list(_H2_LINE.finditer(text))
    if not matches:
        return [
            TaskChunk(
                kind="preamble",
                title="",
                body=text.strip(),
                index=0,
            )
        ]

    first_h2_start = matches[0].start()
    preamble_raw = text[:first_h2_start].strip()

    chunks: list[TaskChunk] = [
        TaskChunk(kind="preamble", title="", body=preamble_raw, index=0)
    ]

    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        block = text[start:end].strip()
        title = m.group(2).strip()
        chunks.append(TaskChunk(kind="h2", title=title, body=block, index=len(chunks)))

    return chunks


def has_h2_sections(chunks: list[TaskChunk]) -> bool:
    """True if markdown had at least one H2 section (chunked mode is meaningful)."""
    return any(c.kind == "h2" for c in chunks)


def validate_chunked_task_markdown(task_content: str) -> None:
    """
    Ensure the task file matches the required structure for chunked generation.

    Requires:
    - At least one H1 line (``# `` at line start, not ``##``).
    - At least one H2 section (``## ``).
    - The first H1 must appear before the first H2.

    Raises:
        ValueError: If the structure is invalid.
    """
    text = task_content.replace("\r\n", "\n")

    h1 = _H1_LINE.search(text)
    if not h1:
        raise ValueError(
            "Task markdown must contain an H1 heading (a single line starting with "
            "'# ' for the presentation title), e.g. '# My presentation'."
        )

    h2 = _H2_LINE.search(text)
    if not h2:
        raise ValueError(
            "Task markdown must contain at least one H2 heading (a line starting with "
            "'## ' for each content slide), e.g. '## Slide 1 — Introduction'."
        )

    if h1.start() >= h2.start():
        raise ValueError(
            "Task markdown must place the H1 presentation title before the first "
            "H2 slide heading (##). Move the '# ...' line above all '## ...' sections."
        )


def section_function_name(section_index: int) -> str:
    """1-based section index -> Python function name (add_section_001)."""
    if section_index < 1:
        raise ValueError("section_index must be >= 1")
    return f"add_section_{section_index:03d}"
