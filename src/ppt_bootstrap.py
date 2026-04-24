"""Deterministic empty PowerPoint creation for incremental generation."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation


def create_empty_ppt(path: Path) -> int:
    """
    Create a minimal .pptx at ``path`` and return baseline slide count.

    Uses python-pptx default template (typically one blank slide). Parent
    directories are created as needed.

    Args:
        path: Destination file path (``.pptx``).

    Returns:
        Number of slides in the saved file (baseline for +1 checks).
    """
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    # Some python-pptx builds ship a template with 0 slides; ensure a baseline slide.
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])
    prs.save(str(path))
    verify = Presentation(str(path))
    return len(verify.slides)
