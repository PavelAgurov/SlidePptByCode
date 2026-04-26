"""Deterministic PowerPoint initialization for incremental generation."""

from __future__ import annotations

import shutil
from pathlib import Path

from pptx import Presentation, presentation

def create_empty_ppt(path: Path) -> int:
    """
    Create an empty .pptx at ``path`` and return slide count (always 0).

    The deck preserves python-pptx's default ``slide_layouts`` (9 built-in
    layouts: ``Title Slide``, ``Title and Content``, ...) so the runner can
    select one for each generated slide. No baseline slide is added; the
    incremental pipeline appends slides itself.
    """
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    # Drop any slides that ship with the default template (defensive: most
    # python-pptx builds already start with 0 slides).
    if len(prs.slides) > 0:
        _drop_all_slides(prs)
    prs.save(str(path))
    verify = Presentation(str(path))
    return len(verify.slides)


def copy_deck_template(template: Path, dest: Path) -> int:
    """
    Copy an existing ``.pptx`` to ``dest`` and return its slide count.

    Verifies the copy opens with python-pptx. Does not modify ``template``.
    """
    src = template.resolve()
    dst = dest.resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    verify = Presentation(str(dst))
    return len(verify.slides)


def _drop_all_slides(prs : presentation.Presentation) -> None:
    """Remove all slides from an open ``Presentation`` (in-memory only).

    Uses python-pptx package internals (``_sldIdLst`` / ``part.drop_rel``).
    Layouts and masters are preserved.
    """
    sldIdLst = prs.slides._sldIdLst # pylint: disable=W0212
    rels = prs.part.rels # pylint: disable=W0212
    for sldId in list(sldIdLst):
        rId = sldId.attrib[ # type: ignore
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        ]
        if rId in rels:
            prs.part.drop_rel(rId)
        sldIdLst.remove(sldId) # type: ignore


def strip_all_slides(pptx_path: Path) -> int:
    """
    Remove all slides from ``pptx_path`` in-place. Layouts/masters preserved.

    Returns:
        Slide count after stripping (always 0).
    """
    p = Path(pptx_path).resolve()
    prs = Presentation(str(p))
    _drop_all_slides(prs)
    prs.save(str(p))
    verify = Presentation(str(p))
    return len(verify.slides)


def init_deck(deck_path: Path, template_pptx: Path | None) -> int:
    """
    Initialize the working deck with **0 slides** and ready-to-use layouts.

    - If ``template_pptx`` is provided: copy it to ``deck_path`` and strip
      all slides; the user's masters/layouts are preserved.
    - Otherwise: write a fresh empty deck with python-pptx defaults.

    Returns:
        Slide count after initialization (always 0).
    """
    if template_pptx is not None:
        copy_deck_template(template_pptx, deck_path)
        return strip_all_slides(deck_path)
    return create_empty_ppt(deck_path)
