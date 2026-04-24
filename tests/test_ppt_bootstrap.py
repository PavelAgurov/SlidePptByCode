"""Tests for ppt_bootstrap.create_empty_ppt."""

from pathlib import Path
from typing import Any, cast

from pptx import Presentation

from src.ppt_bootstrap import create_empty_ppt
from src.task_chunker import TaskChunk


def test_create_empty_ppt_baseline_and_append(tmp_path: Path) -> None:
    p = tmp_path / "empty.pptx"
    baseline = create_empty_ppt(p)
    assert p.exists()
    assert baseline >= 1

    prs = Presentation(str(p))
    n0 = len(prs.slides)
    s = prs.slides.add_slide(prs.slide_layouts[1])
    st = s.shapes.title
    assert st is not None
    st.text = "Scratch slide"
    if len(s.placeholders) > 1:
        cast(Any, s.placeholders[1]).text = "Body text"
    prs.save(str(p))

    from src.validator import validate_scratch_append

    chunk = TaskChunk(kind="h2", title="T", body="## T\nx", index=1)
    v = validate_scratch_append(p, n0, chunk)
    assert v.is_valid
    assert v.slide_count == n0 + 1
