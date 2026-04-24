"""Tests for ppt_bootstrap.create_empty_ppt."""

from pathlib import Path
from typing import Any, cast

from pptx import Presentation

from src.ppt_bootstrap import copy_deck_template, create_empty_ppt
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


def test_copy_deck_template_preserves_source(tmp_path: Path) -> None:
    src = tmp_path / "source.pptx"
    dst = tmp_path / "out" / "copy.pptx"
    n_src = create_empty_ppt(src)
    src_mtime = src.stat().st_mtime_ns
    src_bytes = src.read_bytes()

    n_dst = copy_deck_template(src, dst)
    assert dst.exists()
    assert n_dst == n_src
    assert len(Presentation(str(dst)).slides) == n_src

    assert src.read_bytes() == src_bytes
    assert src.stat().st_mtime_ns == src_mtime
