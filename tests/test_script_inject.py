"""Tests for orchestrator script injection."""

from pathlib import Path

from src.script_inject import inject_slide_paths


def test_inject_slide_paths_preloads_shared(tmp_path: Path) -> None:
    shared = tmp_path / "shared.py"
    shared.write_text("X = 1\n", encoding="utf-8")
    target = tmp_path / "t.pptx"
    code = "assert shared.X == 1\n"
    out = inject_slide_paths(code, target, shared)
    assert "TARGET_PPTX" in out
    assert "SHARED_PY_PATH" in out
    assert "exec_module" in out
    assert "shared =" in out or "= _orch" in out


def test_inject_slide_paths_chosen_layout_index_constant(tmp_path: Path) -> None:
    shared = tmp_path / "shared.py"
    shared.write_text("X = 1\n", encoding="utf-8")
    target = tmp_path / "t.pptx"
    code = "pass\n"

    out = inject_slide_paths(code, target, shared, chosen_layout_index=3)
    assert "CHOSEN_LAYOUT_INDEX = 3" in out

    out2 = inject_slide_paths(code, target, shared)
    assert "CHOSEN_LAYOUT_INDEX" not in out2


def test_inject_slide_paths_resolves_paths(tmp_path: Path) -> None:
    target = tmp_path / "deck.pptx"
    shared = tmp_path / "shared.py"
    shared.write_text("# empty\n", encoding="utf-8")
    out = inject_slide_paths("pass\n", target, shared)
    assert str(target.resolve()) in out or target.name in out
