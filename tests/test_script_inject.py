"""Tests for orchestrator script injection."""

from pathlib import Path

from src.script_inject import inject_h1_paths, inject_h2_slide_paths


def test_inject_h1_paths_contains_literals(tmp_path: Path) -> None:
    deck = tmp_path / "d.pptx"
    shared = tmp_path / "shared.py"
    code = "print(DECK_PPTX_PATH)\n"
    out = inject_h1_paths(code, deck, shared)
    assert "DECK_PPTX_PATH" in out
    assert "SHARED_PY_PATH" in out
    assert str(deck.resolve()) in out or deck.name in out


def test_inject_h2_preloads_shared(tmp_path: Path) -> None:
    shared = tmp_path / "shared.py"
    shared.write_text("X = 1\n", encoding="utf-8")
    target = tmp_path / "t.pptx"
    code = "assert shared.X == 1\n"
    out = inject_h2_slide_paths(code, target, shared)
    assert "TARGET_PPTX" in out
    assert "exec_module" in out
    assert "shared =" in out or "= _orch" in out


def test_inject_h2_chosen_layout_index_constant(tmp_path: Path) -> None:
    shared = tmp_path / "shared.py"
    shared.write_text("X = 1\n", encoding="utf-8")
    target = tmp_path / "t.pptx"
    code = "pass\n"

    out = inject_h2_slide_paths(code, target, shared, chosen_layout_index=3)
    assert "CHOSEN_LAYOUT_INDEX = 3" in out

    out2 = inject_h2_slide_paths(code, target, shared)
    assert "CHOSEN_LAYOUT_INDEX" not in out2
