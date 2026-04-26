from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from src.incremental_runner import run_incremental_pipeline
from src.models import LayoutDescription, LayoutSelection
from src.layout_catalog import LayoutInfo


@pytest.fixture
def mock_config(tmp_path: Path) -> SimpleNamespace:
    out_dir = tmp_path / ".output"
    gen_dir = tmp_path / ".generated"
    out_dir.mkdir()
    gen_dir.mkdir()
    return SimpleNamespace(
        output_dir=out_dir,
        generated_dir=gen_dir,
    )


class _Executor:
    def __init__(self) -> None:
        self._saved: list[str] = []

    def save_code(self, script: str, filename: str) -> Path:
        # Pretend saving to disk; return a dummy path.
        self._saved.append(filename)
        return Path(filename)

    def execute(self, _code_file: Path, *, expected_output_pptx: Path) -> SimpleNamespace:
        # Runner validates scratch/deck by opening PPTX; we bypass by monkeypatching validators.
        return SimpleNamespace(success=True, error_message=None, stdout="", stderr="", traceback="")


def test_layout_selection_called_for_h2_only_when_template(monkeypatch: pytest.MonkeyPatch, mock_config: SimpleNamespace, tmp_path: Path) -> None:
    # Minimal task with H1 + one H2.
    task = "# Title\n\n## Slide 1\n- A\n"
    tpl = tmp_path / "tpl.pptx"
    # Create a valid pptx template file.
    from pptx import Presentation

    Presentation().save(str(tpl))

    generator = Mock()
    # H1 and H2 generation returns IncrementalLlmScriptCode-like objects.
    generator.generate_incremental_h1.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.generate_incremental_slide.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.fix_code_after_error.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.fix_incremental_h2_validation.return_value = SimpleNamespace(code="from pptx import Presentation\n")

    # Layout selection returns a valid index.
    generator.select_layout.return_value = LayoutSelection(explanation="x", selected_layout_index=0)
    generator.describe_layout.return_value = LayoutDescription(description="test layout one-liner")

    executor = _Executor()

    # Avoid touching real validators and shared index building.
    monkeypatch.setattr("src.incremental_runner.validate_incremental_h1_deck", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner.build_shared_index", lambda *_a, **_k: "X=1")
    monkeypatch.setattr("src.incremental_runner.validate_scratch_append", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner.validate_deck_after_append", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner.validate_presentation", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner._sha256_file", lambda *_a, **_k: "0" * 64)

    # Provide layouts from template.
    monkeypatch.setattr("src.incremental_runner.read_layouts", lambda *_a, **_k: [LayoutInfo(index=0, name="L0", placeholders=())])

    ok, _deck, _errs = run_incremental_pipeline(
        task_content=task,
        style_content=None,
        language=None,
        output_filename=str(tmp_path / "out.pptx"),
        generator=generator,
        executor=executor,  # type: ignore[arg-type]
        config=mock_config,  # type: ignore[arg-type]
        max_retries=1,
        slide_max=None,
        template_pptx=tpl,
    )
    assert ok is True
    # Called once per H2; H1 should not call selection.
    assert generator.select_layout.call_count == 1
    assert generator.describe_layout.call_count == 1


def test_template_layout_forces_no_llm_selection(
    monkeypatch: pytest.MonkeyPatch, mock_config: SimpleNamespace, tmp_path: Path
) -> None:
    task = "# Title\n\n## Slide 1\n- A\n"
    tpl = tmp_path / "tpl.pptx"
    from pptx import Presentation

    prs = Presentation()
    prs.save(str(tpl))

    generator = Mock()
    generator.generate_incremental_h1.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.generate_incremental_slide.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.fix_code_after_error.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.fix_incremental_h2_validation.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.select_layout.return_value = LayoutSelection(explanation="x", selected_layout_index=0)
    generator.describe_layout.return_value = LayoutDescription(description="d")

    executor = _Executor()
    monkeypatch.setattr("src.incremental_runner.validate_incremental_h1_deck", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner.build_shared_index", lambda *_a, **_k: "X=1")
    monkeypatch.setattr("src.incremental_runner.validate_scratch_append", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner.validate_deck_after_append", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner.validate_presentation", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner._sha256_file", lambda *_a, **_k: "0" * 64)

    monkeypatch.setattr(
        "src.incremental_runner.read_layouts",
        lambda *_a, **_k: [LayoutInfo(index=0, name="L0", placeholders=())],
    )

    ok, _deck, _errs = run_incremental_pipeline(
        task_content=task,
        style_content=None,
        language=None,
        output_filename=str(tmp_path / "out.pptx"),
        generator=generator,
        executor=executor,  # type: ignore[arg-type]
        config=mock_config,  # type: ignore[arg-type]
        max_retries=1,
        slide_max=None,
        template_pptx=tpl,
        template_layout="L0",
    )
    assert ok is True
    assert generator.select_layout.call_count == 0
    assert generator.describe_layout.call_count == 0


def test_template_layout_unknown_name_errors(
    monkeypatch: pytest.MonkeyPatch, mock_config: SimpleNamespace, tmp_path: Path
) -> None:
    task = "# Title\n\n## Slide 1\n- A\n"
    tpl = tmp_path / "tpl.pptx"
    from pptx import Presentation

    Presentation().save(str(tpl))

    generator = Mock()
    generator.generate_incremental_h1.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.generate_incremental_slide.return_value = SimpleNamespace(code="from pptx import Presentation\n")

    executor = _Executor()
    monkeypatch.setattr("src.incremental_runner.read_layouts", lambda *_a, **_k: [LayoutInfo(index=0, name="L0", placeholders=())])

    ok, _deck, errs = run_incremental_pipeline(
        task_content=task,
        style_content=None,
        language=None,
        output_filename=str(tmp_path / "out.pptx"),
        generator=generator,
        executor=executor,  # type: ignore[arg-type]
        config=mock_config,  # type: ignore[arg-type]
        max_retries=1,
        slide_max=None,
        template_pptx=tpl,
        template_layout="NOPE",
    )
    assert ok is False
    assert any("Template layout name not found" in e for e in errs)


def test_layout_selection_not_called_without_template(monkeypatch: pytest.MonkeyPatch, mock_config: SimpleNamespace, tmp_path: Path) -> None:
    task = "# Title\n\n## Slide 1\n- A\n"
    generator = Mock()
    generator.generate_incremental_h1.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.generate_incremental_slide.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.fix_code_after_error.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.fix_incremental_h2_validation.return_value = SimpleNamespace(code="from pptx import Presentation\n")
    generator.select_layout.return_value = LayoutSelection(explanation="x", selected_layout_index=0)

    executor = _Executor()

    monkeypatch.setattr("src.incremental_runner.validate_incremental_h1_deck", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner.build_shared_index", lambda *_a, **_k: "X=1")
    monkeypatch.setattr("src.incremental_runner.validate_scratch_append", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner.validate_deck_after_append", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner.validate_presentation", lambda *a, **k: SimpleNamespace(is_valid=True, issues=[]))
    monkeypatch.setattr("src.incremental_runner._sha256_file", lambda *_a, **_k: "0" * 64)

    ok, _deck, _errs = run_incremental_pipeline(
        task_content=task,
        style_content=None,
        language=None,
        output_filename=str(tmp_path / "out.pptx"),
        generator=generator,
        executor=executor,  # type: ignore[arg-type]
        config=mock_config,  # type: ignore[arg-type]
        max_retries=1,
        slide_max=None,
        template_pptx=None,
    )
    assert ok is True
    assert generator.select_layout.call_count == 0
    assert generator.describe_layout.call_count == 0

