"""Tests for shared.py pre-loop verification + retry inside the runner."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from src.incremental_runner import _ensure_shared_module
from src.shared_default import DEFAULT_SHARED_PY


def test_ensure_shared_module_no_style_writes_default(tmp_path: Path) -> None:
    g = Mock()
    g.extend_shared_module.return_value = DEFAULT_SHARED_PY
    target = tmp_path / "shared.py"
    log: list[str] = []
    ok = _ensure_shared_module(
        generator=g,
        style_content=None,
        shared_path=target,
        max_retries=3,
        error_log=log,
    )
    assert ok is True
    assert target.read_text(encoding="utf-8").startswith('"""Shared helpers')
    assert log == []


def test_ensure_shared_module_retries_on_syntax_error_then_succeeds(
    tmp_path: Path,
) -> None:
    g = Mock()
    g.extend_shared_module.return_value = "this is ::: not python"
    g.fix_shared_module.return_value = DEFAULT_SHARED_PY
    target = tmp_path / "shared.py"
    log: list[str] = []
    ok = _ensure_shared_module(
        generator=g,
        style_content="brand red",
        shared_path=target,
        max_retries=3,
        error_log=log,
    )
    assert ok is True
    assert g.fix_shared_module.call_count == 1
    assert any("verification attempt 1" in e for e in log)


def test_ensure_shared_module_fails_after_retries(tmp_path: Path) -> None:
    g = Mock()
    bad = "from pptx.dml.color import RGBColor\n# missing default symbols\n"
    g.extend_shared_module.return_value = bad
    g.fix_shared_module.return_value = bad
    target = tmp_path / "shared.py"
    log: list[str] = []
    ok = _ensure_shared_module(
        generator=g,
        style_content="brand X",
        shared_path=target,
        max_retries=2,
        error_log=log,
    )
    assert ok is False
    assert g.fix_shared_module.call_count == 1
    assert any("missing required default symbols" in e for e in log)


def test_ensure_shared_module_rejects_module_with_top_level_import_error(
    tmp_path: Path,
) -> None:
    g = Mock()
    bad = "import nonexistent_module_xyz\n"
    g.extend_shared_module.return_value = bad
    g.fix_shared_module.return_value = DEFAULT_SHARED_PY
    target = tmp_path / "shared.py"
    log: list[str] = []
    ok = _ensure_shared_module(
        generator=g,
        style_content="x",
        shared_path=target,
        max_retries=3,
        error_log=log,
    )
    assert ok is True
    assert g.fix_shared_module.call_count == 1
