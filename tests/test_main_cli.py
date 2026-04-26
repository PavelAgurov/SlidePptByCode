"""CLI argument parsing tests (forced-layout flags)."""

from __future__ import annotations

import sys

import pytest

from src.main import parse_args


def _set_argv(monkeypatch: pytest.MonkeyPatch, args: list[str]) -> None:
    monkeypatch.setattr(sys, "argv", ["src/main.py", "task.md", *args])


def test_template_layout_alias_maps_to_h2(monkeypatch: pytest.MonkeyPatch) -> None:
    _set_argv(monkeypatch, ["--template_layout", "Content"])
    ns = parse_args()
    assert ns.template_layout_h2 == "Content"
    assert ns.template_layout_h1 is None


def test_template_layout_combined_with_new_flags_errors(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _set_argv(
        monkeypatch,
        ["--template_layout", "Content", "--template_layout_h2", "Other"],
    )
    with pytest.raises(SystemExit):
        parse_args()


def test_template_layout_h1_only(monkeypatch: pytest.MonkeyPatch) -> None:
    _set_argv(monkeypatch, ["--template_layout_h1", "Title Slide"])
    ns = parse_args()
    assert ns.template_layout_h1 == "Title Slide"
    assert ns.template_layout_h2 is None
    assert ns.template_layout is None


def test_template_layout_h1_h2_independent(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _set_argv(
        monkeypatch,
        [
            "--template_layout_h1",
            "Title Slide",
            "--template_layout_h2",
            "Content",
        ],
    )
    ns = parse_args()
    assert ns.template_layout_h1 == "Title Slide"
    assert ns.template_layout_h2 == "Content"
