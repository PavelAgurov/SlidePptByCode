"""Colored console formatter (no ANSI in file-style use)."""

import logging

from src.colored_logging import ColorFormatter, stderr_supports_color


def test_color_formatter_plain_when_disabled() -> None:
    fmt = ColorFormatter("%(levelname)s - %(message)s", use_color=False)
    r = logging.LogRecord(
        name="t",
        level=logging.INFO,
        pathname="",
        lineno=0,
        msg="hello",
        args=(),
        exc_info=None,
    )
    r.color_event = "tool_call"
    out = fmt.format(r)
    assert "\033" not in out
    assert "hello" in out


def test_color_formatter_gen_section_bold_underline_blue() -> None:
    fmt = ColorFormatter("%(message)s", use_color=True)
    r = logging.LogRecord(
        name="t",
        level=logging.INFO,
        pathname="",
        lineno=0,
        msg="=== H2 slide 1/5 ===",
        args=(),
        exc_info=None,
    )
    r.color_event = "gen_section"
    out = fmt.format(r)
    assert "\033[1m" in out and "\033[4m" in out and "\033[94m" in out
    assert out.endswith("\033[0m")


def test_color_formatter_wraps_tool_call_when_enabled() -> None:
    fmt = ColorFormatter("%(message)s", use_color=True)
    r = logging.LogRecord(
        name="t",
        level=logging.INFO,
        pathname="",
        lineno=0,
        msg="LLM tool_call: x",
        args=(),
        exc_info=None,
    )
    r.color_event = "tool_call"
    out = fmt.format(r)
    assert out.startswith("\033[36m")
    assert out.endswith("\033[0m")


def test_layout_select_failed_uses_red() -> None:
    fmt = ColorFormatter("%(message)s", use_color=True)
    r = logging.LogRecord(
        name="t",
        level=logging.INFO,
        pathname="",
        lineno=0,
        msg="Slide 1/2: layout selection failed; using default",
        args=(),
        exc_info=None,
    )
    r.color_event = "layout_select_failed"
    out = fmt.format(r)
    assert out.startswith("\033[31m")
    assert out.endswith("\033[0m")


def test_tool_result_uses_dim() -> None:
    fmt = ColorFormatter("%(message)s", use_color=True)
    r = logging.LogRecord(
        name="t",
        level=logging.INFO,
        pathname="",
        lineno=0,
        msg="LLM tool_result: ...",
        args=(),
        exc_info=None,
    )
    r.color_event = "tool_result"
    out = fmt.format(r)
    assert "\033[2m" in out
    assert out.endswith("\033[0m")


def test_stderr_supports_color_is_bool() -> None:
    assert isinstance(stderr_supports_color(), bool)
