"""Console log colors (TTY + NO_COLOR); file logs stay plain."""

from __future__ import annotations

import logging
import os
import sys


def stderr_supports_color() -> bool:
    if not sys.stderr.isatty():
        return False
    if os.environ.get("NO_COLOR", "").strip():
        return False
    return True


class ColorFormatter(logging.Formatter):
    """
    Wrap formatted lines in ANSI colors when ``use_color`` is True.

    Drive colors via ``logger.info(..., extra={"color_event": "..."})``:
    - ``tool_call`` — model invoked a tool (cyan)
    - ``tool_result`` — tool response payload size / snippet first fetch (dim gray)
    - ``snippet_reuse`` — snippet cache reuse hint (dim gray)
    - ``llm_structured_ok`` — final structured parse succeeded (green)
    - ``llm_round`` — debug: parse+tools round (dim)
    - ``tool_error`` — tool guard / unknown tool (red, often WARNING level)
    - ``gen_section`` — start of a new LLM block (H1 once, each H2 slide once; not retries)

    ``ERROR`` / ``WARNING`` without ``color_event`` still get red / yellow.
    """

    RESET = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"
    BRIGHT_BLUE = "\033[94m"
    CYAN = "\033[36m"
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    RED = "\033[31m"
    DIM = "\033[2m"

    def __init__(self, fmt: str, *, use_color: bool | None = None) -> None:
        super().__init__(fmt)
        if use_color is None:
            use_color = stderr_supports_color()
        self._use_color = use_color

    def format(self, record: logging.LogRecord) -> str:
        msg = super().format(record)
        if not self._use_color:
            return msg

        ev = getattr(record, "color_event", None)
        if ev == "tool_call":
            return f"{self.CYAN}{msg}{self.RESET}"
        if ev == "tool_result":
            return f"{self.DIM}{msg}{self.RESET}"
        if ev == "snippet_reuse":
            return f"{self.DIM}{msg}{self.RESET}"
        if ev == "llm_structured_ok":
            return f"{self.GREEN}{msg}{self.RESET}"
        if ev == "llm_round":
            return f"{self.DIM}{msg}{self.RESET}"
        if ev == "gen_section":
            return (
                f"{self.BOLD}{self.UNDERLINE}{self.BRIGHT_BLUE}{msg}"
                f"{self.RESET}"
            )
        if ev in ("tool_error", "snippet_error"):
            return f"{self.RED}{msg}{self.RESET}"

        if record.levelno >= logging.ERROR:
            return f"{self.RED}{msg}{self.RESET}"
        if record.levelno == logging.WARNING:
            return f"{self.YELLOW}{msg}{self.RESET}"
        return msg
