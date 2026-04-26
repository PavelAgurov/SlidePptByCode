"""Default ``shared.py`` source used by the incremental pipeline.

This module is the deterministic fallback / starting point for the
``shared.py`` helper module that slide scripts import. When the user
provides ``--style``, the LLM extends this default; otherwise the runner
writes ``DEFAULT_SHARED_PY`` verbatim to disk.

All names defined at the top level of ``DEFAULT_SHARED_PY`` are treated as
required: the verification step ensures they are still present after any
LLM-driven extension.
"""

from __future__ import annotations

DEFAULT_SHARED_PY: str = '''"""Shared helpers and palette for incremental slide scripts."""
from pptx.dml.color import RGBColor


def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """Convert "#RRGGBB" (or "RRGGBB") to an (r, g, b) tuple of ints."""
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Expected 6 hex digits, got: {hex_color!r}")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def rgb(hex_color: str) -> RGBColor:
    """Build an ``RGBColor`` from a hex string like "#76CCBE"."""
    return RGBColor(*hex_to_rgb(hex_color))


# Neutral default palette. When --style is provided, the LLM extends this
# module with brand-specific colors; otherwise these neutrals are used.
WHITE = rgb("#FFFFFF")
BLACK = rgb("#000000")
GRAY_LIGHT = rgb("#F2F2F2")
GRAY = rgb("#888888")
GRAY_DARK = rgb("#333333")
PRIMARY = rgb("#1F4E79")
ACCENT = rgb("#2E75B6")
'''


def default_shared_symbol_names() -> list[str]:
    """Names that must remain present in ``shared.py`` after any LLM extension."""
    return [
        "hex_to_rgb",
        "rgb",
        "WHITE",
        "BLACK",
        "GRAY_LIGHT",
        "GRAY",
        "GRAY_DARK",
        "PRIMARY",
        "ACCENT",
    ]
