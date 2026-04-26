"""Prefix generated scripts with orchestrator-controlled paths (deterministic)."""

from __future__ import annotations

from pathlib import Path


def _path_literal(p: Path) -> str:
    return repr(str(p.resolve()))


def inject_slide_paths(
    code: str,
    target_pptx: Path,
    shared_py: Path,
    *,
    chosen_layout_index: int | None = None,
) -> str:
    """
    Inject TARGET_PPTX, SHARED_PY_PATH, optional CHOSEN_LAYOUT_INDEX, and
    preload ``shared`` via importlib.

    Generated slide code should use ``shared`` for helpers and
    ``Presentation(TARGET_PPTX)`` for the deck to modify. The same script is
    used for the preamble (first slide on a 0-slide deck) and for every H2
    slide; both append exactly one slide.
    """
    chosen = ""
    if chosen_layout_index is not None:
        chosen = f"CHOSEN_LAYOUT_INDEX = {int(chosen_layout_index)}\n"

    header = f"""# --- injected by orchestrator (do not remove) ---
from pathlib import Path
import importlib.util as _orch_importlib_util

TARGET_PPTX = Path({_path_literal(target_pptx)})
SHARED_PY_PATH = Path({_path_literal(shared_py)})
{chosen}_orch_spec = _orch_importlib_util.spec_from_file_location(
    "orch_shared", str(SHARED_PY_PATH)
)
shared = _orch_importlib_util.module_from_spec(_orch_spec)
assert _orch_spec.loader is not None
_orch_spec.loader.exec_module(shared)
# --- end injected ---

"""
    return header + code.lstrip("\n")


# Backwards-compatible alias
inject_h2_slide_paths = inject_slide_paths
