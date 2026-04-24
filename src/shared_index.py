"""Build a compact symbol index from ``shared.py`` for incremental H2 prompts."""

from __future__ import annotations

import ast
from pathlib import Path


def build_shared_index(shared_py: Path) -> str:
    """
    Parse ``shared.py`` and emit module-level constants and function signatures.

    Private names (leading ``_``) are omitted. Used as the source of truth for
    ``shared`` attribute names in incremental H2 scripts.
    """
    text = shared_py.read_text(encoding="utf-8")
    tree = ast.parse(text, filename=str(shared_py))
    lines: list[str] = []

    for node in tree.body:
        if isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name) and not t.id.startswith("_"):
                    lines.append(ast.unparse(node))
                    break
        elif isinstance(node, ast.AnnAssign):
            if isinstance(node.target, ast.Name) and not node.target.id.startswith("_"):
                lines.append(ast.unparse(node))
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            if node.name.startswith("_"):
                continue
            args = ast.unparse(node.args)
            if not args.startswith("("):
                args = f"({args})"
            doc = ast.get_docstring(node)
            first = (doc or "").split("\n", 1)[0].strip() if doc else ""
            suffix = f'  # "{first}"' if first else ""
            lines.append(f"def {node.name}{args}:{suffix}")

    if not lines:
        return "(no public module-level symbols detected in shared.py)"
    return "\n".join(lines)
