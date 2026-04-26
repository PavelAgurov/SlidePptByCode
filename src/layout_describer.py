"""LLM layout descriptions for template master layouts, with JSON on-disk cache."""

from __future__ import annotations

import hashlib
import json
import logging
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from src.code_generator import CodeGenerator
from src.layout_catalog import LayoutInfo
from src.prompts import LAYOUT_DESC_CACHE_PROMPT_VERSION

logger = logging.getLogger(__name__)


def cache_path(template_pptx: Path, generated_dir: Path) -> Path:
    d = generated_dir / "layout"
    return d / f"{Path(template_pptx).stem}.json"


def template_sha256(template_pptx: Path) -> str:
    h = hashlib.sha256()
    h.update(Path(template_pptx).read_bytes())
    return h.hexdigest()


def load_cached(
    cache_file: Path, expected_sha: str, expected_prompt_version: int = LAYOUT_DESC_CACHE_PROMPT_VERSION
) -> dict[int, str]:
    if not cache_file.is_file():
        return {}
    try:
        data: dict[str, Any] = json.loads(cache_file.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as e:
        logger.debug("layout description cache: read/parse failed (%s)", e)
        return {}
    if str(data.get("template_sha256", "")) != expected_sha:
        return {}
    if int(data.get("prompt_version", -1)) != expected_prompt_version:
        return {}
    out: dict[int, str] = {}
    for row in data.get("layouts") or []:
        if not isinstance(row, dict):
            continue
        try:
            idx = int(row["index"])
        except (TypeError, KeyError, ValueError):
            continue
        desc = str(row.get("description", "")).strip()
        if desc:
            out[idx] = desc
    return out


def save_cache(
    cache_file: Path,
    *,
    template_pptx: Path,
    sha: str,
    layouts: list[LayoutInfo],
    descriptions: dict[int, str],
) -> None:
    rows: list[dict[str, str | int]] = []
    for li in sorted(layouts, key=lambda x: x.index):
        d = (descriptions.get(li.index) or "").strip()
        if d:
            rows.append({"index": li.index, "name": li.name, "description": d})
    payload = {
        "template": str(Path(template_pptx).resolve()),
        "template_sha256": sha,
        "prompt_version": LAYOUT_DESC_CACHE_PROMPT_VERSION,
        "generated_at": datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z"),
        "layouts": rows,
    }
    cache_file.parent.mkdir(parents=True, exist_ok=True)
    raw = json.dumps(payload, ensure_ascii=True, indent=2)
    fd, tmp = tempfile.mkstemp(
        dir=str(cache_file.parent),
        prefix=cache_file.name + ".",
        suffix=".tmp",
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(raw)
        os.replace(tmp, cache_file)
    except OSError:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def ensure_layout_descriptions(
    *,
    generator: CodeGenerator,
    template_pptx: Path,
    layouts: list[LayoutInfo],
    generated_dir: Path,
) -> dict[int, str]:
    """
    Return index -> one-line English description. Missing indices fall back to
    compact layout lines without ``desc:`` in selection prompts.
    """
    if not layouts:
        return {}
    cpath = cache_path(template_pptx, generated_dir)
    try:
        sha = template_sha256(template_pptx)
    except OSError as e:
        logger.warning("layout description: cannot read template for sha (%s): %s", template_pptx, e)
        return {}
    cached = load_cached(cpath, sha)
    valid = {li.index for li in layouts}
    merged: dict[int, str] = {
        i: t for i, t in cached.items() if i in valid and t.strip()
    }
    before_llm = set(merged.keys())
    missing = [li for li in layouts if li.index not in merged]

    for li in missing:
        try:
            r = generator.describe_layout(li)
            text = (r.description or "").strip()
            if text:
                merged[li.index] = text
        except Exception as e:  # noqa: BLE001 — optional step; any failure is fine
            logger.warning("layout description: LLM failed for layout index=%s: %s", li.index, e)

    if set(merged.keys()) != before_llm:
        try:
            save_cache(cpath, template_pptx=template_pptx, sha=sha, layouts=layouts, descriptions=merged)
        except OSError as e:
            logger.warning("layout description: cannot write cache %s: %s", cpath, e)
    return merged
