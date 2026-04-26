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

from pydantic import ValidationError

from src.code_generator import CodeGenerator
from src.layout_catalog import LayoutInfo
from src.models import LayoutDescription
from src.prompts import LAYOUT_DESC_CACHE_PROMPT_VERSION

logger = logging.getLogger(__name__)


_DEFAULT_TEMPLATE_KEY = "default-pptx"


def _default_template_sha() -> str:
    """Stable, version-aware sha for python-pptx's built-in default layouts."""
    try:
        import pptx  # noqa: WPS433

        version = getattr(pptx, "__version__", "unknown")
    except Exception:  # noqa: BLE001
        version = "unknown"
    h = hashlib.sha256()
    h.update(f"python-pptx-default-layouts:{version}".encode("utf-8"))
    return h.hexdigest()


def cache_path(template_pptx: Path | None, generated_dir: Path) -> Path:
    d = generated_dir / "layout"
    if template_pptx is None:
        return d / f"{_DEFAULT_TEMPLATE_KEY}.json"
    return d / f"{Path(template_pptx).stem}.json"


def template_sha256(template_pptx: Path | None) -> str:
    if template_pptx is None:
        return _default_template_sha()
    h = hashlib.sha256()
    h.update(Path(template_pptx).read_bytes())
    return h.hexdigest()


def load_cached(
    cache_file: Path, expected_sha: str, expected_prompt_version: int = LAYOUT_DESC_CACHE_PROMPT_VERSION
) -> dict[int, LayoutDescription]:
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
    out: dict[int, LayoutDescription] = {}
    for row in data.get("layouts") or []:
        if not isinstance(row, dict):
            continue
        try:
            idx = int(row["index"])
        except (TypeError, KeyError, ValueError):
            continue
        desc = str(row.get("description", "")).strip()
        if not desc:
            continue
        st = row.get("slide_type")
        if st is None:
            continue
        raw_img = row.get("slide_has_image_placeholder")
        if raw_img is None:
            continue
        if isinstance(raw_img, bool):
            has_img = raw_img
        elif isinstance(raw_img, int) and raw_img in (0, 1):
            has_img = bool(raw_img)
        else:
            continue
        try:
            out[idx] = LayoutDescription(
                slide_type=st,
                slide_has_image_placeholder=has_img,
                description=desc,
            )
        except ValidationError:
            continue
    return out


def save_cache(
    cache_file: Path,
    *,
    template_pptx: Path | None,
    sha: str,
    layouts: list[LayoutInfo],
    descriptions: dict[int, LayoutDescription],
) -> None:
    rows: list[dict[str, str | int]] = []
    for li in sorted(layouts, key=lambda x: x.index):
        ld = descriptions.get(li.index)
        if ld is None:
            continue
        d = (ld.description or "").strip()
        if not d:
            continue
        rows.append(
            {
                "index": li.index,
                "name": li.name,
                "slide_type": ld.slide_type,
                "slide_has_image_placeholder": ld.slide_has_image_placeholder,
                "description": d,
            }
        )
    payload = {
        "template": (
            str(Path(template_pptx).resolve())
            if template_pptx is not None
            else f"<{_DEFAULT_TEMPLATE_KEY}>"
        ),
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
    template_pptx: Path | None,
    layouts: list[LayoutInfo],
    generated_dir: Path,
) -> dict[int, LayoutDescription]:
    """
    Return index -> ``LayoutDescription`` (slide_type, slide_has_image_placeholder, description).

    Missing indices fall back to compact layout lines without ``desc:`` in
    selection prompts.

    When ``template_pptx`` is ``None``, descriptions are computed/cached for
    python-pptx's built-in default layouts under a synthetic key.
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
    merged: dict[int, LayoutDescription] = {
        i: ld
        for i, ld in cached.items()
        if i in valid and (ld.description or "").strip()
    }
    before_llm = set(merged.keys())
    missing = [li for li in layouts if li.index not in merged]

    for li in missing:
        try:
            r = generator.describe_layout(li)
            if (r.description or "").strip():
                merged[li.index] = r
        except Exception as e:  # noqa: BLE001 — optional step; any failure is fine
            logger.warning("layout description: LLM failed for layout index=%s: %s", li.index, e)

    if set(merged.keys()) != before_llm:
        try:
            save_cache(cpath, template_pptx=template_pptx, sha=sha, layouts=layouts, descriptions=merged)
        except OSError as e:
            logger.warning("layout description: cannot write cache %s: %s", cpath, e)
    return merged
