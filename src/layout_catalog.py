"""Read slide layouts (masters) and placeholders from a .pptx template.

This module is used for selecting the best layout for a slide when a user
provides a template deck.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation


@dataclass(frozen=True, slots=True)
class LayoutPlaceholder:
    idx: int
    type: str
    name: str
    prompt_text: str
    left_in: float
    top_in: float
    width_in: float
    height_in: float


@dataclass(frozen=True, slots=True)
class LayoutInfo:
    index: int
    name: str
    placeholders: tuple[LayoutPlaceholder, ...]


def _safe_inches(value: object) -> float:
    """
    Convert an EMU-like value to inches as float.

    python-pptx returns shape coordinates as Emu, which provides `.inches`.
    Some placeholders may have missing geometry; in that case return 0.0.
    """

    try:
        inches = float(getattr(value, "inches"))
    except Exception:
        return 0.0
    # Keep reasonably compact for prompt payloads.
    return round(inches, 2)


def _ph_type_name(ph) -> str:
    try:
        t = ph.placeholder_format.type
    except Exception:
        return "UNKNOWN"
    try:
        # python-pptx formats this as e.g. "TITLE (1)".
        raw = str(t).strip()
        head = raw.split(".", 1)[-1]
        head = head.split("(", 1)[0].strip()
        head = head.split(" ", 1)[0].strip()
        return head or "UNKNOWN"
    except Exception:
        return "UNKNOWN"


def read_layouts(pptx_path: Path) -> list[LayoutInfo]:
    """
    Read all slide layouts and their placeholders from a pptx deck.

    Raises on invalid pptx / IO errors; callers should decide whether to
    fallback to default behavior.
    """

    prs = Presentation(str(Path(pptx_path)))
    out: list[LayoutInfo] = []

    for i, layout in enumerate(prs.slide_layouts):
        placeholders: list[LayoutPlaceholder] = []
        for ph in layout.placeholders:
            try:
                idx = int(ph.placeholder_format.idx)
            except Exception:
                # If idx is not accessible we cannot address the placeholder reliably.
                continue

            try:
                name = str(getattr(ph, "name", "")) or ""
            except Exception:
                name = ""

            prompt_text = ""
            try:
                tf = getattr(ph, "text_frame", None)
                if tf is not None:
                    prompt_text = (tf.text or "").strip()
            except Exception:
                prompt_text = ""

            placeholders.append(
                LayoutPlaceholder(
                    idx=idx,
                    type=_ph_type_name(ph),
                    name=name,
                    prompt_text=prompt_text,
                    left_in=_safe_inches(getattr(ph, "left", None)),
                    top_in=_safe_inches(getattr(ph, "top", None)),
                    width_in=_safe_inches(getattr(ph, "width", None)),
                    height_in=_safe_inches(getattr(ph, "height", None)),
                )
            )

        out.append(
            LayoutInfo(
                index=i,
                name=str(getattr(layout, "name", "")) or "",
                placeholders=tuple(placeholders),
            )
        )

    return out


def format_layout_card_full(li: LayoutInfo) -> str:
    lines = [
        f'### Layout #{li.index} — \"{li.name}\"',
        "Placeholders:",
    ]
    if not li.placeholders:
        lines.append("- (no placeholders)")
        return "\n".join(lines)

    for ph in li.placeholders:
        lines.append(
            (
                f'- idx={ph.idx} type={ph.type} name=\"{ph.name}\" '
                f'prompt=\"{ph.prompt_text}\" '
                f'bbox=({ph.left_in},{ph.top_in},{ph.width_in}x{ph.height_in}) in'
            )
        )
    return "\n".join(lines)


def format_layout_card_compact(li: LayoutInfo) -> str:
    types = ", ".join(ph.type for ph in li.placeholders) if li.placeholders else ""
    if not types:
        types = "(none)"
    return (
        f'- #{li.index} \"{li.name}\" — placeholders: {len(li.placeholders)} ({types})'
    )


def format_layouts_catalog_compact(items: list[LayoutInfo]) -> str:
    if not items:
        return "(no layouts)"
    return "\n".join(format_layout_card_compact(x) for x in items)


def format_layouts_catalog_with_descriptions(
    items: list[LayoutInfo],
    descriptions: dict[int, str] | None,
) -> str:
    """
    Like ``format_layouts_catalog_compact`` but appends a short ``desc:`` line per layout
    when a non-empty description exists for that index.
    """

    if not items:
        return "(no layouts)"
    if not descriptions:
        return format_layouts_catalog_compact(items)
    out: list[str] = []
    for li in items:
        line = format_layout_card_compact(li)
        d = (descriptions.get(li.index) or "").strip()
        if d:
            line = f"{line}  desc: {d}"
        out.append(line)
    return "\n".join(out)

