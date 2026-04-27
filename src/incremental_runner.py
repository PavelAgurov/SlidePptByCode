"""Unified incremental slide pipeline (one loop for preamble + H2 slides)."""

from __future__ import annotations

import ast
import hashlib
import importlib.util
import logging
import shutil
import traceback as _tb
from pathlib import Path

from pptx import Presentation

from src.code_executor import CodeExecutor
from src.code_generator import CodeGenerator
from src.config import Settings
from src.layout_catalog import LayoutInfo, read_layouts
from src.layout_describer import ensure_layout_descriptions
from src.models import LayoutDescription
from src.ppt_bootstrap import init_deck
from src.script_inject import inject_slide_paths
from src.shared_default import default_shared_symbol_names
from src.shared_index import build_shared_index
from src.task_chunker import TaskChunk, extract_deck_title, split_into_chunks
from src.validator import (
    prefix_slide_digests,
    validate_deck_after_append,
    validate_presentation,
    validate_scratch_append,
)

logger = logging.getLogger(__name__)


def _sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def resolve_deck_path(output_filename: str | None) -> Path:
    """Final deck path under ``.output`` (default name if None)."""
    if output_filename:
        return Path(output_filename)
    return Path(".output") / "presentation.pptx"


def _resolve_forced_layout(
    layouts: list[LayoutInfo], name: str | None, *, flag: str
) -> tuple[LayoutInfo | None, str | None]:
    """Return (layout, error). ``layout`` is None when ``name`` is None."""
    if not name:
        return None, None
    target = name.strip()
    match = next((li for li in layouts if li.name == target), None)
    if match is None:
        available = ", ".join(repr(li.name) for li in layouts if li.name)
        return None, (
            f"{flag} layout name not found: {name!r}. "
            f"Available layout names: {available}"
        )
    return match, None


def _verify_shared_module(shared_path: Path) -> str | None:
    """Return None on success or a traceback string describing the failure."""
    try:
        text = shared_path.read_text(encoding="utf-8")
    except OSError as e:
        return f"OSError reading shared.py: {e}"
    try:
        ast.parse(text, filename=str(shared_path))
    except SyntaxError as e:
        return f"SyntaxError in shared.py: {e}"
    try:
        spec = importlib.util.spec_from_file_location(
            "orch_shared_verify", str(shared_path)
        )
        if spec is None or spec.loader is None:
            return "importlib could not build a spec for shared.py"
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
    except Exception:  # noqa: BLE001
        return "ImportError while exec_module shared.py:\n" + _tb.format_exc()
    missing = [n for n in default_shared_symbol_names() if not hasattr(mod, n)]
    if missing:
        return (
            "shared.py is missing required default symbols: "
            + ", ".join(missing)
        )
    return None


def _ensure_shared_module(
    *,
    generator: CodeGenerator,
    style_content: str | None,
    shared_path: Path,
    max_retries: int,
    error_log: list[str],
) -> bool:
    """
    Generate ``shared.py``, write to disk, and verify it is importable and
    contains all default symbols. Retries via ``fix_shared_module``.
    """
    code = generator.extend_shared_module(style_content)
    shared_path.parent.mkdir(parents=True, exist_ok=True)
    for attempt in range(max(1, max_retries)):
        shared_path.write_text(code, encoding="utf-8")
        problem = _verify_shared_module(shared_path)
        if problem is None:
            logger.info(
                "shared.py verified (attempt %s/%s) at %s",
                attempt + 1,
                max_retries,
                shared_path,
            )
            return True
        error_log.append(
            f"shared.py verification attempt {attempt + 1}: "
            + problem.splitlines()[0]
        )
        if attempt >= max_retries - 1:
            return False
        try:
            code = generator.fix_shared_module(prev_code=code, traceback=problem)
        except Exception as e:  # noqa: BLE001
            error_log.append(f"fix_shared_module failed: {e}")
            return False
    return False


def run_incremental_pipeline(
    *,
    task_content: str,
    style_content: str | None,
    language: str | None,
    output_filename: str | None,
    generator: CodeGenerator,
    executor: CodeExecutor,
    config: Settings,
    max_retries: int,
    slide_max: int | None = None,
    template_pptx: Path | None = None,
    forced_layout_h1: str | None = None,
    forced_layout_h2: str | None = None,
) -> tuple[bool, Path | None, list[str]]:
    """
    Run the unified slide-by-slide generation: prepare a 0-slide deck,
    generate ``shared.py`` once, then loop over chunks ``[preamble, *h2_to_process]``
    appending exactly one slide per chunk.

    Returns:
        (success, deck_path_or_none, error_log)
    """
    error_log: list[str] = []
    chunks = split_into_chunks(task_content)
    h2_all = [c for c in chunks if c.kind == "h2"]
    preamble = next(c for c in chunks if c.kind == "preamble")
    if not h2_all:
        error_log.append("No H2 sections after validation")
        return False, None, error_log

    h2_to_process = h2_all if slide_max is None else h2_all[:slide_max]
    partial_run = slide_max is not None and len(h2_to_process) < len(h2_all)

    deck_title = extract_deck_title(task_content)
    if not deck_title:
        error_log.append("Missing H1 deck title")
        return False, None, error_log

    deck_path = resolve_deck_path(output_filename)
    deck_path.parent.mkdir(parents=True, exist_ok=True)

    shared_path = (config.generated_dir / "shared.py").resolve()
    scratch_dir = (config.output_dir / "_scratch").resolve()
    scratch_dir.mkdir(parents=True, exist_ok=True)

    if template_pptx is not None:
        tpl = Path(template_pptx).resolve()
        if not tpl.is_file():
            error_log.append(f"Template not found: {tpl}")
            return False, None, error_log
        if deck_path.resolve() == tpl:
            error_log.append(
                "Output deck path equals template path; use a different --output "
                "so the template file is not overwritten."
            )
            return False, None, error_log

    try:
        n_slides = init_deck(deck_path, template_pptx)
        logger.info(
            "Initialized deck at %s (%s slides; template=%s)",
            deck_path,
            n_slides,
            template_pptx if template_pptx else "<default python-pptx>",
        )
    except OSError as e:
        error_log.append(f"Failed to initialize deck: {e}")
        return False, None, error_log
    except Exception as e:  # noqa: BLE001
        error_log.append(f"init_deck failed: {e}")
        return False, None, error_log

    try:
        layouts: list[LayoutInfo] = read_layouts(deck_path)
    except Exception as e:  # noqa: BLE001
        error_log.append(f"Cannot read layouts from deck: {e}")
        return False, None, error_log
    if not layouts:
        error_log.append("Deck has no master slide layouts")
        return False, None, error_log

    forced_h1, err = _resolve_forced_layout(
        layouts, forced_layout_h1, flag="--template_layout_h1"
    )
    if err:
        error_log.append(err)
        return False, None, error_log
    forced_h2, err = _resolve_forced_layout(
        layouts, forced_layout_h2, flag="--template_layout_h2"
    )
    if err:
        error_log.append(err)
        return False, None, error_log

    layout_descriptions: dict[int, LayoutDescription] = ensure_layout_descriptions(
        generator=generator,
        template_pptx=Path(template_pptx).resolve() if template_pptx else None,
        layouts=layouts,
        generated_dir=config.generated_dir,
    )

    if not _ensure_shared_module(
        generator=generator,
        style_content=style_content,
        shared_path=shared_path,
        max_retries=max_retries,
        error_log=error_log,
    ):
        error_log.append("shared.py could not be generated/verified within retries")
        return False, None, error_log

    try:
        locked_shared_fp = _sha256_file(shared_path)
    except OSError as e:
        error_log.append(f"Cannot hash shared.py: {e}")
        return False, None, error_log

    try:
        shared_index = build_shared_index(shared_path)
    except (OSError, SyntaxError, UnicodeError) as e:
        error_log.append(f"Cannot build shared.py symbol index: {e}")
        return False, None, error_log

    _ALLOWED_SLIDE_TYPES: dict[str, frozenset[str]] = {
        "preamble": frozenset({"header"}),
        "h2": frozenset({"content"}),
    }

    def _select_layout(slide_markdown: str, kind: str) -> LayoutInfo | None:
        allowed_types = _ALLOWED_SLIDE_TYPES[kind]
        filtered = [
            li
            for li in layouts
            if (ld := layout_descriptions.get(li.index)) is not None
            and ld.slide_type in allowed_types
            and (kind != "h2" or not ld.slide_has_image_placeholder)
        ]
        if not filtered:
            return None
        desc_subset = {li.index: layout_descriptions[li.index] for li in filtered}
        sel = generator.select_layout(
            slide_markdown=slide_markdown,
            deck_title=deck_title,
            layouts=filtered,
            style_content=style_content,
            language=language,
            descriptions=desc_subset,
        )
        idx = int(sel.selected_layout_index)
        pick = next((li for li in filtered if li.index == idx), None)
        if pick is not None:
            return pick
        logger.warning(
            "Layout selection out of range (%s); retrying once",
            idx,
        )
        sel2 = generator.select_layout(
            slide_markdown=(
                slide_markdown
                + "\n\n(Your previous selected_layout_index was out of range; choose only from the allowed values.)"
            ),
            deck_title=deck_title,
            layouts=filtered,
            style_content=style_content,
            language=language,
            descriptions=desc_subset,
        )
        idx2 = int(sel2.selected_layout_index)
        pick2 = next((li for li in filtered if li.index == idx2), None)
        if pick2 is not None:
            return pick2
        logger.warning(
            "Layout selection out of range after retry (%s)", idx2
        )
        return None

    # Unified loop: preamble first, then each H2 chunk -> one appended slide.
    pipeline_chunks: list[tuple[TaskChunk, str, LayoutInfo | None]] = []
    pipeline_chunks.append((preamble, "preamble", forced_h1))
    for c in h2_to_process:
        pipeline_chunks.append((c, "h2", forced_h2))
    total_slides = len(pipeline_chunks)

    for ord1, (chunk, kind, forced) in enumerate(pipeline_chunks, start=1):
        logger.debug("Slide %s/%s (kind=%s)", ord1, total_slides, kind)

        deck_slide_count_before = len(Presentation(str(deck_path)).slides)
        prefix_before = prefix_slide_digests(deck_path, deck_slide_count_before)

        if forced is not None:
            chosen_layout: LayoutInfo | None = forced
        else:
            chosen_layout = _select_layout(chunk.body, kind)
            if chosen_layout is None:
                # No safe fallback exists for the preamble (Title Slide layouts
                # have no BODY placeholder, so pick_title_and_content_layout
                # would crash). Treat as a generation error and let the outer
                # retry (here, just log and fail this slide) rerun selection.
                if kind == "preamble":
                    error_log.append(
                        "Preamble layout selection failed twice; aborting "
                        "(no safe default fallback)."
                    )
                    return False, None, error_log
                logger.info(
                    "Slide %s/%s: layout selection failed; using default layout picking via snippets",
                    ord1,
                    total_slides,
                    extra={"color_event": "layout_select"},
                )
        chosen_layout_index = (
            chosen_layout.index if chosen_layout is not None else None
        )
        if chosen_layout is not None:
            logger.info(
                "Slide %s/%s (%s): chosen layout #%s %r",
                ord1,
                total_slides,
                kind,
                chosen_layout.index,
                chosen_layout.name,
                extra={"color_event": "layout_select"},
            )

        slide_code = generator.generate_incremental_slide(
            chunk.body,
            deck_title,
            ord1,
            total_slides,
            style_content,
            language,
            shared_index=shared_index,
            chosen_layout=chosen_layout,
        ).code

        deck_backup = deck_path.with_suffix(deck_path.suffix + ".bak")

        for attempt in range(max_retries):
            scratch_path = scratch_dir / f"slide_{ord1:03d}_try{attempt + 1}.pptx"
            shutil.copy2(deck_path, scratch_path)
            baseline = deck_slide_count_before

            scr_script = inject_slide_paths(
                slide_code,
                scratch_path,
                shared_path,
                chosen_layout_index=chosen_layout_index,
            )
            scr_file = executor.save_code(
                scr_script, f"incremental_slide_{ord1:03d}_scratch_{attempt + 1}.py"
            )
            scr_res = executor.execute(
                scr_file, expected_output_pptx=scratch_path
            )
            if not scr_res.success:
                msg = scr_res.error_message or "Scratch execution failed"
                error_log.append(
                    f"Slide {ord1} ({kind}) scratch attempt {attempt + 1}: {msg}"
                )
                if attempt >= max_retries - 1:
                    return False, None, error_log
                slide_code = generator.fix_code_after_error(
                    slide_code,
                    scr_res,
                    incremental=True,
                    shared_index=shared_index,
                    chosen_layout=chosen_layout,
                ).code
                continue

            vs = validate_scratch_append(scratch_path, baseline, chunk)
            if not vs.is_valid:
                err_text = "; ".join(vs.issues)
                error_log.append(
                    f"Slide {ord1} ({kind}) scratch validation attempt {attempt + 1}: {err_text}"
                )
                if attempt >= max_retries - 1:
                    return False, None, error_log
                slide_code = generator.fix_incremental_h2_validation(
                    slide_code,
                    vs.issues,
                    deck_title,
                    chunk.body,
                    shared_index=shared_index,
                    chosen_layout=chosen_layout,
                ).code
                continue

            shutil.copy2(deck_path, deck_backup)
            deck_script = inject_slide_paths(
                slide_code,
                deck_path,
                shared_path,
                chosen_layout_index=chosen_layout_index,
            )
            deck_file = executor.save_code(
                deck_script, f"incremental_slide_{ord1:03d}_deck_{attempt + 1}.py"
            )
            deck_res = executor.execute(deck_file, expected_output_pptx=deck_path)
            if not deck_res.success:
                shutil.copy2(deck_backup, deck_path)
                msg = deck_res.error_message or "Deck execution failed"
                error_log.append(
                    f"Slide {ord1} ({kind}) deck attempt {attempt + 1}: {msg}"
                )
                if attempt >= max_retries - 1:
                    return False, None, error_log
                slide_code = generator.fix_code_after_error(
                    slide_code,
                    deck_res,
                    incremental=True,
                    shared_index=shared_index,
                    chosen_layout=chosen_layout,
                ).code
                continue

            vd = validate_deck_after_append(
                deck_path, deck_slide_count_before, prefix_before, chunk
            )
            if not vd.is_valid:
                shutil.copy2(deck_backup, deck_path)
                err_text = "; ".join(vd.issues)
                error_log.append(
                    f"Slide {ord1} ({kind}) deck validation attempt {attempt + 1}: {err_text}"
                )
                if attempt >= max_retries - 1:
                    return False, None, error_log
                slide_code = generator.fix_incremental_h2_validation(
                    slide_code,
                    vd.issues,
                    deck_title,
                    chunk.body,
                    shared_index=shared_index,
                    chosen_layout=chosen_layout,
                ).code
                continue

            try:
                if _sha256_file(shared_path) != locked_shared_fp:
                    shutil.copy2(deck_backup, deck_path)
                    error_log.append(
                        f"Slide {ord1}: shared.py changed during slide generation (forbidden)"
                    )
                    if attempt >= max_retries - 1:
                        return False, None, error_log
                    slide_code = generator.fix_incremental_h2_validation(
                        slide_code,
                        ["shared.py must not be modified"],
                        deck_title,
                        chunk.body,
                        shared_index=shared_index,
                        chosen_layout=chosen_layout,
                    ).code
                    continue
            except OSError as e:
                shutil.copy2(deck_backup, deck_path)
                error_log.append(f"Slide {ord1}: cannot verify shared.py: {e}")
                return False, None, error_log

            try:
                deck_backup.unlink(missing_ok=True)
            except OSError:
                pass
            break
        else:
            return False, None, error_log

    if partial_run:
        logger.info(
            "Partial run (--slide_max): skipping full-deck validation against task spec"
        )
        return True, deck_path, error_log

    final_val = validate_presentation(deck_path, task_content)
    if not final_val.is_valid:
        error_log.append(
            "Final validation: " + "; ".join(final_val.issues)
        )
        return False, deck_path, error_log

    return True, deck_path, error_log
