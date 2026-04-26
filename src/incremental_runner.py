"""Incremental H1 + per-H2 slide orchestration (deterministic paths, scratch then deck)."""

from __future__ import annotations

import hashlib
import logging
import shutil
from pathlib import Path

from pptx import Presentation

from src.code_executor import CodeExecutor
from src.code_generator import CodeGenerator
from src.config import Settings
from src.ppt_bootstrap import copy_deck_template, create_empty_ppt
from src.script_inject import inject_h1_paths, inject_h2_slide_paths
from src.task_chunker import extract_deck_title, split_into_chunks
from src.shared_index import build_shared_index
from src.layout_catalog import LayoutInfo, read_layouts
from src.layout_describer import ensure_layout_descriptions
from src.validator import (
    prefix_slide_digests,
    validate_deck_after_append,
    validate_incremental_h1_deck,
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
    template_layout: str | None = None,
) -> tuple[bool, Path | None, list[str]]:
    """
    Run incremental generation: H1 (deck stub + shared), then each H2 slide.

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
    num_h2_total = len(h2_all)
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

    deck_from_template = template_pptx is not None
    if template_layout and not deck_from_template:
        error_log.append("--template_layout requires --template")
        return False, None, error_log
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
            n_slides = copy_deck_template(tpl, deck_path)
            logger.info(
                "Initialized deck from template %s (%s slides)", tpl, n_slides
            )
        except OSError as e:
            error_log.append(f"Failed to copy template deck: {e}")
            return False, None, error_log
        except Exception as e:
            error_log.append(f"Template deck is not a valid .pptx: {e}")
            return False, None, error_log
    else:
        create_empty_ppt(deck_path)

    layouts: list[LayoutInfo] = []
    if deck_from_template:
        try:
            layouts = read_layouts(deck_path)
        except Exception as e:
            logger.warning(
                "read_layouts failed (%s); proceeding without layout selection",
                e,
            )
            layouts = []
    use_layout_selection = bool(layouts)

    forced_layout: LayoutInfo | None = None
    if template_layout:
        if not layouts:
            error_log.append(
                f"--template_layout was provided ({template_layout!r}) but no layouts could be read from the template"
            )
            return False, None, error_log
        target = template_layout.strip()
        forced_layout = next((li for li in layouts if li.name == target), None)
        if forced_layout is None:
            available = ", ".join(repr(li.name) for li in layouts if li.name)
            error_log.append(
                f"Template layout name not found: {template_layout!r}. Available layout names: {available}"
            )
            return False, None, error_log
        logger.info(
            "Forcing template layout for all H2 slides: #%s %r",
            forced_layout.index,
            forced_layout.name,
            extra={"color_event": "layout_select"},
        )
        use_layout_selection = False

    layout_descriptions: dict[int, str] = {}
    if use_layout_selection and template_pptx is not None:
        layout_descriptions = ensure_layout_descriptions(
            generator=generator,
            template_pptx=Path(template_pptx).resolve(),
            layouts=layouts,
            generated_dir=config.generated_dir,
        )

    def _select_layout_with_validation(slide_markdown: str) -> LayoutInfo | None:
        if not use_layout_selection:
            return None
        allowed = list(range(len(layouts)))
        sel = generator.select_layout(
            slide_markdown=slide_markdown,
            deck_title=deck_title,
            layouts=layouts,
            style_content=style_content,
            language=language,
            descriptions=layout_descriptions,
        )
        idx = int(sel.selected_layout_index)
        if 0 <= idx < len(layouts):
            return layouts[idx]

        logger.warning(
            "Layout selection out of range (%s), retrying once. Allowed: %s",
            idx,
            allowed,
        )
        sel2 = generator.select_layout(
            slide_markdown=(
                slide_markdown
                + "\n\n(Your previous selected_layout_index was out of range; choose only from the allowed values.)"
            ),
            deck_title=deck_title,
            layouts=layouts,
            style_content=style_content,
            language=language,
            descriptions=layout_descriptions,
        )
        idx2 = int(sel2.selected_layout_index)
        if 0 <= idx2 < len(layouts):
            return layouts[idx2]
        logger.warning(
            "Layout selection out of range after retry (%s); falling back to default layout picking.",
            idx2,
        )
        return None

    # --- H1 ---
    h1_code = generator.generate_incremental_h1(
        preamble.body,
        deck_title,
        style_content,
        language,
        deck_from_template=deck_from_template,
    ).code

    for attempt in range(max_retries):
        logger.debug("H1 attempt %s/%s (execute/validate loop)", attempt + 1, max_retries)
        script = inject_h1_paths(h1_code, deck_path, shared_path)
        code_file = executor.save_code(script, f"incremental_h1_{attempt + 1}.py")
        result = executor.execute(code_file, expected_output_pptx=deck_path)
        if not result.success:
            msg = result.error_message or "H1 execution failed"
            error_log.append(f"H1 attempt {attempt + 1}: {msg}")
            if attempt >= max_retries - 1:
                return False, None, error_log
            h1_code = generator.fix_code_after_error(
                h1_code, result, incremental=True
            ).code
            continue

        v = validate_incremental_h1_deck(deck_path, shared_path, deck_title)
        if v.is_valid:
            break
        err = "H1 validation: " + "; ".join(v.issues)
        error_log.append(f"H1 attempt {attempt + 1}: {err}")
        if attempt >= max_retries - 1:
            return False, None, error_log
        h1_code = generator.fix_incremental_h1_validation(
            h1_code, v.issues, preamble.body
        ).code
    else:
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

    # --- H2 slides ---
    for ord1, chunk in enumerate(h2_to_process, start=1):
        logger.debug("H2 slide %s/%s (orchestrator step)", ord1, num_h2_total)

        deck_slide_count_before = len(Presentation(str(deck_path)).slides)
        prefix_before = prefix_slide_digests(deck_path, deck_slide_count_before)

        chosen_layout = _select_layout_with_validation(chunk.body)
        if forced_layout is not None:
            chosen_layout = forced_layout
        chosen_layout_index = chosen_layout.index if chosen_layout is not None else None
        if chosen_layout is not None:
            logger.info(
                "Slide %s/%s: chosen layout #%s %r",
                ord1,
                num_h2_total,
                chosen_layout.index,
                chosen_layout.name,
                extra={"color_event": "layout_select"},
            )
        elif use_layout_selection:
            logger.info(
                "Slide %s/%s: layout selection failed; using default layout picking",
                ord1,
                num_h2_total,
                extra={"color_event": "layout_select"},
            )

        slide_code = generator.generate_incremental_slide(
            chunk.body,
            deck_title,
            ord1,
            num_h2_total,
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

            scr_script = inject_h2_slide_paths(
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
                    f"Slide {ord1} scratch attempt {attempt + 1}: {msg}"
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
                err = "; ".join(vs.issues)
                error_log.append(
                    f"Slide {ord1} scratch validation attempt {attempt + 1}: {err}"
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
            deck_script = inject_h2_slide_paths(
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
                    f"Slide {ord1} deck attempt {attempt + 1}: {msg}"
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
                err = "; ".join(vd.issues)
                error_log.append(
                    f"Slide {ord1} deck validation attempt {attempt + 1}: {err}"
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
