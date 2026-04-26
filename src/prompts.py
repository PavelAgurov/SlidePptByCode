"""System prompts and templates for LLM interactions."""

from .code_merger import CHUNK_MERGE_MARKER
from .layout_catalog import LayoutInfo, format_layout_card_full, format_layouts_catalog_with_descriptions
from .snippets import formatted_snippet_catalog_for_prompt

_ERROR_PROMPT_CODE_MAX = 16_000


def _tail_if_too_long(text: str, max_len: int) -> str:
    """Keep the end of ``text`` (tracebacks and errors often end with the exception line)."""
    if len(text) <= max_len:
        return text
    marker = f"... [truncated from {len(text)} chars; showing end]\n"
    budget = max_len - len(marker)
    return marker + text[-budget:]


_SNIPPET_TOOL_RULES = """
TOOL ``get_code_snippet``:
- Before writing non-trivial python-pptx code (slides, bullets, colors, charts, **tables**), call
  ``get_code_snippet`` with the matching ``snippet_id`` from the catalog below.
- Reuse snippet text from the conversation; do not re-request the same id unless you need it again.
- Do not invent python-pptx APIs; align with the snippets you fetched.

""" + formatted_snippet_catalog_for_prompt()


SYSTEM_PROMPT = """You are an expert Python developer specializing in creating PowerPoint presentations using the python-pptx library.

Your task is to generate complete, executable Python code that creates a PowerPoint presentation based on the provided specification.

REQUIREMENTS:
1. The code must be self-contained and immediately executable
2. Save the output .pptx file to the '.output/' directory
3. Use descriptive filename (e.g., 'company_q4_results.pptx')
4. Follow python-pptx best practices
5. Handle errors gracefully
6. When style guidelines are provided, apply colors and formatting consistently throughout the presentation
7. Use only the Python standard library plus python-pptx (and its normal dependencies). Do not import requests, Pillow, matplotlib, or other third-party packages for HTTP, downloads, or plotting unless the task text explicitly requires them. Prefer python-pptx shapes, text, and colors for illustrations

STRUCTURE YOUR CODE:
- Use ``from pptx import Presentation``, ``from pptx.util import Inches, Pt``, ``from pptx.dml.color import RGBColor`` as needed.
- Create ``prs = Presentation()``, set optional slide dimensions, add slides, then save under ``Path('.output') / 'your_file.pptx'`` with ``output_path.parent.mkdir(exist_ok=True)`` and ``prs.save(output_path)``.
- End with ``if __name__ == '__main__':`` calling your entry function and ``print`` a success line including the saved path.

""" + _SNIPPET_TOOL_RULES + """

WHEN STYLE GUIDELINES ARE PROVIDED:
1. Parse HEX colors (e.g., #76CCBE) and convert to RGBColor via the hex_to_rgb snippet pattern
2. Apply primary colors for backgrounds where the guide specifies
3. Use accent colors sparingly for emphasis
4. Use secondary colors for highlights, warnings, or data visualization
5. Use grey scale for text and functional elements when appropriate
6. Follow any "DO NOT" rules from the style guide
7. Maintain brand consistency throughout all slides

CHART PATTERNS (text rules — use tool snippets for code):

Charts are created from CSV-like data in the task specification. The task may specify chart type in two ways:
1. **Directive format**: A line like "chart: line" or "chart: pie" (takes priority)
2. **Free text**: Natural language like "Display as a line chart" or "Покажи как круговую диаграмму"
3. **Auto-detect**: If neither is specified, choose the most appropriate chart type based on data shape

Supported chart directives:
- chart: column → XL_CHART_TYPE.COLUMN_CLUSTERED (default for multi-series)
- chart: column_stacked → XL_CHART_TYPE.COLUMN_STACKED
- chart: bar → XL_CHART_TYPE.BAR_CLUSTERED
- chart: line → XL_CHART_TYPE.LINE
- chart: line_markers → XL_CHART_TYPE.LINE_MARKERS
- chart: pie → XL_CHART_TYPE.PIE (default for single-series proportions)
- chart: doughnut → XL_CHART_TYPE.DOUGHNUT
- chart: area → XL_CHART_TYPE.AREA
- chart: scatter → XL_CHART_TYPE.XY_SCATTER

Important chart notes:
- Use blank layout (6) for chart slides to have full control
- Add title as text box, not title placeholder
- Parse CSV data carefully - handle headers and data rows
- Apply style guide colors to chart series when provided
- Typical chart dimensions: 8x5 inches for full-width charts
- Position: x=1", y=1.5", title at y=0.5"

IMPORTANT NOTES:
- Always create the .output directory if it doesn't exist
- Use pathlib.Path for cross-platform compatibility
- Include a main guard: if __name__ == '__main__'
- Add a print statement when saving successfully
- Handle multiple slides as specified in the task
- Use appropriate slide layouts (0=title, 1=title+content, 6=blank)
- Parse the task carefully to extract slide count and content
- For chart slides: use blank layout (6), parse CSV data, look for "chart:" directives, apply style colors to series
- Chart slide titles should be text boxes, not title placeholders
- If the task markdown contains a **pipe table** (rows with ``|`` columns), render it as a real table using
  ``slide.shapes.add_table`` (snippet ``table_markdown_grid``), not as a single text box or bullet list of raw ``|`` characters

Generate clean, well-structured code that follows these guidelines."""


ERROR_FIX_PROMPT_TEMPLATE = """The code you generated encountered an error during execution.

ORIGINAL CODE:
```python
{original_code}
```

ERROR DETAILS:
Error Message: {error_message}

FULL TRACEBACK:
{traceback}

STDOUT:
{stdout}

STDERR:
{stderr}

INSTRUCTIONS:
1. Analyze the error carefully
2. Identify the root cause
3. Generate corrected code that fixes the specific issue
4. Maintain the overall structure and intent
5. Ensure the fix doesn't introduce new issues
6. Preserve any style guidelines and color formatting that were applied

Common issues to check:
- Missing imports
- Incorrect method calls or parameters
- Index errors (accessing non-existent slides/shapes)
- Type errors (wrong data types)
- File path issues
- None reference errors
- Layout placeholder issues

Generate the complete corrected code (not just the fix)."""


INCREMENTAL_EXECUTION_ERROR_FIX_TEMPLATE = """EXECUTION FAILED — produce a corrected full script in the structured ``code`` field.

ERROR MESSAGE:
{error_message}

FULL TRACEBACK:
{traceback}

STDOUT:
{stdout}

STDERR:
{stderr}

{shared_block}

INSTRUCTIONS:
1. Fix the root cause shown in the traceback.
2. Call ``get_code_snippet`` if you need canonical python-pptx examples.
3. For incremental H2: use attribute names from the shared symbol index above exactly — do not invent names.
4. If the traceback mentions ``KeyError`` / ``no placeholder on this slide with idx`` / ``placeholders[...]``, stop using numeric placeholder indices: pick a title+content layout and resolve the body via ``PP_PLACEHOLDER_TYPE`` (see ``get_code_snippet`` with ``incremental_h2_skeleton`` or ``content_slide``).
5. Return the complete corrected script (entire file body you would save as .py), not a minimal diff.

ORIGINAL SCRIPT (may be truncated from the start if very long):
```python
{original_code}
```
"""


VALIDATION_FIX_PROMPT_TEMPLATE = """The code executed successfully but the generated presentation failed validation.

ORIGINAL CODE:
```python
{original_code}
```

VALIDATION ISSUES:
- Expected slides: {expected_slides}
- Actual slides: {actual_slides}
- Has titles: {has_titles}
- Issues found: {issues}

TASK SPECIFICATION:
{task_content}

INSTRUCTIONS:
1. Review the task specification carefully
2. Count the expected number of slides (look for "## Slide N" markers)
3. Ensure each slide has appropriate content
4. Fix any missing slides or incorrect content
5. Ensure all slides have titles
6. Preserve any style guidelines and color formatting that were applied in the original code

Generate the complete corrected code that matches the task specification."""


CHUNK_MERGE_MARKER_LINE = f"    {CHUNK_MERGE_MARKER}"


def _append_chunk_style_language(
    msg: str,
    style_content: str | None,
    language: str | None,
    *,
    style_apply_line: str,
) -> str:
    if style_content:
        msg += f"\n\n## STYLE GUIDELINES:\n\n{style_content}\n\n{style_apply_line}"
    if language:
        msg += f"\n\n## LANGUAGE REQUIREMENT:\n\nAll visible text MUST be in {language}."
    return msg


def format_chunk_title_deck_user_message(
    title_block_markdown: str,
    deck_title: str,
    total_h2_slides: int,
    style_content: str | None,
    language: str | None,
    output_filename: str | None,
) -> str:
    """
    Part 1 of chunked build: imports, helpers, create_presentation shell, and title/H1 block only.

    ``title_block_markdown`` is everything before the first ``##`` (includes the ``#`` deck line).
    """
    parts = [
        "You are generating **PART 1 of 2+** of a python-pptx script (chunked build). "
        f"Later parts add **{total_h2_slides}** content slide(s): one LLM pass per `##` heading, "
        "each merged as `add_section_NNN(prs)`.",
        "",
        "ROLE — TITLE DECK (H1 block only):",
        "- This part covers ONLY the presentation title from the `# ...` line and any markdown "
        "between that H1 and the first `##` slide section. Do **not** implement body slides for `##` sections here.",
        "",
        "HARD CONTRACT:",
        "1. Include all imports and helper functions needed for the **entire** deck (charts, tables, etc. in later parts).",
        "2. Define `def add_preamble_slides(prs):` that adds ONLY slide(s) for the title block below: "
        "use the H1 text as the deck title and place any intervening text (before the first `##`) on the title slide "
        "or as minimal supporting layout — still within the title/opening block, not H2 content slides.",
        "3. Define `def create_presentation():` that:",
        "   - Creates `prs = Presentation()`, sets slide dimensions,",
        "   - Calls `add_preamble_slides(prs)`,",
        "   - Then contains a single line with **exactly** this comment (four spaces before `#`):",
        f"      {CHUNK_MERGE_MARKER_LINE.strip()}",
        "   - Then sets `output_path` (pathlib.Path), ensures parent dir exists, `prs.save(output_path)`, prints success.",
        "4. End with `if __name__ == '__main__': create_presentation()`.",
        "5. Do NOT define or call `add_section_*` — those are merged from later parts.",
        "",
        f"Presentation title (H1): {deck_title}",
        f"Content slides to be generated in later parts (`##` count): {total_h2_slides}",
        "",
        "## TITLE-BLOCK MARKDOWN (H1 through line before first ##):",
        title_block_markdown,
    ]
    msg = "\n".join(parts)
    msg = _append_chunk_style_language(
        msg,
        style_content,
        language,
        style_apply_line="Apply these consistently for the whole deck.",
    )
    if output_filename:
        msg += (
            f"\n\n## OUTPUT FILENAME:\n\nSave to exactly: {output_filename}\n"
            "Use this path in create_presentation."
        )
    return msg


def format_chunk_h2_slide_user_message(
    section_markdown: str,
    function_name: str,
    part_number: int,
    total_parts: int,
    section_ordinal: int,
    num_h2_slides: int,
    deck_title: str,
    style_content: str | None,
    language: str | None,
) -> str:
    """Part k>1: exactly one `def add_section_NNN(prs):` implementing one `##` content slide."""
    parts = [
        f"You are generating **PART {part_number} of {total_parts}** (chunked build).",
        "",
        "ROLE — CONTENT SLIDE (single `##` section):",
        "- Add **exactly one** content slide for this `##` block. This is not the opening title slide from part 1; "
        "do not recreate the full title layout unless the section text clearly requires a duplicate.",
        f"- Deck name for context only (footer, small label, branding if needed): **{deck_title}**.",
        "",
        "HARD CONTRACT:",
        f"1. Output **only** one function, named **exactly** `{function_name}(prs):` (same spelling).",
        "2. Inside it, add **exactly one** slide: map the `##` line to the slide title; put bullets, tables-as-text, "
        "charts, or body copy from the section markdown on that slide.",
        "3. Do NOT create `Presentation()`, do NOT call `prs.save`, do NOT add `if __name__`.",
        "4. You may use `from pptx...` imports at the top of your snippet if required; prefer `prs` and python-pptx APIs.",
        "",
        f"Content slide {section_ordinal} of {num_h2_slides} (document order).",
        "",
        "## SECTION MARKDOWN (one ## slide):",
        section_markdown,
    ]
    msg = "\n".join(parts)
    return _append_chunk_style_language(
        msg,
        style_content,
        language,
        style_apply_line="Match the visual style from part 1 (imports live only in part 1).",
    )


def format_error_fix_prompt(
    original_code: str,
    error_message: str,
    traceback: str,
    stdout: str,
    stderr: str
) -> str:
    """Format the error fix prompt with actual error details."""
    return ERROR_FIX_PROMPT_TEMPLATE.format(
        original_code=original_code,
        error_message=error_message,
        traceback=traceback,
        stdout=stdout or "(empty)",
        stderr=stderr or "(empty)"
    )


def format_incremental_execution_error_fix_prompt(
    original_code: str,
    error_message: str,
    traceback: str,
    stdout: str,
    stderr: str,
    *,
    shared_index: str | None = None,
) -> str:
    """Error-first fix prompt for incremental H1/H2 execution failures."""
    trimmed = _tail_if_too_long(original_code, _ERROR_PROMPT_CODE_MAX)
    if shared_index:
        shared_block = (
            "CURRENT shared.py SYMBOL INDEX (source of truth — match these names exactly):\n"
            "```text\n"
            f"{shared_index}\n"
            "```\n"
        )
    else:
        shared_block = ""
    return INCREMENTAL_EXECUTION_ERROR_FIX_TEMPLATE.format(
        error_message=error_message,
        traceback=traceback or "(no traceback)",
        stdout=stdout or "(empty)",
        stderr=stderr or "(empty)",
        shared_block=shared_block,
        original_code=trimmed,
    )


def format_validation_fix_prompt(
    original_code: str,
    expected_slides: int,
    actual_slides: int,
    has_titles: bool,
    issues: list[str],
    task_content: str
) -> str:
    """Format the validation fix prompt with validation details."""
    return VALIDATION_FIX_PROMPT_TEMPLATE.format(
        original_code=original_code,
        expected_slides=expected_slides,
        actual_slides=actual_slides,
        has_titles=has_titles,
        issues=", ".join(issues) if issues else "None",
        task_content=task_content
    )


# ---------------------------------------------------------------------------
# Incremental pipeline (H1 deck stub + shared.py, then H2 append per slide)
# ---------------------------------------------------------------------------

SYSTEM_PROMPT_INCREMENTAL_H1 = """You are an expert Python developer using python-pptx.

STRUCTURED OUTPUT (critical):
- You return **one** Python module in the ``code`` field: normal ``.py`` source, not markdown, not "here is shared.py" as the only payload.
- The ``code`` string **must** contain an import of python-pptx, e.g. ``from pptx import Presentation`` (and typically ``from pptx.util import Inches, Pt``, etc.) **before** any deck logic.
- Put the shared helper **source** inside your script as a triple-quoted string (or build it with concatenation) and pass it to ``SHARED_PY_PATH.write_text(..., encoding='utf-8')``. Do **not** return the shared module alone without ``from pptx import Presentation`` and deck code.

INCREMENTAL H1 CONTRACT (orchestrator injects path constants at the top of the file):
- Variables ``DECK_PPTX_PATH`` and ``SHARED_PY_PATH`` are already defined (do not redefine).
- Open the existing deck with ``prs = Presentation(str(DECK_PPTX_PATH))``. Do NOT call ``Presentation()`` with no arguments to create a new template for the deck.
- Implement the title / H1 block from the markdown: fill the first slide(s) as appropriate. Prefer editing an existing title slide (or the most suitable existing slide) rather than adding many new slides, unless the task clearly needs more. Do not remove existing slides unless the task explicitly requires it.
- Write the shared style/helper module to ``SHARED_PY_PATH`` using ``SHARED_PY_PATH.write_text(...)`` (UTF-8). This file must be valid Python and importable; put palette helpers (e.g. hex_to_rgb), layout helpers, and any constants used across slides there.
- After writing ``shared.py``, you may load it with importlib from ``SHARED_PY_PATH`` and use it when building the H1 slide(s), or duplicate minimal logic — but the file on disk must exist and be usable by later slide scripts.
- Save only to ``DECK_PPTX_PATH``: ``prs.save(str(DECK_PPTX_PATH))`` (parent directory already exists).
- End with ``if __name__ == '__main__':`` calling a single entry function (e.g. ``main()`` or ``run()``) that performs all steps.
- Use only the Python standard library plus python-pptx. Do not import requests, Pillow, matplotlib, etc., unless the task explicitly requires them.
- Do not print the deck path for orchestration; optional logs are fine.

""" + _SNIPPET_TOOL_RULES + """

Call ``get_code_snippet`` with ``incremental_h1_skeleton`` if you need a minimal end-to-end pattern for this contract."""


SYSTEM_PROMPT_INCREMENTAL_H2 = """You are an expert Python developer using python-pptx.

STRUCTURED OUTPUT (critical):
- The ``code`` field is **one** runnable ``.py`` file. It **must** include ``from pptx import Presentation`` or ``import pptx`` (in addition to any other imports).

INCREMENTAL SLIDE (H2) CONTRACT (orchestrator injects constants at the top of the file):
- ``TARGET_PPTX`` is the path to the presentation to modify (scratch or final deck).
- ``shared`` is already loaded from ``SHARED_PY_PATH`` — use ``shared`` for styling/helpers; do NOT write to ``SHARED_PY_PATH`` and do not redefine it.
- Use **only** attribute names that appear in the ``CURRENT shared.py SYMBOL INDEX`` block in the user message (exact spelling). Never invent names like ``shared.elf_green`` if the index shows ``ELF_GREEN``.
- Open with ``prs = Presentation(str(TARGET_PPTX))``.
- Append **exactly one** new slide at the end: ``prs.slides.add_slide(...)``. Do not remove slides. Do not modify shapes/text on slides whose index is less than the slide count before your addition (append-only for existing slides).
- **Layouts and placeholders:** The same script runs on a **default scratch** deck and on a **template** deck; placeholder **idx** values differ (e.g. ``slide.placeholders[1]`` may raise ``KeyError`` on branded masters). Do **not** assume ``prs.slide_layouts[1]`` or ``slide.placeholders[1]`` for the body. Choose a layout that has both a title-type and a body-type placeholder using ``PP_PLACEHOLDER_TYPE`` (see ``incremental_h2_skeleton`` / ``content_slide`` snippets), then find the body by iterating ``slide.placeholders`` and matching type — not by numeric index.
- If a `## SELECTED LAYOUT` block is present in the user message, use the injected `CHOSEN_LAYOUT_INDEX` constant and address placeholders strictly by their `idx` (from the layout card). When that block is absent, fall back to `pick_title_and_content_layout`.
- If the section markdown contains a **pipe table** (``|`` columns), use ``get_code_snippet`` with ``table_markdown_grid`` and build a native ``Table`` via ``add_table``. Do **not** paste the markdown table into ``body.text`` or a bullet list as plain text.
- Save back to the same file: ``prs.save(str(TARGET_PPTX))``.
- End with ``if __name__ == '__main__':`` calling one entry function.
- Use only the Python standard library plus python-pptx.
- Do not print paths for orchestration; optional logs are fine.

""" + _SNIPPET_TOOL_RULES + """

Call ``get_code_snippet`` with ``incremental_h2_skeleton`` for a minimal append-one-slide pattern."""


SYSTEM_PROMPT_LAYOUT_SELECTION = """You choose the best PowerPoint master slide layout for ONE content slide.

You are given:
- The slide content in markdown
- A compact catalog of available `prs.slide_layouts` from the user's template deck

Your job:
- Pick the best matching layout index for this slide
- Return structured output only (see response model)

SELECTION HEURISTICS:
- Prefer layouts whose name clearly matches the intent (e.g., \"Agenda\", \"Section Header\", \"Quote\", \"Comparison\").
- Use placeholder TYPE counts as a weak signal: e.g. many BODY placeholders often indicates agenda/list layouts.
- If the content is table-like or chart-like, prefer layouts that likely have BODY/OBJECT areas over title-only layouts.

HARD RULES:
- `selected_layout_index` MUST be one of the allowed indices listed in the user message.
- Do not output any markdown or extra fields; return only the structured object."""

# Bumped when `SYSTEM_PROMPT_LAYOUT_DESCRIBE` meaningfully changes (invalidates on-disk cache).
LAYOUT_DESC_CACHE_PROMPT_VERSION = 1

SYSTEM_PROMPT_LAYOUT_DESCRIBE = """You describe ONE PowerPoint master slide layout for a downstream agent that will pick a layout per slide.

Use the layout name, placeholder types, placeholder names (often human-readable), any prompt/placeholder text from the template (e.g. time ranges or \"Section header here\"), and rough bbox/geometry to infer the slide's *intent* (what it is for), not just placeholder counts.
Do not equate a layout named \"Content\" with generic body text if prompt text and structure indicate a table of contents, agenda, or section list.

Return only structured output (see response model). One English sentence, no markdown, at most 160 characters."""


INCREMENTAL_H2_VALIDATION_FIX_TEMPLATE = """The slide script ran but incremental validation failed.

{shared_block}

VALIDATION ISSUES:
{issues}

CONTEXT:
- Deck title (H1): {deck_title}
- Slide section markdown:
{section_markdown}

ORIGINAL CODE:
```python
{original_code}
```

INSTRUCTIONS:
1. Fix the code so it still opens ``TARGET_PPTX``, appends exactly one slide, saves to ``TARGET_PPTX``, and uses ``shared`` for styling.
2. Do not write ``SHARED_PY_PATH`` or change shared on disk.
3. Use attribute names from the shared symbol index above exactly when referencing ``shared``.
4. Call ``get_code_snippet`` if you need canonical python-pptx patterns.
5. Return the complete corrected script (full file)."""


def format_incremental_h1_user_message(
    preamble_markdown: str,
    deck_title: str,
    style_content: str | None,
    language: str | None,
    deck_from_template: bool = False,
) -> str:
    """User message for H1: preamble markdown + deck title."""
    parts = [
        "Generate the H1 (title deck) step for an incremental build.",
        "",
        "Return a single Python program in `code` that imports python-pptx, writes `shared.py` via SHARED_PY_PATH.write_text, then opens and saves the deck at DECK_PPTX_PATH.",
        "",
    ]
    if deck_from_template:
        parts.extend(
            [
                "DECK SOURCE: The file at DECK_PPTX_PATH is a **copy of the user's template** `.pptx`. "
                "It may already contain **multiple** slides (e.g. branding). Implement the H1 title block "
                "by **editing** the most appropriate existing slide(s) — usually the first title slide — "
                "without deleting template slides unless the task explicitly requires that.",
                "",
            ]
        )
    parts.extend(
        [
            f"Presentation title (from H1): {deck_title}",
            "",
            "## MARKDOWN FOR H1 (everything before the first `##`, including the `#` line):",
            preamble_markdown,
        ]
    )
    msg = "\n".join(parts)
    return _append_chunk_style_language(
        msg,
        style_content,
        language,
        style_apply_line="Apply these consistently; encode reusable styling in shared.py.",
    )


def format_incremental_h2_user_message(
    section_markdown: str,
    deck_title: str,
    section_ordinal: int,
    num_h2_slides: int,
    style_content: str | None,
    language: str | None,
    shared_index: str = "",
    *,
    chosen_layout: object | None = None,
) -> str:
    """User message for one H2 slide: single section body."""
    parts = [
        "Generate one slide-append script for an incremental build.",
        "",
        f"Deck title (context): {deck_title}",
        f"Content slide {section_ordinal} of {num_h2_slides} (document order).",
        "",
    ]
    if chosen_layout is not None:
        # Avoid importing LayoutInfo here to keep prompts.py lightweight; the runner passes
        # an object from layout_catalog that has `.index`, `.name`, and a formatter is
        # applied before passing to the model. We embed a pre-formatted string block.
        # The CodeGenerator will pass `chosen_layout_card_full` as a string via this param
        # (see code_generator wiring).
        parts.extend(
            [
                "## SELECTED LAYOUT (already chosen by orchestrator — DO NOT pick another):",
                str(chosen_layout).rstrip(),
                "",
            ]
        )
    if shared_index.strip():
        parts.extend(
            [
                "## CURRENT shared.py SYMBOL INDEX (source of truth for ``shared`` names):",
                "```text",
                shared_index.strip(),
                "```",
                "",
            ]
        )
    parts.extend(
        [
            "## SECTION MARKDOWN (one `##` slide):",
            section_markdown,
        ]
    )
    msg = "\n".join(parts)
    return _append_chunk_style_language(
        msg,
        style_content,
        language,
        style_apply_line="Match styling via the ``shared`` module; do not duplicate large palettes inline.",
    )


def format_layout_describe_user_message(*, layout: LayoutInfo) -> str:
    """Static instructions first (prompt-cache friendly), layout card last."""

    static = (
        "Describe ONE master slide layout for a downstream layout-selection agent.\n"
        "Focus on INTENT (what kind of slide it is for) — distinguish it from layouts "
        "with similar placeholder type counts.\n"
        "Return one sentence (<=160 chars), no markdown.\n"
        "\n"
        "## LAYOUT:\n"
    )
    return static + format_layout_card_full(layout)


def format_layout_selection_user_message(
    *,
    slide_markdown: str,
    deck_title: str,
    layouts: list[LayoutInfo],
    allowed_indices: list[int],
    style_content: str | None,
    language: str | None,
    descriptions: dict[int, str] | None = None,
) -> str:
    catalog = format_layouts_catalog_with_descriptions(layouts, descriptions)
    parts = [
        f'Choose a layout for one H2 content slide of deck "{deck_title}".',
        "",
        "## SLIDE CONTENT (markdown):",
        slide_markdown,
        "",
        "## AVAILABLE LAYOUTS (compact catalog):",
        catalog,
        "",
        "Allowed selected_layout_index values: " + ", ".join(str(i) for i in allowed_indices),
    ]
    msg = "\n".join(parts)
    return _append_chunk_style_language(
        msg,
        style_content,
        language,
        style_apply_line="Use these as weak signals when choosing a branded layout.",
    )


def format_incremental_h2_validation_fix_prompt(
    original_code: str,
    issues: list[str],
    deck_title: str,
    section_markdown: str,
    *,
    shared_index: str = "",
) -> str:
    """Validation fix prompt for a single H2 slide script."""
    if shared_index.strip():
        shared_block = (
            "## CURRENT shared.py SYMBOL INDEX (source of truth for ``shared`` names):\n"
            "```text\n"
            f"{shared_index.strip()}\n"
            "```\n"
        )
    else:
        shared_block = ""
    return INCREMENTAL_H2_VALIDATION_FIX_TEMPLATE.format(
        shared_block=shared_block,
        original_code=original_code,
        issues="\n".join(f"- {i}" for i in issues) if issues else "- (none)",
        deck_title=deck_title,
        section_markdown=section_markdown,
    )


INCREMENTAL_H1_VALIDATION_FIX_TEMPLATE = """The H1 deck script ran but validation failed.

ORIGINAL CODE:
```python
{original_code}
```

ISSUES:
{issues}

H1 MARKDOWN (preamble only):
{preamble_markdown}

INSTRUCTIONS:
1. Keep using ``DECK_PPTX_PATH`` and ``SHARED_PY_PATH`` from the injected header; do not remove the orchestrator prefix.
2. Ensure ``SHARED_PY_PATH`` exists and contains importable Python helpers used for styling.
3. Ensure the deck saved at ``DECK_PPTX_PATH`` reflects the H1 title block and opens correctly.
4. Return the complete corrected script."""


def format_incremental_h1_validation_fix_prompt(
    original_code: str,
    issues: list[str],
    preamble_markdown: str,
) -> str:
    return INCREMENTAL_H1_VALIDATION_FIX_TEMPLATE.format(
        original_code=original_code,
        issues="\n".join(f"- {i}" for i in issues) if issues else "- (none)",
        preamble_markdown=preamble_markdown,
    )
