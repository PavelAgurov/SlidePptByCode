"""System prompts and templates for LLM interactions."""

from .code_merger import CHUNK_MERGE_MARKER
from .layout_catalog import LayoutInfo, format_layout_card_full, format_layouts_catalog_with_descriptions
from .models import LayoutDescription
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
# Incremental pipeline (shared.py once, then unified per-slide append)
# ---------------------------------------------------------------------------

SYSTEM_PROMPT_SHARED_EXTEND = """You extend a small Python helper module ``shared.py`` used by per-slide scripts.

CONTRACT:
- Output a complete Python module in the ``code`` field. No markdown fences.
- You are given a DEFAULT module body and STYLE GUIDELINES.
- Keep **every** symbol from the default (functions and constants) with the same names and the same call signatures. You may extend the body of a function or its docstring, but the public name must stay; do not rename or delete anything.
- Add named ``RGBColor`` palette constants for colors mentioned in the style guidelines, using descriptive uppercase names (e.g. ``ELF_GREEN``, ``DEEP_SAFFRON``). Use the existing ``rgb("#RRGGBB")`` helper.
- Allowed imports: only ``from pptx.dml.color import RGBColor``. Do not import anything else, especially not ``pptx.Presentation``, ``pathlib``, etc.
- The module must be importable as-is (no top-level side effects beyond constant definitions, no I/O, no print).
- No deck logic: shared.py is consumed by other scripts that build slides."""


SYSTEM_PROMPT_INCREMENTAL_SLIDE = """You are an expert Python developer using python-pptx.

STRUCTURED OUTPUT (critical):
- The ``code`` field is **one** runnable ``.py`` file. It **must** include ``from pptx import Presentation`` or ``import pptx`` (in addition to any other imports).

USER MESSAGE STRUCTURE (critical — read carefully):
- The user message uses XML-style tags. Treat each tag as an isolated container.
- ``<slide_markdown>`` is the **only** source of slide content. Render exactly what is inside it.
- ``<chosen_layout>``, ``<shared_symbols>``, ``<deck_context>``, ``<style_note>``, and ``<language_requirement>`` are **instructions/metadata**, NOT slide content. Never put their text on the slide.
- If a brand style applies, the palette is already encoded as constants on ``shared.*`` (see ``<shared_symbols>``). NEVER paste color tables, hex lists, or palette guides onto the slide as content — reference ``shared.*`` constants in code instead.

PER-SLIDE APPEND CONTRACT (orchestrator injects constants at the top of the file):
- ``TARGET_PPTX`` is the path to the presentation to modify (scratch or final deck). The deck may currently have **0 or more** slides.
- ``shared`` is already loaded from ``SHARED_PY_PATH`` — use ``shared`` for styling/helpers; do NOT write to ``SHARED_PY_PATH`` and do not redefine it.
- Use **only** attribute names that appear in the ``CURRENT shared.py SYMBOL INDEX`` block in the user message (exact spelling). Never invent names like ``shared.elf_green`` if the index shows ``ELF_GREEN``.
- Open with ``prs = Presentation(str(TARGET_PPTX))``.
- Append **exactly one** new slide at the end: ``prs.slides.add_slide(...)``. Do not remove slides. Do not modify shapes/text on slides whose index is less than the slide count before your addition (append-only for existing slides). Appending also works when the deck currently has zero slides — the new slide simply becomes slide 0.
- **Layouts and placeholders:** Different layouts have different placeholder ``idx`` values. Always use the ``CHOSEN_LAYOUT_INDEX`` constant the orchestrator injects. The user message contains a ``SELECTED LAYOUT`` block listing every placeholder ``idx`` and its ``type``; address placeholders strictly by exact ``idx`` from that card. Do not assume numeric indices like ``slide.placeholders[1]`` and do not call ``pick_title_and_content_layout``.
- For title-style layouts (TITLE + SUBTITLE, no BODY), the SUBTITLE placeholder is normal text; set its ``text`` (or build via ``text_frame``).
- If the section markdown contains a **pipe table** (``|`` columns), use ``get_code_snippet`` with ``table_markdown_grid`` and build a native ``Table`` via ``add_table``. Do **not** paste the markdown table into ``body.text`` or a bullet list as plain text.
- Save back to the same file: ``prs.save(str(TARGET_PPTX))``.
- End with ``if __name__ == '__main__':`` calling one entry function.
- Use only the Python standard library plus python-pptx.
- Do not print paths for orchestration; optional logs are fine.

""" + _SNIPPET_TOOL_RULES + """

Call ``get_code_snippet`` with ``incremental_h2_skeleton`` for a minimal append-one-slide pattern."""


# Backwards-compatible alias (legacy name imported elsewhere)
SYSTEM_PROMPT_INCREMENTAL_H2 = SYSTEM_PROMPT_INCREMENTAL_SLIDE


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
- Catalog lines may end with ``zones=N`` (distinct visual content regions from the template metadata). Use as a weak signal (e.g. side-by-side vs single canvas).
- If the content is table-like or chart-like, prefer layouts that likely have BODY/OBJECT areas over title-only layouts.

HARD RULES:
- `selected_layout_index` MUST be one of the allowed indices listed in the user message.
- Do not output any markdown or extra fields; return only the structured object.
- Do not make up answer, use only provided information.
"""

# Bumped when `SYSTEM_PROMPT_LAYOUT_DESCRIBE` meaningfully changes (invalidates on-disk cache).
LAYOUT_DESC_CACHE_PROMPT_VERSION = 5

SYSTEM_PROMPT_LAYOUT_DESCRIBE = """You describe ONE PowerPoint master slide layout for a downstream agent that will pick a layout per slide.

Use the layout name, placeholder types, placeholder names (often human-readable), any prompt/placeholder text from the template (e.g. time ranges or \"Section header here\"), rough bbox/geometry, and the **Decorative shapes** section (non-placeholder rectangles, freeforms, pictures on the layout) to infer the slide's *intent* (what it is for), not just placeholder counts.
Do not equate a layout named \"Content\" with generic body text if prompt text and structure indicate a table of contents, agenda, or section list.

``content`` vs ``header`` (critical):
- ``content`` — the slide is meant to carry **body material**: bullets, paragraphs, charts, tables, images, or freeform shapes. Count it as ``content`` if there is a BODY (or similar) **text/content placeholder**, **or** if there is **no** such placeholder but most of the canvas is **plain empty space** clearly intended for arbitrary content (e.g. only footer/date/slide-number chrome and a large blank area — still ``content``, not ``other``).
- ``header`` — the slide is dominated by a **large centered title** (and maybe subtitle); it is **not** a general-purpose canvas for dense body copy. Decorative branding alone does not make it ``content`` if the layout's role is clearly a title/cover/section opener.

You MUST set ``slide_type`` (structured field) to exactly one of:
- ``header`` — title slide, cover, or section opener: dominant title/subtitle; not a body-content canvas.
- ``agenda`` — table of contents, agenda, outline of upcoming sections.
- ``content`` — body slide: content/text placeholder and/or large usable empty area for content as above.
- ``other`` — does not fit the above (e.g. picture-with-caption only, vertical-title specialty). Use sparingly; do **not** use ``other`` for minimal-placeholder decks that are clearly meant for freeform body slides.

You MUST set ``slide_has_image_placeholder`` (boolean): ``true`` if the layout has a dedicated **PICTURE** / image placeholder (a slot meant for a user-supplied photo or large graphic — check placeholder **type** names in the card, e.g. PICTURE). ``false`` if there is no such placeholder (text/body/subtitle only and/or blank canvas). Small fixed brand marks in the master are **not** a picture placeholder.

You MUST set ``content_zones_count`` (integer >= 1): the number of distinct top-level **visual content regions** on the layout.
- ``1`` — single content area: one BODY/CONTENT (or similar) placeholder, **or** no body placeholder but one large empty canvas intended for arbitrary material. Footer / logo / slide-number chrome do **not** split the canvas into a second zone.
- ``2`` — clear split into two top-level regions: e.g. two BODY placeholders side-by-side, or one BODY + one PICTURE, **or** a large decorative shape under **Decorative shapes** (background rectangle, chevron, freeform panel) that visibly partitions the slide into two major areas (e.g. left black panel + right white panel).
- ``3+`` — three or more such top-level regions (e.g. three-column comparison).
Count **zones**, not bullet items inside one placeholder. Purely decorative accents that do not define a separate main content area do not increase the count.

Return only structured output (see response model). ``description``: one English sentence, no markdown, at most 160 characters.
"""


INCREMENTAL_H2_VALIDATION_FIX_TEMPLATE = """The slide script ran but incremental validation failed.

{shared_block}

<validation_issues>
{issues}
</validation_issues>

<deck_context>
  <deck_title>{deck_title}</deck_title>
</deck_context>

<slide_markdown>
{section_markdown}
</slide_markdown>

<original_code>
```python
{original_code}
```
</original_code>

INSTRUCTIONS:
1. Fix the code so it still opens ``TARGET_PPTX``, appends exactly one slide, saves to ``TARGET_PPTX``, and uses ``shared`` for styling.
2. Do not write ``SHARED_PY_PATH`` or change shared on disk.
3. Use attribute names from the shared symbol index above exactly when referencing ``shared``.
4. Use ONLY the content inside <slide_markdown> as slide text; do NOT paste anything from <validation_issues>, <deck_context>, or any style/palette tables.
5. Call ``get_code_snippet`` if you need canonical python-pptx patterns.
6. Return the complete corrected script (full file)."""


def format_shared_extend_user_message(
    *, default_shared_py: str, style_content: str
) -> str:
    """User message for extending the default shared.py with style guidelines."""
    return (
        "Extend the DEFAULT shared.py module to encode the brand palette and any "
        "reusable styling helpers implied by the STYLE GUIDELINES.\n"
        "\n"
        "Rules:\n"
        "- Keep every name from the default (functions, constants).\n"
        "- Add new RGBColor palette constants for brand colors using uppercase names.\n"
        "- Allowed import: only `from pptx.dml.color import RGBColor`.\n"
        "- Module must be importable as-is; no top-level I/O.\n"
        "\n"
        "## DEFAULT shared.py:\n"
        "```python\n"
        f"{default_shared_py}\n"
        "```\n"
        "\n"
        "## STYLE GUIDELINES:\n"
        f"{style_content}\n"
    )


def format_shared_extend_fix_user_message(
    *, prev_code: str, traceback: str
) -> str:
    """User message for fixing shared.py after verification failed."""
    return (
        "The shared.py module you produced fails verification. Fix it.\n"
        "\n"
        "Rules (unchanged):\n"
        "- Keep every name from the original default module (functions, constants).\n"
        "- Allowed import: only `from pptx.dml.color import RGBColor`.\n"
        "- Module must parse and import cleanly with no top-level I/O.\n"
        "\n"
        "## VERIFICATION ERROR:\n"
        "```text\n"
        f"{traceback}\n"
        "```\n"
        "\n"
        "## PREVIOUS shared.py:\n"
        "```python\n"
        f"{prev_code}\n"
        "```\n"
        "\n"
        "Return the complete corrected module.\n"
    )


def format_incremental_slide_user_message(
    section_markdown: str,
    deck_title: str,
    style_content: str | None,
    language: str | None,
    shared_index: str = "",
    *,
    chosen_layout: object | None = None,
) -> str:
    """
    User message for one slide append (preamble or H2 — uniform contract).

    Inputs are wrapped in XML-style tags so the model never confuses
    instructions, the slide source markdown, the layout card, and the
    shared-symbol index with each other. The raw ``style_content`` is
    intentionally **not** embedded here: brand colors / palette are already
    encoded in ``shared.py`` (see ``<shared_symbols>``); injecting the style
    guide here causes the model to paste palette tables onto the slide.
    """
    parts: list[str] = [
        "You build ONE slide. Use ONLY the markdown inside <slide_markdown> as",
        "slide content. Do NOT include text from any other tag (deck title,",
        "layout card, shared symbols, language directive) in slide shapes.",
        "Brand palette is exposed via the `shared` module (see <shared_symbols>);",
        "use those constants for colors and never paste palette tables onto slides.",
        "",
        "<deck_context>",
        f"  <deck_title>{deck_title}</deck_title>",
        "</deck_context>",
        "",
    ]
    if chosen_layout is not None:
        # The orchestrator pre-formats the layout card; we embed it verbatim.
        layout_block = str(chosen_layout).rstrip()
        parts.extend(
            [
                "<chosen_layout note=\"already chosen by orchestrator — DO NOT pick another\">",
                layout_block,
                "</chosen_layout>",
                "",
            ]
        )
    if shared_index.strip():
        parts.extend(
            [
                "<shared_symbols note=\"exact names available on the `shared` module\">",
                shared_index.strip(),
                "</shared_symbols>",
                "",
            ]
        )
    if style_content and style_content.strip():
        # We DO NOT include the raw style markdown — it is already encoded in
        # shared.py. We only flag that a brand style applies, so the model
        # prefers shared.* over hardcoded colors.
        parts.extend(
            [
                "<style_note>",
                "A brand style is in effect. Its palette is encoded in `shared.*`.",
                "Do not embed the style guide as slide content; reference shared.* constants only.",
                "</style_note>",
                "",
            ]
        )
    if language:
        parts.extend(
            [
                "<language_requirement>",
                f"All visible text on the slide MUST be in {language}.",
                "</language_requirement>",
                "",
            ]
        )
    parts.extend(
        [
            "<slide_markdown>",
            section_markdown.rstrip(),
            "</slide_markdown>",
        ]
    )
    return "\n".join(parts)


# Backwards-compatible alias: the runner used to call this for H2 slides.
def format_incremental_h2_user_message(
    section_markdown: str,
    deck_title: str,
    section_ordinal: int,  # kept for backwards-compat (unused in slide message)
    num_h2_slides: int,    # kept for backwards-compat (unused in slide message)
    style_content: str | None,
    language: str | None,
    shared_index: str = "",
    *,
    chosen_layout: object | None = None,
) -> str:
    """Deprecated alias for :func:`format_incremental_slide_user_message`."""
    del section_ordinal, num_h2_slides
    return format_incremental_slide_user_message(
        section_markdown=section_markdown,
        deck_title=deck_title,
        style_content=style_content,
        language=language,
        shared_index=shared_index,
        chosen_layout=chosen_layout,
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
    descriptions: dict[int, LayoutDescription] | None = None,
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
            "<shared_symbols note=\"exact names available on the `shared` module\">\n"
            f"{shared_index.strip()}\n"
            "</shared_symbols>\n"
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


