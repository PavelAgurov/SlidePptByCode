# PowerPoint Presentation Generator

An LLM-powered agent that generates PowerPoint presentations from text specifications using Python and python-pptx.

## Overview

This tool takes a text/markdown description of a presentation and uses a configurable chat model via **OpenRouter** (default in config: `gpt-4.1-mini`) to:
1. Generate Python code using python-pptx library
2. Execute the code to create a .pptx file
3. Validate the output (slide count, titles)
4. Automatically fix errors through an iterative retry loop

## Features

- **Automated Code Generation**: LLM generates complete python-pptx code from natural language descriptions
- **Chart & Graph Support**: Create native PowerPoint charts (column, line, pie, etc.) from CSV data with brand colors
- **Style Guidelines Support**: Apply custom styling rules (colors, fonts, tone) via `--style` parameter
- **Smart Error Recovery**: Automatically detects and fixes execution errors through retry mechanism
- **Validation**: Verifies generated presentations match specifications (slide count, titles, content)
- **Structured Output**: Uses Pydantic models for reliable LLM responses
- **Safe Execution**: Runs generated code in isolated subprocess with timeout
- **Comprehensive Logging**: Detailed logs in `.logs/` directory

## Installation

### Prerequisites

- Python 3.12+
- OpenRouter API key ([Get one here](https://openrouter.ai/keys))

### Setup

1. Clone the repository:
```bash
git clone <repository-url>
cd SlidePptCode
```

2. Create a virtual environment (one-time; needs any Python 3.12+ on your PATH):
```bash
python -m venv .venv
```

3. Install dependencies (always use the venv interpreter):
```powershell
# Windows (PowerShell)
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```
```bash
# macOS / Linux
./.venv/bin/python -m pip install -r requirements.txt
```

4. Configure environment:
```bash
cp .env.example .env
# Edit .env and add your OPENROUTER_API_KEY
```

## Usage

### Basic Usage

Generate a presentation from a task file. Example markdown and styles live under **`data_sample/`** (for example `data_sample/data01.md`). Unit tests still use `tests/fixtures/sample_task.md`; your own decks can live anywhere (e.g. a local `data/` folder).

**Windows (PowerShell):**
```powershell
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md
```

**macOS / Linux:**
```bash
./.venv/bin/python src/main.py data_sample/data01.md
```

Unless you pass `--output`, the deck is written to **`.output/presentation.pptx`**.

### Advanced Options

```powershell
# Windows — same pattern: .\.venv\Scripts\python.exe src\main.py <TASK> [flags]

# Enable verbose logging
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --verbose

# Output file name (under .output/ if you pass a bare name)
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --output my_deck.pptx

# Language hint for slide content (passed into the LLM prompt)
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --lang Russian

# Apply style guidelines (repo sample)
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --style data_sample\style.md

# Use different model
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --model "gpt-4.1"

# Custom retry limit
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --max-retries 5

# Custom execution timeout
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --timeout 120

# Limit how many H2 (##) content slides to generate (preamble/H1 always runs)
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --slide_max 3

# Use a branded .pptx as the layout source (deck is copied to the output path,
# all slides removed; masters/layouts kept; the original template file is not modified)
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --template path\to\brand.pptx

# Force a specific layout name (works with or without --template; without
# --template, names come from python-pptx defaults like "Title Slide",
# "Title and Content", "Section Header", ...)
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md `
    --template_layout_h1 "Title Slide" `
    --template_layout_h2 "Title and Content"

# Combine multiple options
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --style data_sample\style-epam.md --model "gpt-4.1" --verbose
```

On **Windows PowerShell**, unquoted model names that contain `-` (for example `gpt-4.1`) can be parsed as expressions instead of a single argument. Use quotes as above, or a single token such as `--model=gpt-4.1`.

#### Layout selection flags

- `--template FILE` — copy a `.pptx` and use its master layouts. The runner strips any pre-existing slides from the copy, so the deck always starts with **0 slides** and only your master layouts are kept.
- `--template_layout_h1 NAME` — force a layout by name for the **H1/preamble slide** (skips LLM layout selection for it). Works with or without `--template`. Without `--template`, the name must match a python-pptx default layout (e.g. `Title Slide`, `Title Only`).
- `--template_layout_h2 NAME` — force a layout by name for **every H2 content slide**. Same rules as above.
- `--template_layout NAME` — *deprecated* alias for `--template_layout_h2`. Cannot be combined with the new `_h1`/`_h2` flags.
- `--slide_max N` — limit the number of H2 slides generated. The H1/preamble slide is always produced and is **not** counted; partial runs skip full-deck validation against the task spec.

### Style Guidelines

The `--style` parameter allows you to specify visual and content style guidelines for the presentation. This is useful for maintaining brand consistency or applying specific formatting rules.

**Example style file** (repo sample: `data_sample/style.md`):
```markdown
Use string MS Office styles
Slides should be readable.
Use White and Blue colors as base colors.
Use plain English for text.
```

Style guidelines can include:
- **Visual Style**: Colors, fonts, layouts
- **Tone and Language**: Professional, casual, technical level
- **Data Presentation**: How to format numbers, tables, charts
- **Emphasis**: What to highlight, how to structure content

The style content is injected into the LLM prompt along with the task specification, allowing the AI to generate code that follows your styling preferences.

### Input Format

Create a markdown file with slide specifications:

```markdown
# Presentation Title

## Slide 1 — Introduction

**Title:** Welcome
**Subtitle:** Getting Started

Content for the first slide...

## Slide 2 — Main Points

**Title:** Key Features

- Feature 1
- Feature 2
- Feature 3

## Slide 3 — Conclusion

**Title:** Summary

Final remarks...
```

**Required structure for task files:** the CLI always runs **chunked** generation (one LLM pass for the title block — everything from the first `#` through the line before the first `##` — then one pass per `##` slide). Your markdown **must** include:

- An **H1** line (`# Presentation title`) before any **H2** slide heading.
- At least one **H2** line (`## Slide title`) for content slides.

If this structure is missing or the first `##` appears before the first `#`, the program exits with a clear `ValueError` (no single-pass fallback). Chunking is always on; a previous `--chunk-by-h2` CLI flag is removed.

The tool will:
- Parse each top-level `## ...` heading as one content slide and count H2 sections for validation
- Extract titles and content for each slide
- Generate appropriate python-pptx code
- Create the presentation in `.output/` directory

### Creating Charts and Graphs

The tool supports creating native PowerPoint charts from CSV data. Charts are fully editable in PowerPoint after generation.

**Specifying Chart Type:**

You can specify chart type in two ways:

1. **Chart directive** (recommended): Add `chart: <type>` before or after the CSV data
2. **Natural language**: Describe the chart in text (e.g., "Display as a line chart")

**Supported chart types:**
- `chart: column` - Clustered column chart (default for multi-series)
- `chart: column_stacked` - Stacked column chart
- `chart: bar` - Horizontal bar chart
- `chart: line` - Line chart
- `chart: line_markers` - Line chart with markers
- `chart: pie` - Pie chart (default for proportions)
- `chart: doughnut` - Doughnut chart
- `chart: area` - Area chart
- `chart: scatter` - Scatter plot

**Example with charts:** the repo includes **`data_sample/data03_charts_test.md`** (three slides with `chart:` directives and CSV). Open that file for the full markdown.

Minimal shape (see `data_sample/data03_charts_test.md` for complete CSV blocks):

```markdown
# Sales Dashboard

## Slide 1 — Revenue Trends

chart: line_markers

month,product_A,product_B,product_C
2026-01,120000,95000,60000
```

Generate (optional EPAM-oriented style sample):
```powershell
.\.venv\Scripts\python.exe src\main.py data_sample\data03_charts_test.md --style data_sample\style-epam.md
```

When using style guidelines, chart colors automatically match your brand colors for consistency.

## Project Structure

```
SlidePptCode/
├── src/
│   ├── main.py                 # CLI entry point
│   ├── incremental_runner.py   # Slide-by-slide pipeline orchestration
│   ├── config.py               # Settings (pydantic-settings)
│   ├── models.py               # Pydantic models for LLM structured output
│   ├── llm_client.py           # OpenRouter-compatible client
│   ├── code_generator.py       # Prompting, codegen, retries
│   ├── code_executor.py        # Subprocess execution with timeout
│   ├── validator.py            # Deck vs. task validation
│   ├── prompts.py              # System prompts and templates
│   ├── task_chunker.py         # Markdown → H1 / H2 chunks
│   ├── ppt_bootstrap.py        # Empty deck or template copy + strip slides
│   ├── layout_catalog.py       # Read layouts from a .pptx
│   ├── layout_describer.py     # LLM one-line descriptions per layout
│   ├── shared_default.py       # Default shared.py source (no --style)
│   ├── shared_index.py         # Helpers for shared module handling
│   ├── script_inject.py        # Inject paths into generated slide scripts
│   ├── snippets.py             # Small reusable prompt/code fragments
│   ├── code_merger.py          # Legacy / helper merge utilities
│   └── colored_logging.py      # Console log colors
├── tests/
│   ├── test_*.py               # Unit tests
│   └── fixtures/               # sample_task.md (used by tests, not the main demos)
├── data_sample/                # Example tasks + styles (safe to run with the CLI)
│   ├── data01.md               # Longer narrative / financial-style deck
│   ├── data02.md               # Short example with CSV table
│   ├── data03_charts_test.md   # chart: directives + charts
│   ├── style.md, style2.md     # Generic style guideline samples
│   └── style-epam.md           # Brand-oriented style sample
├── .generated/                 # Generated slide scripts + shared.py (auto-created)
├── .output/                    # Generated .pptx (+ _scratch/ during runs)
├── .logs/                      # agent.log (auto-created)
├── .env                        # Your secrets (not committed)
├── .env.example                # OPENROUTER_API_KEY placeholder
└── requirements.txt            # Python dependencies
```

Optional: add your own `data/`, `styles/`, or `templates/` folders; the shipped **`data_sample/`** tree is enough to try the tool without creating paths.

## Configuration

Edit `.env` file or set environment variables:

```bash
# Required
OPENROUTER_API_KEY=your_api_key_here

# Optional (defaults shown)
MODEL_ID=gpt-4.1-mini
BASE_URL=https://openrouter.ai/api/v1
TEMPERATURE=0
MAX_RETRIES=50
EXECUTION_TIMEOUT=60
```

### Model Options

- `gpt-4.1-mini` - Faster and cheaper (default)
- `gpt-4.1` - Better code quality, higher cost

## Testing

### Run Tests

Use the project virtual environment (see workspace rules):

```powershell
# Windows — all tests
.\.venv\Scripts\pytest.exe

# One file
.\.venv\Scripts\pytest.exe tests\test_validator.py -v

# With coverage
.\.venv\Scripts\pytest.exe tests -v --cov=src --cov-report=html
```

```bash
# macOS / Linux
./.venv/bin/pytest
./.venv/bin/pytest tests/test_validator.py -v
```

## How It Works

### Pipeline Overview (incremental, slide-by-slide)

The CLI runs a **unified incremental pipeline**: every slide — including the H1 / preamble — is appended to the deck by its own LLM-generated script with the exact same contract.

1. **Load configuration** from `.env`.
2. **Parse and chunk the task**: H1 preamble + one chunk per `##` heading.
3. **Initialize the deck with 0 slides**: either save a fresh empty `.pptx` (default python-pptx layouts) or copy `--template` and strip all of its existing slides. Only master layouts survive.
4. **Read layouts** from the deck and (optionally) cache one-line LLM descriptions for each layout (cached on disk under `.generated/layout/`). Default python-pptx layouts are cached under `default-pptx.json`.
5. **Generate `shared.py`** once for the whole run:
   - With no `--style`: write a hardcoded default module verbatim (no LLM call).
   - With `--style`: ask the LLM to **extend** the default with brand-specific palette constants while keeping every default symbol.
   - Then verify the resulting file with `ast.parse` + `importlib.exec_module` and check that all required default symbols are still present. On failure, retry with `fix_shared_module` up to `--max-retries`. Failure aborts the run before any slide is generated.
6. **Slide loop** (preamble first, then each H2 chunk in order). For every chunk:
   1. Pick a layout — either the forced `--template_layout_h1` / `--template_layout_h2`, or LLM layout selection (with two attempts; failure on the preamble aborts the run because there is no safe default fallback for title-only layouts).
   2. Generate a per-slide append script with the orchestrator-injected `TARGET_PPTX`, `SHARED_PY_PATH`, and `CHOSEN_LAYOUT_INDEX`.
   3. Run on a **scratch copy** of the deck and validate that exactly one slide was appended and has readable content.
   4. Run on the real deck, validate, and fall back to a `.bak` copy on failure.
   5. Confirm `shared.py` was not modified by the slide script.
   6. Retry on any failure up to `--max-retries` with full traceback / validation issues fed back to the LLM.
7. **Final validation** of the deck against the task spec (skipped when `--slide_max` produces a partial run).
8. **Save results**: generated scripts in `.generated/`, the deck in `.output/`, full agent log in `.logs/agent.log`.

### Error Recovery

The agent implements a smart retry loop:

- **Execution Errors**: Captures traceback and stderr, sends to LLM for fixing
- **Validation Errors**: Detects slide count mismatches or missing titles, requests corrections
- **Timeout Handling**: Terminates long-running code, generates faster alternatives
- **Conversation History**: Maintains context across retries for better fixes

### Structured Output

Uses Pydantic models with the OpenAI-compatible **structured outputs** / JSON-schema style responses from the API. The incremental pipeline mainly uses:

- **`IncrementalLlmScriptCode`** — one Python file per H1 or H2 slide append
- **`SharedModuleCode`** — optional brand extensions for `.generated/shared.py` when `--style` is set
- **`LayoutDescription`** / **`LayoutSelection`** — layout catalog text and per-slide layout choice

Legacy / other flows may still use **`GeneratedCode`** (full script with `code`, `explanation`, `expected_output_filename`).

## Examples

### Example 1: Simple Presentation

Task files must include an **H1** line before any **H2** slide (see [Input Format](#input-format)). Example:

```markdown
# My Presentation

## Slide 1 — Title
**Title:** Welcome

## Slide 2 — Content
**Title:** Main Point
- Item 1
- Item 2
```

Default output: **`.output/presentation.pptx`** (three slides: H1 preamble plus two H2 slides). Use `--slide_max 1` if you only want the preamble plus one H2.

### Example 2: Shipped samples in `data_sample/`

| File | Purpose |
|------|---------|
| `data_sample/data01.md` | Larger deck (multiple H2 sections, narrative + metrics) |
| `data_sample/data02.md` | Short task with a CSV block |
| `data_sample/data03_charts_test.md` | Charts via `chart:` + CSV |
| `data_sample/style.md`, `style2.md`, `style-epam.md` | Use with `--style …` |

For automated tests only, the minimal markdown fixture is `tests/fixtures/sample_task.md`.

## Troubleshooting

### Common Issues

**API Key Error:**
```
Error: OPENROUTER_API_KEY not set
```
Solution: Add your API key to `.env` file

**Timeout Errors:**
```
Execution timed out after 60 seconds
```
Solution: Increase timeout with `--timeout 120`

**Validation Failures:**
```
Slide count mismatch: expected 5, got 4
```
Solution: Check input file slide markers or increase `--max-retries`

**Import Errors:**
```
ModuleNotFoundError: No module named 'pptx'
```
Solution: Install dependencies with the venv interpreter (`python -m pip install -r requirements.txt` inside the venv, or `.\.venv\Scripts\python.exe -m pip install -r requirements.txt` on Windows).

### Debug Mode

Enable verbose logging to see detailed execution flow:

```powershell
.\.venv\Scripts\python.exe src\main.py data_sample\data01.md --verbose
```

Check logs in `.logs/agent.log` for full details.

## Development

### Adding New Features

1. Update models in `src/models.py`
2. Modify prompts in `src/prompts.py`
3. Add logic in respective modules
4. Write tests in `tests/`
5. Run full test suite

### Code Style

- Type hints for all functions
- Docstrings for classes and public methods
- Logging for important operations
- Error handling with clear messages

## License

MIT License

Copyright (c) 2026 SlidePptCode contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Add tests for new features
4. Submit a pull request

## Acknowledgments

- Built with [python-pptx](https://python-pptx.readthedocs.io/)
- Powered by [OpenRouter](https://openrouter.ai/)
- Uses OpenAI's structured output feature
