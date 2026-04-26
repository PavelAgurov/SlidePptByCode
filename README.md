# PowerPoint Presentation Generator

An LLM-powered agent that generates PowerPoint presentations from text specifications using Python and python-pptx.

## Overview

This tool takes a text/markdown description of a presentation and uses GPT-4.1 (via OpenRouter) to:
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

2. Create and activate virtual environment:
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
# or
source .venv/bin/activate  # Unix/macOS
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Configure environment:
```bash
cp .env.example .env
# Edit .env and add your OPENROUTER_API_KEY
```

## Usage

### Basic Usage

Generate a presentation from a task file:

```bash
python src/main.py data/data01.md
```

### Advanced Options

```bash
# Enable verbose logging
python src/main.py data/data01.md --verbose

# Apply style guidelines
python src/main.py data/data01.md --style styles/style.md

# Use different model
python src/main.py data/data01.md --model "gpt-4.1"

# Custom retry limit
python src/main.py data/data01.md --max-retries 5

# Custom execution timeout
python src/main.py data/data01.md --timeout 120

# Limit how many H2 (##) content slides to generate (preamble/H1 always runs)
python src/main.py data/data01.md --slide_max 3

# Use a branded .pptx as the layout source (existing slides are stripped;
# masters/layouts are kept)
python src/main.py data/data01.md --template templates/brand.pptx

# Force a specific layout name (works with or without --template; without
# --template, names come from python-pptx defaults like "Title Slide",
# "Title and Content", "Section Header", ...)
python src/main.py data/data01.md \
    --template_layout_h1 "Title Slide" \
    --template_layout_h2 "Title and Content"

# Combine multiple options
python src/main.py data/data01.md --style styles/style.md --model "gpt-4.1" --verbose
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

**Example style file** (`styles/style.md`):
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
- Parse `## Slide N` markers to determine expected slide count
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

**Example with charts** (`data/charts_example.md`):

```markdown
# Sales Dashboard

## Slide 1 — Revenue Trends

chart: line_markers

month,product_A,product_B,product_C
2025-01,120000,95000,60000
2025-02,135000,102000,72000
2025-03,150000,110000,80000
2025-04,170000,130000,95000

## Slide 2 — Market Share

chart: pie

product,percentage
Product A,45
Product B,30
Product C,25

## Slide 3 — Quarterly Performance

chart: column_stacked

quarter,revenue,costs
Q1,355000,297000
Q2,515000,395000
Q3,605000,460000
Q4,720000,550000
```

Generate the presentation:
```bash
python src/main.py data/charts_example.md --style styles/style-epam.md
```

When using style guidelines, chart colors automatically match your brand colors for consistency.

## Project Structure

```
SlidePptCode/
├── src/
│   ├── main.py              # CLI entry point and orchestration
│   ├── config.py            # Configuration management (pydantic-settings)
│   ├── models.py            # Pydantic models for structured output
│   ├── llm_client.py        # OpenRouter/OpenAI client wrapper
│   ├── code_generator.py    # Code generation and error fixing logic
│   ├── code_executor.py     # Safe code execution in subprocess
│   ├── validator.py         # Presentation validation
│   └── prompts.py           # System prompts and templates
├── tests/
│   ├── test_*.py            # Unit tests
│   └── fixtures/            # Test data
├── data/                    # Input task specifications
├── .generated/              # Generated Python code (auto-created)
├── .output/                 # Generated .pptx files (auto-created)
├── .logs/                   # Application logs (auto-created)
├── .env                     # Environment configuration
└── requirements.txt         # Python dependencies
```

## Configuration

Edit `.env` file or set environment variables:

```bash
# Required
OPENROUTER_API_KEY=your_api_key_here

# Optional (defaults shown)
MODEL_ID=gpt-4.1-mini
BASE_URL=https://openrouter.ai/api/v1
TEMPERATURE=0
MAX_RETRIES=3
EXECUTION_TIMEOUT=60
```

### Model Options

- `gpt-4.1-mini` - Faster and cheaper (default)
- `gpt-4.1` - Better code quality, higher cost

## Testing

### Run Tests

```bash
# All tests
python -m pytest tests/ -v

# Specific test file
python -m pytest tests/test_validator.py -v

# With coverage
python -m pytest tests/ -v --cov=src --cov-report=html
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

Uses Pydantic models with OpenAI's structured output feature:

```python
class GeneratedCode(BaseModel):
    code: str                      # Complete Python code
    explanation: str                # Approach description
    expected_output_filename: str   # Expected .pptx filename
```

This ensures reliable, parseable responses from the LLM.

## Examples

### Example 1: Simple Presentation

Input (`data/simple.md`):
```markdown
## Slide 1 — Title
**Title:** My Presentation

## Slide 2 — Content
**Title:** Main Point
- Item 1
- Item 2
```

Output: `simple_presentation.pptx` with 2 slides

### Example 2: Complex Presentation

See `data/data01.md` for a full example (8 slides with financial data, tables, formatting).

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
Solution: Ensure virtual environment is activated and dependencies installed

### Debug Mode

Enable verbose logging to see detailed execution flow:

```bash
python src/main.py data/data01.md --verbose
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

[Your License Here]

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
