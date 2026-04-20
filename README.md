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
- **Style Guidelines Support**: Apply custom styling rules (colors, fonts, tone) via `--style` parameter
- **Smart Error Recovery**: Automatically detects and fixes execution errors through retry mechanism
- **Validation**: Verifies generated presentations match specifications (slide count, titles)
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
python src/main.py data/data01.md --style data/style.md

# Use different model
python src/main.py data/data01.md --model gpt-4.1

# Custom retry limit
python src/main.py data/data01.md --max-retries 5

# Custom execution timeout
python src/main.py data/data01.md --timeout 120

# Combine multiple options
python src/main.py data/data01.md --style data/style.md --model gpt-4.1 --verbose
```

### Style Guidelines

The `--style` parameter allows you to specify visual and content style guidelines for the presentation. This is useful for maintaining brand consistency or applying specific formatting rules.

**Example style file** (`data/style.md`):
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

The tool will:
- Parse `## Slide N` markers to determine expected slide count
- Extract titles and content for each slide
- Generate appropriate python-pptx code
- Create the presentation in `.output/` directory

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
│   ├── test_integration.py  # Integration tests with real LLM
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

### Run Unit Tests

```bash
# All unit tests
python -m pytest tests/ -v --ignore=tests/test_integration.py

# Specific test file
python -m pytest tests/test_validator.py -v

# With coverage
python -m pytest tests/ -v --cov=src --cov-report=html
```

### Run Integration Tests

Integration tests make real API calls to OpenRouter:

```bash
python -m pytest tests/test_integration.py -v -m integration
```

## How It Works

### Pipeline Overview

1. **Load Configuration**: Read settings from `.env`
2. **Parse Task**: Read input file and extract presentation requirements
3. **Generate Code**: LLM creates python-pptx code based on specifications
4. **Execute Code**: Run generated code in isolated subprocess
5. **Validate Output**: Check slide count, titles, file existence
6. **Retry on Error**: If execution or validation fails, send error details back to LLM for fixing
7. **Save Results**: Store generated code in `.generated/`, presentation in `.output/`

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
