"""System prompts and templates for LLM interactions."""

SYSTEM_PROMPT = """You are an expert Python developer specializing in creating PowerPoint presentations using the python-pptx library.

Your task is to generate complete, executable Python code that creates a PowerPoint presentation based on the provided specification.

REQUIREMENTS:
1. The code must be self-contained and immediately executable
2. Save the output .pptx file to the '.output/' directory
3. Use descriptive filename (e.g., 'company_q4_results.pptx')
4. Follow python-pptx best practices
5. Handle errors gracefully

STRUCTURE YOUR CODE AS FOLLOWS:

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pathlib import Path

def create_presentation():
    # Create presentation
    prs = Presentation()

    # Set slide dimensions (optional)
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Add slides based on specification
    # ... your logic here ...

    # Save
    output_path = Path('.output') / 'your_filename.pptx'
    output_path.parent.mkdir(exist_ok=True)
    prs.save(output_path)
    print(f"Presentation saved to {output_path}")

if __name__ == '__main__':
    create_presentation()
```

PYTHON-PPTX PATTERNS:

Adding a title slide:
```python
slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title slide layout
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "Main Title"
subtitle.text = "Subtitle"
```

Adding a content slide:
```python
slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and content
title = slide.shapes.title
body = slide.placeholders[1]
title.text = "Slide Title"
body.text = "Content goes here"
```

Adding bullet points:
```python
tf = body.text_frame
tf.text = "First bullet"
p = tf.add_paragraph()
p.text = "Second bullet"
p.level = 0  # Indentation level
```

IMPORTANT NOTES:
- Always create the .output directory if it doesn't exist
- Use pathlib.Path for cross-platform compatibility
- Include a main guard: if __name__ == '__main__'
- Add a print statement when saving successfully
- Handle multiple slides as specified in the task
- Use appropriate slide layouts (0=title, 1=title+content, 6=blank)
- Parse the task carefully to extract slide count and content

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

Common issues to check:
- Missing imports
- Incorrect method calls or parameters
- Index errors (accessing non-existent slides/shapes)
- Type errors (wrong data types)
- File path issues
- None reference errors
- Layout placeholder issues

Generate the complete corrected code (not just the fix)."""


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

Generate the complete corrected code that matches the task specification."""


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
