"""Presentation validation logic."""

import logging
import re
from pathlib import Path
from pptx import Presentation
from .models import ValidationResult

logger = logging.getLogger(__name__)


def parse_expected_slide_count(task_content: str) -> int:
    """
    Parse expected slide count from task content.

    Looks for patterns like "## Slide N" to count expected slides.

    Args:
        task_content: Text content of the task

    Returns:
        Expected number of slides
    """
    # Find all "## Slide N" patterns
    slide_markers = re.findall(r'##\s+Slide\s+(\d+)', task_content, re.IGNORECASE)

    if slide_markers:
        # Get the highest slide number
        max_slide = max(int(num) for num in slide_markers)
        logger.debug(f"Found {len(slide_markers)} slide markers, max: {max_slide}")
        return max_slide

    # Fallback: count lines that look like slide headers
    slide_headers = re.findall(r'##\s+[^#\n]+', task_content)
    logger.debug(f"Fallback: found {len(slide_headers)} potential slide headers")
    return len(slide_headers)


def validate_presentation(
    pptx_path: Path,
    task_content: str
) -> ValidationResult:
    """
    Validate a generated presentation.

    Checks:
    - File exists and is not empty
    - Can be opened as a valid .pptx
    - Slide count matches expected
    - Slides have titles

    Args:
        pptx_path: Path to .pptx file
        task_content: Original task specification

    Returns:
        ValidationResult with validation details
    """
    logger.info(f"Validating presentation: {pptx_path}")

    issues = []

    # Check file exists
    if not pptx_path.exists():
        logger.error(f"File does not exist: {pptx_path}")
        return ValidationResult(
            valid=False,
            slide_count=0,
            has_titles=False,
            issues=["File does not exist"]
        )

    # Check file is not empty
    if pptx_path.stat().st_size == 0:
        logger.error(f"File is empty: {pptx_path}")
        return ValidationResult(
            valid=False,
            slide_count=0,
            has_titles=False,
            issues=["File is empty"]
        )

    try:
        # Open presentation
        prs = Presentation(str(pptx_path))
        slide_count = len(prs.slides)
        logger.info(f"Presentation has {slide_count} slides")

        # Parse expected slide count
        expected_slide_count = parse_expected_slide_count(task_content)
        logger.info(f"Expected {expected_slide_count} slides")

        # Check slide count
        if slide_count != expected_slide_count:
            issues.append(
                f"Slide count mismatch: expected {expected_slide_count}, got {slide_count}"
            )

        # Check if slides have titles
        slides_with_titles = 0
        for i, slide in enumerate(prs.slides, 1):
            if slide.shapes.title:
                title_text = slide.shapes.title.text.strip()
                if title_text:
                    slides_with_titles += 1
                else:
                    issues.append(f"Slide {i} has empty title")
            else:
                issues.append(f"Slide {i} has no title shape")

        has_titles = slides_with_titles == slide_count

        # Determine if valid
        valid = len(issues) == 0

        logger.info(f"Validation result: valid={valid}, issues={len(issues)}")

        return ValidationResult(
            valid=valid,
            slide_count=slide_count,
            expected_slide_count=expected_slide_count,
            has_titles=has_titles,
            issues=issues
        )

    except Exception as e:
        logger.error(f"Error validating presentation: {e}")
        return ValidationResult(
            valid=False,
            slide_count=0,
            has_titles=False,
            issues=[f"Error opening presentation: {str(e)}"]
        )
