"""Tests for Pydantic models."""

import pytest
from pydantic import ValidationError
from src.models import (
    GeneratedCode,
    IncrementalLlmScriptCode,
    CodeExecutionResult,
    ValidationResult,
)


def test_generated_code_valid():
    """Test valid GeneratedCode model."""
    code = GeneratedCode(
        code="from pptx import Presentation\nprs = Presentation()",
        explanation="Creates a presentation",
        expected_output_filename="test.pptx"
    )

    assert "from pptx import" in code.code
    assert code.explanation == "Creates a presentation"
    assert code.expected_output_filename == "test.pptx"


def test_incremental_llm_script_code_allows_no_pptx_import_in_model():
    """Incremental parse model does not enforce pptx import (runtime still must)."""
    m = IncrementalLlmScriptCode(
        code="print('only shared snippet')",
        explanation="x",
    )
    assert "pptx" not in m.code


def test_generated_code_missing_import():
    """Test GeneratedCode validation fails without python-pptx import."""
    with pytest.raises(ValidationError) as exc_info:
        GeneratedCode(
            code="print('hello')",
            explanation="No pptx import",
            expected_output_filename="test.pptx"
        )

    assert "must import python-pptx" in str(exc_info.value)


def test_code_execution_result_success():
    """Test successful CodeExecutionResult."""
    from pathlib import Path

    result = CodeExecutionResult(
        success=True,
        output_file=Path("test.pptx"),
        stdout="Success"
    )

    assert result.success is True
    assert result.output_file == Path("test.pptx")
    assert result.error_message is None


def test_code_execution_result_failure():
    """Test failed CodeExecutionResult."""
    result = CodeExecutionResult(
        success=False,
        error_message="File not found",
        stderr="Error: file missing",
        traceback="Traceback..."
    )

    assert result.success is False
    assert result.error_message == "File not found"
    assert result.output_file is None


def test_validation_result_valid():
    """Test valid ValidationResult."""
    result = ValidationResult(
        valid=True,
        slide_count=8,
        expected_slide_count=8,
        has_titles=True,
        issues=[]
    )

    assert result.is_valid is True
    assert result.slide_count == 8


def test_validation_result_invalid():
    """Test invalid ValidationResult."""
    result = ValidationResult(
        valid=False,
        slide_count=5,
        expected_slide_count=8,
        has_titles=True,
        issues=["Slide count mismatch"]
    )

    assert result.is_valid is False
    assert len(result.issues) == 1
