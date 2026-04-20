"""Tests for code generator."""

import pytest
from unittest.mock import Mock, MagicMock
from src.code_generator import CodeGenerator
from src.models import GeneratedCode, CodeExecutionResult, ValidationResult


@pytest.fixture
def mock_llm_client():
    """Create mock LLM client."""
    return Mock()


@pytest.fixture
def generator(mock_llm_client):
    """Create code generator instance."""
    return CodeGenerator(mock_llm_client)


def test_generate_initial_code(generator, mock_llm_client):
    """Test initial code generation."""
    # Mock LLM response
    mock_response = GeneratedCode(
        code="from pptx import Presentation\nprs = Presentation()",
        explanation="Creates a presentation",
        expected_output_filename="test.pptx"
    )
    mock_llm_client.generate_structured.return_value = mock_response

    # Generate code
    result = generator.generate_initial_code("Create a simple presentation")

    assert result.code == mock_response.code
    assert len(generator.conversation_history) == 3  # system + user + assistant
    mock_llm_client.generate_structured.assert_called_once()


def test_generate_initial_code_with_style(generator, mock_llm_client):
    """Test initial code generation with style guidelines."""
    # Mock LLM response
    mock_response = GeneratedCode(
        code="from pptx import Presentation\nprs = Presentation()",
        explanation="Creates a styled presentation",
        expected_output_filename="test.pptx"
    )
    mock_llm_client.generate_structured.return_value = mock_response

    # Generate code with style
    style_content = "Use blue colors and bold fonts"
    result = generator.generate_initial_code("Create a presentation", style_content)

    assert result.code == mock_response.code
    assert generator.style_content == style_content

    # Check that style was included in user message
    user_message = generator.conversation_history[1]["content"]
    assert "STYLE GUIDELINES" in user_message
    assert "Use blue colors" in user_message
    assert len(generator.conversation_history) == 3
    mock_llm_client.generate_structured.assert_called_once()


def test_fix_code_after_error(generator, mock_llm_client):
    """Test code fixing after error."""
    # Setup conversation history
    generator.conversation_history = [
        {"role": "system", "content": "system prompt"},
        {"role": "user", "content": "task"}
    ]

    # Mock LLM response
    mock_response = GeneratedCode(
        code="from pptx import Presentation\nfixed code",
        explanation="Fixed the error",
        expected_output_filename="test.pptx"
    )
    mock_llm_client.generate_structured.return_value = mock_response

    # Create error result
    error_result = CodeExecutionResult(
        success=False,
        error_message="ImportError",
        traceback="Traceback...",
        stderr="Error details"
    )

    # Fix code
    result = generator.fix_code_after_error("original code", error_result)

    assert result.code == mock_response.code
    assert len(generator.conversation_history) == 4  # added user request + assistant response
    mock_llm_client.generate_structured.assert_called_once()


def test_fix_code_after_validation(generator, mock_llm_client):
    """Test code fixing after validation failure."""
    generator.conversation_history = [
        {"role": "system", "content": "system prompt"}
    ]

    # Mock LLM response
    mock_response = GeneratedCode(
        code="from pptx import Presentation\nfixed validation",
        explanation="Fixed validation issues",
        expected_output_filename="test.pptx"
    )
    mock_llm_client.generate_structured.return_value = mock_response

    # Create validation result
    validation_result = ValidationResult(
        valid=False,
        slide_count=2,
        expected_slide_count=3,
        has_titles=True,
        issues=["Slide count mismatch"]
    )

    # Fix code
    result = generator.fix_code_after_validation(
        "original code",
        validation_result,
        "task content"
    )

    assert result.code == mock_response.code
    assert len(generator.conversation_history) == 3
    mock_llm_client.generate_structured.assert_called_once()
