"""Tests for code generator."""

import pytest
from unittest.mock import Mock
from src.code_generator import CodeGenerator
from src.models import GeneratedCode, CodeExecutionResult, LayoutSelection, ValidationResult
from src.layout_catalog import LayoutInfo


@pytest.fixture
def mock_llm_client() -> Mock:
    """Create mock LLM client."""
    m = Mock()

    def _tools_return(*_a: object, **_k: object) -> object:
        return m.generate_structured_with_snippet_tools.return_value

    m.generate_structured_with_snippet_tools.side_effect = _tools_return
    return m


@pytest.fixture
def generator(mock_llm_client: Mock) -> CodeGenerator:
    """Create code generator instance."""
    return CodeGenerator(mock_llm_client)


def test_generate_initial_code(generator: CodeGenerator, mock_llm_client: Mock) -> None:
    """Test initial code generation."""
    # Mock LLM response
    mock_response = GeneratedCode(
        code="from pptx import Presentation\nprs = Presentation()",
        explanation="Creates a presentation",
        expected_output_filename="test.pptx"
    )
    mock_llm_client.generate_structured_with_snippet_tools.return_value = mock_response

    # Generate code
    result = generator.generate_initial_code("Create a simple presentation")

    assert result.code == mock_response.code
    assert len(generator.conversation_history) == 3  # system + user + assistant
    mock_llm_client.generate_structured_with_snippet_tools.assert_called_once()


def test_generate_initial_code_with_style(generator: CodeGenerator, mock_llm_client: Mock) -> None:
    """Test initial code generation with style guidelines."""
    # Mock LLM response
    mock_response: GeneratedCode = GeneratedCode(
        code="from pptx import Presentation\nprs = Presentation()",
        explanation="Creates a styled presentation",
        expected_output_filename="test.pptx"
    )
    mock_llm_client.generate_structured_with_snippet_tools.return_value = mock_response

    # Generate code with style
    style_content = "Use blue colors and bold fonts"
    result: GeneratedCode = generator.generate_initial_code("Create a presentation", style_content)

    assert result.code == mock_response.code
    assert generator.style_content == style_content

    # Check that style was included in user message
    user_message = generator.conversation_history[1]["content"]
    assert "STYLE GUIDELINES" in user_message
    assert "Use blue colors" in user_message
    assert len(generator.conversation_history) == 3
    mock_llm_client.generate_structured_with_snippet_tools.assert_called_once()


def test_fix_code_after_error(generator: CodeGenerator, mock_llm_client: Mock) -> None:
    """Test code fixing after error."""
    # Setup conversation history
    generator.conversation_history = [
        {"role": "system", "content": "system prompt"},
        {"role": "user", "content": "task"}
    ]

    # Mock LLM response
    mock_response: GeneratedCode = GeneratedCode(
        code="from pptx import Presentation\nfixed code",
        explanation="Fixed the error",
        expected_output_filename="test.pptx"
    )
    mock_llm_client.generate_structured_with_snippet_tools.return_value = mock_response

    # Create error result
    error_result: CodeExecutionResult = CodeExecutionResult(
        success=False,
        error_message="ImportError",
        traceback="Traceback...",
        stderr="Error details"
    )

    # Fix code
    result = generator.fix_code_after_error("original code", error_result)

    assert result.code == mock_response.code
    assert len(generator.conversation_history) == 4  # added user request + assistant response
    mock_llm_client.generate_structured_with_snippet_tools.assert_called_once()


def test_fix_code_after_error_incremental_compacts_history(
    generator: CodeGenerator, mock_llm_client: Mock
) -> None:
    """Incremental execution error uses compact history and shared index in prompt."""
    from src.models import IncrementalLlmScriptCode

    generator.conversation_history = [
        {"role": "system", "content": "SYS_H2"},
        {"role": "user", "content": "TASK_BODY"},
        {"role": "assistant", "content": "Code:\nbad\n\nExplanation: x"},
    ]
    mock_llm_client.generate_structured_with_snippet_tools.return_value = (
        IncrementalLlmScriptCode(
            code="from pptx import Presentation\nfixed",
            explanation="ok",
        )
    )
    err = CodeExecutionResult(
        success=False,
        error_message="boom",
        traceback="Traceback...",
        stderr="e",
    )
    generator.fix_code_after_error(
        "bad code",
        err,
        incremental=True,
        shared_index="ELF_GREEN = ...",
    )
    assert len(generator.conversation_history) == 5
    last_user = generator.conversation_history[3]["content"]
    assert "EXECUTION FAILED" in last_user
    assert "Traceback" in last_user
    assert "ELF_GREEN" in last_user
    assert "bad code" in last_user
    mock_llm_client.generate_structured_with_snippet_tools.assert_called_once()


def test_fix_code_after_validation(generator: CodeGenerator, mock_llm_client: Mock) -> None:
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
    mock_llm_client.generate_structured_with_snippet_tools.return_value = mock_response

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
    mock_llm_client.generate_structured_with_snippet_tools.assert_called_once()


def test_select_layout_does_not_mutate_conversation_history(
    generator: CodeGenerator, mock_llm_client: Mock
) -> None:
    mock_llm_client.generate_structured.return_value = LayoutSelection(
        explanation="ok",
        selected_layout_index=0,
    )
    generator.conversation_history = [
        {"role": "system", "content": "SYS"},
        {"role": "user", "content": "U"},
    ]
    layouts = [LayoutInfo(index=0, name="L0", placeholders=())]
    out = generator.select_layout(
        slide_markdown="## x",
        deck_title="Deck",
        layouts=layouts,
    )
    assert out.selected_layout_index == 0
    assert len(generator.conversation_history) == 2
    mock_llm_client.generate_structured.assert_called_once()
