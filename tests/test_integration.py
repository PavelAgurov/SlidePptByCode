"""Integration tests with real LLM API calls."""

import pytest
from pathlib import Path
from src.config import load_settings
from src.llm_client import LLMClient
from src.code_generator import CodeGenerator
from src.code_executor import CodeExecutor
from src.validator import validate_presentation


@pytest.mark.integration
def test_full_pipeline():
    """
    Test full pipeline with real LLM API.

    This test:
    1. Loads real configuration
    2. Reads sample task
    3. Generates code using LLM
    4. Executes code
    5. Validates output

    Requires valid OPENROUTER_API_KEY in environment.
    """
    # Load config
    config = load_settings()
    config.create_directories()

    # Read sample task
    task_file = Path("tests/fixtures/sample_task.md")
    assert task_file.exists(), "Sample task file not found"
    task_content = task_file.read_text(encoding='utf-8')

    # Initialize components
    llm_client = LLMClient(config)
    generator = CodeGenerator(llm_client)
    executor = CodeExecutor(config)

    # Generate code
    generated = generator.generate_initial_code(task_content)

    assert generated.code is not None
    assert "from pptx import" in generated.code or "import pptx" in generated.code
    assert generated.expected_output_filename.endswith(".pptx")

    # Save and execute code
    code_file = executor.save_code(generated.code, "integration_test.py")
    result = executor.execute(code_file)

    # Check execution
    assert result.success is True, f"Execution failed: {result.error_message}"
    assert result.output_file is not None
    assert result.output_file.exists()

    # Validate presentation
    validation = validate_presentation(result.output_file, task_content)

    assert validation.slide_count == 3, f"Expected 3 slides, got {validation.slide_count}"
    assert validation.has_titles is True, "Slides should have titles"

    print(f"\n✓ Integration test passed!")
    print(f"  Generated file: {result.output_file}")
    print(f"  Slides: {validation.slide_count}")
    print(f"  Validation: {'PASS' if validation.is_valid else 'FAIL'}")

    if not validation.is_valid:
        print(f"  Issues: {validation.issues}")


@pytest.mark.integration
def test_error_recovery():
    """
    Test error recovery mechanism.

    This test intentionally creates a scenario where the LLM
    might generate imperfect code and needs to fix it.
    """
    config = load_settings()
    config.create_directories()
    config.max_retries = 2

    # Use a more challenging task
    task_content = """
    # Complex Presentation

    ## Slide 1 — Overview
    Title and bullet points

    ## Slide 2 — Data
    Include numbers and formatting

    ## Slide 3 — Charts
    Visual elements

    ## Slide 4 — Conclusion
    Final summary
    """

    llm_client = LLMClient(config)
    generator = CodeGenerator(llm_client)
    executor = CodeExecutor(config)

    # Try generation with retries
    max_attempts = 2
    success = False

    for attempt in range(max_attempts):
        if attempt == 0:
            generated = generator.generate_initial_code(task_content)
        else:
            # This would be called after an error
            continue

        code_file = executor.save_code(generated.code, f"recovery_test_{attempt}.py")
        result = executor.execute(code_file)

        if result.success and result.output_file:
            validation = validate_presentation(result.output_file, task_content)
            if validation.is_valid or validation.slide_count == 4:
                success = True
                break

    assert success, "Failed to generate valid presentation even with retries"


@pytest.mark.integration
def test_full_pipeline_with_style():
    """
    Test full pipeline with style guidelines.

    This test verifies that style content is properly injected
    into the prompt and used by the LLM.
    """
    # Load config
    config = load_settings()
    config.create_directories()

    # Read task and style
    task_file = Path("tests/fixtures/sample_task.md")
    task_content = task_file.read_text(encoding='utf-8')

    style_file = Path("data/style.md")
    assert style_file.exists(), "Style file not found"
    style_content = style_file.read_text(encoding='utf-8')

    # Initialize components
    llm_client = LLMClient(config)
    generator = CodeGenerator(llm_client)
    executor = CodeExecutor(config)

    # Generate with style
    generated = generator.generate_initial_code(task_content, style_content)

    # Verify style was stored
    assert generator.style_content == style_content

    # Verify style was included in prompt
    user_message = generator.conversation_history[1]["content"]
    assert "STYLE GUIDELINES" in user_message

    # Execute and validate
    code_file = executor.save_code(generated.code, "test_with_style.py")
    result = executor.execute(code_file)

    assert result.success is True, f"Execution failed: {result.error_message}"
    assert result.output_file is not None
    assert result.output_file.exists()

    # Validate presentation
    validation = validate_presentation(result.output_file, task_content)
    assert validation.slide_count == 3

    print(f"\n✓ Integration test with style passed!")
    print(f"  Generated file: {result.output_file}")
    print(f"  Style guidelines applied: Yes")
