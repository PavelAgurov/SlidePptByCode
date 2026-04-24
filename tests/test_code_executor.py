"""Tests for code executor."""

import pytest
from pathlib import Path
from unittest.mock import Mock, patch
from src.config import Settings
from src.code_executor import CodeExecutor


@pytest.fixture
def mock_config(tmp_path):
    """Create mock configuration for tests."""
    generated_dir = tmp_path / ".generated"
    output_dir = tmp_path / ".output"
    generated_dir.mkdir()
    output_dir.mkdir()

    return Settings(
        openrouter_api_key="test",
        generated_dir=generated_dir,
        output_dir=output_dir,
        execution_timeout=10
    )


@pytest.fixture
def executor(mock_config):
    """Create code executor instance."""
    return CodeExecutor(mock_config)


def test_save_code(executor, tmp_path):
    """Test saving code to file."""
    code = "from pptx import Presentation\nprint('test')"

    code_path = executor.save_code(code, "test.py")

    assert code_path.exists()
    assert code_path.name.endswith("_test.py")
    assert code_path.read_text(encoding='utf-8') == code


def test_execute_success(executor, mock_config):
    """Test successful code execution."""
    # Create a simple Python file that creates a .pptx
    # Use absolute path to output directory
    output_dir = mock_config.output_dir
    code = f"""
from pptx import Presentation
from pathlib import Path

prs = Presentation()
output_path = Path(r'{output_dir}') / 'test_output.pptx'
output_path.parent.mkdir(exist_ok=True)
prs.save(str(output_path))
print('Created presentation')
"""

    code_file = executor.save_code(code, "success.py")
    result = executor.execute(code_file)

    assert result.success is True, f"Failed: {result.error_message}"
    assert result.output_file is not None
    assert result.output_file.exists()
    assert "Created presentation" in (result.stdout or "")


def test_execute_failure(executor):
    """Test failed code execution."""
    # Code that will fail
    code = "raise ValueError('Test error')"

    code_file = executor.save_code(code, "failure.py")
    result = executor.execute(code_file)

    assert result.success is False
    assert result.error_message is not None
    assert result.output_file is None
    assert "ValueError" in (result.stderr or "")
    assert "Test error" in (result.stderr or "")
    assert "ValueError" in result.error_message
    assert "Test error" in result.error_message


def test_execute_timeout(executor, mock_config):
    """Test execution timeout."""
    # Code that will timeout
    code = "import time\ntime.sleep(100)"

    # Set very short timeout
    executor.execution_timeout = 1

    code_file = executor.save_code(code, "timeout.py")
    result = executor.execute(code_file)

    assert result.success is False
    assert "timed out" in result.error_message.lower()


def test_execute_no_output(executor):
    """Test execution without creating .pptx file."""
    code = "print('No pptx created')"

    code_file = executor.save_code(code, "no_output.py")
    result = executor.execute(code_file)

    assert result.success is False
    assert "No .pptx file created" in result.error_message
