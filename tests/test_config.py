"""Tests for configuration management."""

import pytest
from pathlib import Path
from src.config import Settings


def test_settings_defaults():
    """Test default settings values."""
    # Create settings with minimal required fields
    settings = Settings(openrouter_api_key="test_key")

    assert settings.model_id == "gpt-4.1"
    assert settings.base_url == "https://openrouter.ai/api/v1"
    assert settings.temperature == 0
    assert settings.max_retries == 3
    assert settings.execution_timeout == 60
    assert settings.generated_dir == Path(".generated")
    assert settings.output_dir == Path(".output")


def test_settings_custom_values():
    """Test custom settings values."""
    settings = Settings(
        openrouter_api_key="custom_key",
        model_id="gpt-4.1",
        temperature=0.5,
        max_retries=5
    )

    assert settings.openrouter_api_key == "custom_key"
    assert settings.model_id == "gpt-4.1"
    assert settings.temperature == 0.5
    assert settings.max_retries == 5


def test_create_directories(tmp_path):
    """Test directory creation."""
    test_generated = tmp_path / "test_generated"
    test_output = tmp_path / "test_output"

    settings = Settings(
        openrouter_api_key="test",
        generated_dir=test_generated,
        output_dir=test_output
    )

    assert not test_generated.exists()
    assert not test_output.exists()

    settings.create_directories()

    assert test_generated.exists()
    assert test_output.exists()
