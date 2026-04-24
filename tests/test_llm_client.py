"""Tests for LLMClient session token usage accumulation."""

from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest

from src.llm_client import LLMClient, SessionTokenUsage
from src.models import GeneratedCode
from src.snippets import SnippetCallCache


def _make_response(
    parsed: GeneratedCode,
    *,
    prompt: int,
    completion: int,
    total: int | None,
    cached: int = 0,
) -> SimpleNamespace:
    details = SimpleNamespace(cached_tokens=cached) if cached else None
    usage = SimpleNamespace(
        prompt_tokens=prompt,
        completion_tokens=completion,
        total_tokens=total,
        prompt_tokens_details=details,
    )
    message = SimpleNamespace(parsed=parsed)
    choice = SimpleNamespace(message=message)
    return SimpleNamespace(choices=[choice], usage=usage)


@pytest.fixture
def mock_config() -> MagicMock:
    cfg = MagicMock()
    cfg.base_url = "https://example.test/v1"
    cfg.openrouter_api_key = "sk-test"
    cfg.model_id = "test-model"
    cfg.temperature = 0.0
    cfg.llm_max_tokens = 4096
    return cfg


@patch("src.llm_client.OpenAI")
def test_session_token_usage_sums_two_calls(mock_openai_class: MagicMock, mock_config: MagicMock) -> None:
    mock_api = MagicMock()
    mock_openai_class.return_value = mock_api

    sample = GeneratedCode(
        code="from pptx import Presentation\nx=1\n",
        explanation="test",
        expected_output_filename="out.pptx",
    )
    responses = [
        _make_response(sample, prompt=100, completion=50, total=150, cached=10),
        _make_response(sample, prompt=200, completion=80, total=None, cached=20),
    ]
    mock_api.beta.chat.completions.parse.side_effect = responses

    client = LLMClient(mock_config)
    fmt = GeneratedCode
    client.generate_structured([{"role": "user", "content": "a"}], fmt)
    client.generate_structured([{"role": "user", "content": "b"}], fmt)

    u = client.session_token_usage()
    assert u == SessionTokenUsage(
        prompt_tokens=300,
        completion_tokens=130,
        cached_tokens=30,
        total_tokens=150 + 280,
    )


@patch("src.llm_client.OpenAI")
def test_session_token_usage_missing_total_falls_back(mock_openai_class: MagicMock, mock_config: MagicMock) -> None:
    mock_api = MagicMock()
    mock_openai_class.return_value = mock_api

    sample = GeneratedCode(
        code="from pptx import Presentation\nx=1\n",
        explanation="test",
        expected_output_filename="out.pptx",
    )
    usage = SimpleNamespace(
        prompt_tokens=40,
        completion_tokens=60,
        total_tokens=None,
        prompt_tokens_details=None,
    )
    message = SimpleNamespace(parsed=sample)
    mock_api.beta.chat.completions.parse.return_value = SimpleNamespace(
        choices=[SimpleNamespace(message=message)],
        usage=usage,
    )

    client = LLMClient(mock_config)
    client.generate_structured([{"role": "user", "content": "x"}], GeneratedCode)
    u = client.session_token_usage()
    assert u.total_tokens == 100
    assert u.cached_tokens == 0


@patch("src.llm_client.OpenAI")
def test_generate_structured_with_snippet_tools_tool_then_parsed(
    mock_openai_class: MagicMock, mock_config: MagicMock
) -> None:
    """First parse returns tool_calls; second returns structured output."""
    mock_api = MagicMock()
    mock_openai_class.return_value = mock_api

    parsed = GeneratedCode(
        code="from pptx import Presentation\nok\n",
        explanation="done",
        expected_output_filename="out.pptx",
    )
    tool_fn = SimpleNamespace(
        name="get_code_snippet",
        arguments='{"snippet_id":"title_slide"}',
    )
    tool_call = SimpleNamespace(
        id="call_abc",
        type="function",
        function=tool_fn,
    )
    msg_tool = SimpleNamespace(parsed=None, content=None, tool_calls=[tool_call])
    msg_done = SimpleNamespace(parsed=parsed, content=None, tool_calls=None)

    usage1 = SimpleNamespace(
        prompt_tokens=10,
        completion_tokens=5,
        total_tokens=15,
        prompt_tokens_details=None,
    )
    usage2 = SimpleNamespace(
        prompt_tokens=20,
        completion_tokens=10,
        total_tokens=30,
        prompt_tokens_details=None,
    )
    mock_api.beta.chat.completions.parse.side_effect = [
        SimpleNamespace(
            choices=[SimpleNamespace(message=msg_tool)],
            usage=usage1,
        ),
        SimpleNamespace(
            choices=[SimpleNamespace(message=msg_done)],
            usage=usage2,
        ),
    ]

    client = LLMClient(mock_config)
    out = client.generate_structured_with_snippet_tools(
        [{"role": "user", "content": "hi"}],
        GeneratedCode,
        SnippetCallCache(),
    )
    assert out == parsed
    assert mock_api.beta.chat.completions.parse.call_count == 2
    u = client.session_token_usage()
    assert u.prompt_tokens == 30
    assert u.completion_tokens == 15
