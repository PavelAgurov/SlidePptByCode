"""OpenAI client wrapper for OpenRouter API."""

import logging
from collections.abc import Iterable
from dataclasses import dataclass
from typing import Any, TypeVar, cast

from openai import OpenAI
from openai.types.chat import ChatCompletionMessageParam
from pydantic import BaseModel

from .config import Settings

logger = logging.getLogger(__name__)

T = TypeVar("T", bound=BaseModel)


@dataclass(frozen=True)
class SessionTokenUsage:
    """Cumulative token counts for all successful API calls in this client session."""

    prompt_tokens: int
    completion_tokens: int
    cached_tokens: int
    total_tokens: int


class LLMClient:
    """Wrapper around OpenAI client configured for OpenRouter."""

    def __init__(self, config: Settings):
        """Initialize LLM client with configuration."""
        self.client = OpenAI(
            base_url=config.base_url,
            api_key=config.openrouter_api_key
        )
        self.model_id = config.model_id
        self.temperature = config.temperature
        self.max_tokens = config.llm_max_tokens
        logger.info(
            "Initialized LLM client with model: %s, max_tokens=%s",
            self.model_id,
            self.max_tokens,
        )
        self._usage_prompt = 0
        self._usage_completion = 0
        self._usage_cached = 0
        self._usage_total = 0

    def session_token_usage(self) -> SessionTokenUsage:
        """Return summed usage since this client was constructed."""
        return SessionTokenUsage(
            prompt_tokens=self._usage_prompt,
            completion_tokens=self._usage_completion,
            cached_tokens=self._usage_cached,
            total_tokens=self._usage_total,
        )

    def _record_usage(self, usage: object | None) -> None:
        if usage is None:
            return
        prompt = int(getattr(usage, "prompt_tokens", None) or 0)
        completion = int(getattr(usage, "completion_tokens", None) or 0)
        total_attr = getattr(usage, "total_tokens", None)
        total = (
            int(total_attr)
            if total_attr is not None
            else prompt + completion
        )
        cached = 0
        details = getattr(usage, "prompt_tokens_details", None)
        if details is not None:
            cached = int(getattr(details, "cached_tokens", None) or 0)
        self._usage_prompt += prompt
        self._usage_completion += completion
        self._usage_cached += cached
        self._usage_total += total

    def generate_structured(
        self,
        messages: list[dict[str, Any]],
        response_format: type[T],
    ) -> T:
        """
        Generate structured output using Pydantic model.

        Args:
            messages: List of message dicts with 'role' and 'content'
            response_format: Pydantic model class for structured output

        Returns:
            Instance of response_format model with LLM response

        Raises:
            Exception: If API call fails
        """
        logger.debug(f"Sending request to LLM: {len(messages)} messages")

        try:
            response = self.client.beta.chat.completions.parse(
                model=self.model_id,
                messages=cast(
                    Iterable[ChatCompletionMessageParam],
                    messages,
                ),
                response_format=response_format,
                temperature=self.temperature,
                max_tokens=self.max_tokens,
            )

            self._record_usage(response.usage)

            result = response.choices[0].message.parsed
            logger.debug(f"Received structured response: {type(result).__name__}")
            if result is None:
                raise RuntimeError("LLM returned no parsed structured output")
            return result

        except Exception as e:
            logger.error(f"LLM API call failed: {e}")
            raise
