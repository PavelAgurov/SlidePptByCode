"""OpenAI client wrapper for OpenRouter API."""

import logging
from collections.abc import Iterable
from typing import Any, TypeVar, cast

from openai import OpenAI
from openai.types.chat import ChatCompletionMessageParam
from pydantic import BaseModel

from .config import Settings

logger = logging.getLogger(__name__)

T = TypeVar("T", bound=BaseModel)


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

            result = response.choices[0].message.parsed
            logger.debug(f"Received structured response: {type(result).__name__}")
            if result is None:
                raise RuntimeError("LLM returned no parsed structured output")
            return result

        except Exception as e:
            logger.error(f"LLM API call failed: {e}")
            raise
