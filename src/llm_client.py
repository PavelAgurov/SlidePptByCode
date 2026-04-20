"""OpenAI client wrapper for OpenRouter API."""

import logging
from typing import TypeVar, Type
from openai import OpenAI
from pydantic import BaseModel
from .config import Settings

logger = logging.getLogger(__name__)

T = TypeVar('T', bound=BaseModel)


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
        logger.info(f"Initialized LLM client with model: {self.model_id}")

    def generate_structured(
        self,
        messages: list[dict],
        response_format: Type[T]
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
                messages=messages,
                response_format=response_format,
                temperature=self.temperature
            )

            result = response.choices[0].message.parsed
            logger.debug(f"Received structured response: {type(result).__name__}")
            return result

        except Exception as e:
            logger.error(f"LLM API call failed: {e}")
            raise
