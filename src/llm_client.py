"""OpenAI client wrapper for OpenRouter API."""

import json
import logging
from collections.abc import Iterable
from dataclasses import dataclass
from typing import Any, TypeVar, cast

from openai import OpenAI
from openai.types.chat import ChatCompletionMessageParam
from pydantic import BaseModel

from .config import Settings
from .snippets import get_code_snippet_tools, handle_get_code_snippet

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

    def generate_structured_with_snippet_tools(
        self,
        messages: list[dict[str, Any]],
        response_format: type[T],
        snippet_cache: Any,
        *,
        max_tool_rounds: int = 12,
    ) -> T:
        """
        Structured output with ``get_code_snippet`` tool rounds, then parsed result.

        Uses ``beta.chat.completions.parse`` with ``tools`` on each round (gpt-4.1 family).
        """
        tools = get_code_snippet_tools()
        msgs: list[dict[str, Any]] = list(messages)
        rounds = 0

        while rounds < max_tool_rounds:
            rounds += 1
            logger.debug(
                "LLM parse+tools round %s/%s",
                rounds,
                max_tool_rounds,
                extra={"color_event": "llm_round"},
            )
            response = self.client.beta.chat.completions.parse(
                model=self.model_id,
                messages=cast(Iterable[ChatCompletionMessageParam], msgs),
                response_format=response_format,
                tools=tools,
                temperature=self.temperature,
                max_tokens=self.max_tokens,
            )
            self._record_usage(response.usage)
            msg = response.choices[0].message

            if msg.parsed is not None:
                logger.info(
                    "LLM structured output: %s (after %s tool round(s))",
                    getattr(response_format, "__name__", str(response_format)),
                    rounds - 1,
                    extra={"color_event": "llm_structured_ok"},
                )
                return msg.parsed

            if not msg.tool_calls:
                logger.error(
                    "LLM returned neither parsed structured output nor tool_calls",
                    extra={"color_event": "tool_error"},
                )
                raise RuntimeError(
                    "LLM returned neither parsed structured output nor tool_calls"
                )

            assistant_payload: dict[str, Any] = {
                "role": "assistant",
                "content": msg.content,
            }
            assistant_payload["tool_calls"] = [
                {
                    "id": tc.id,
                    "type": getattr(tc, "type", None) or "function",
                    "function": {
                        "name": tc.function.name,
                        "arguments": tc.function.arguments or "{}",
                    },
                }
                for tc in msg.tool_calls
            ]
            msgs.append(assistant_payload)

            for tc in msg.tool_calls:
                name = tc.function.name
                arguments = tc.function.arguments or "{}"
                snippet_id = ""
                if name == "get_code_snippet":
                    try:
                        parsed_args = json.loads(arguments)
                        snippet_id = str(parsed_args.get("snippet_id", ""))
                    except (json.JSONDecodeError, TypeError):
                        snippet_id = "(bad json)"
                logger.info(
                    "LLM tool_call: %s%s",
                    name,
                    f" snippet_id={snippet_id!r}" if snippet_id else "",
                    extra={"color_event": "tool_call"},
                )
                if name == "get_code_snippet":
                    try:
                        body = handle_get_code_snippet(arguments, snippet_cache)
                    except RuntimeError as e:
                        body = str(e)
                        logger.warning(
                            "LLM tool guard: %s",
                            body,
                            extra={"color_event": "tool_error"},
                        )
                else:
                    body = f"Unknown tool: {name}"
                    logger.warning(
                        "LLM unknown tool: %s",
                        name,
                        extra={"color_event": "tool_error"},
                    )
                reuse = body.startswith("(Snippet") or body.startswith("Unknown snippet_id")
                logger.info(
                    "LLM tool_result: %s -> %d chars%s",
                    name,
                    len(body),
                    " [reuse/short]" if reuse else "",
                    extra={"color_event": "tool_result"},
                )
                msgs.append(
                    {
                        "role": "tool",
                        "tool_call_id": tc.id,
                        "content": body,
                    }
                )

        logger.error(
            "Exceeded max_tool_rounds=%s without structured output",
            max_tool_rounds,
            extra={"color_event": "tool_error"},
        )
        raise RuntimeError(
            f"Exceeded max_tool_rounds={max_tool_rounds} without structured output"
        )
