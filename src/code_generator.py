"""Code generation and error fixing using LLM."""

import logging
from .llm_client import LLMClient
from .models import GeneratedCode, CodeExecutionResult, ValidationResult
from .prompts import (
    SYSTEM_PROMPT,
    format_error_fix_prompt,
    format_validation_fix_prompt
)

logger = logging.getLogger(__name__)


class CodeGenerator:
    """Generates and fixes Python code using LLM."""

    def __init__(self, llm_client: LLMClient):
        """Initialize code generator with LLM client."""
        self.llm_client = llm_client
        self.conversation_history: list[dict] = []
        self.style_content: str | None = None

    def generate_initial_code(self, task_content: str, style_content: str | None = None) -> GeneratedCode:
        """
        Generate initial Python code for presentation task.

        Args:
            task_content: Text description of the presentation to create
            style_content: Optional style guidelines to apply

        Returns:
            GeneratedCode with code, explanation, and expected filename
        """
        logger.info("Generating initial code")

        # Store style for potential retries
        self.style_content = style_content

        # Build user message with task content
        user_message = f"Create a PowerPoint presentation based on this specification:\n\n{task_content}"

        # Inject style guidelines if provided
        if style_content:
            user_message += f"\n\n## STYLE GUIDELINES:\n\n{style_content}\n\nPlease apply these style guidelines when generating the presentation code."
            logger.info("Style guidelines included in prompt")

        # Build messages
        self.conversation_history = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_message}
        ]

        # Get structured response
        result = self.llm_client.generate_structured(
            messages=self.conversation_history,
            response_format=GeneratedCode
        )

        # Add assistant response to history
        self.conversation_history.append({
            "role": "assistant",
            "content": f"Code:\n{result.code}\n\nExplanation: {result.explanation}\nExpected output: {result.expected_output_filename}"
        })

        logger.info(f"Generated code, expected output: {result.expected_output_filename}")
        return result

    def fix_code_after_error(
        self,
        original_code: str,
        error_result: CodeExecutionResult
    ) -> GeneratedCode:
        """
        Fix code after execution error.

        Args:
            original_code: The code that failed
            error_result: Execution result with error details

        Returns:
            GeneratedCode with fixed code
        """
        logger.info(f"Fixing code after error: {error_result.error_message}")

        # Format error fix prompt
        error_prompt = format_error_fix_prompt(
            original_code=original_code,
            error_message=error_result.error_message or "Unknown error",
            traceback=error_result.traceback or "(no traceback)",
            stdout=error_result.stdout or "",
            stderr=error_result.stderr or ""
        )

        # Add to conversation
        self.conversation_history.append({
            "role": "user",
            "content": error_prompt
        })

        # Get fixed code
        result = self.llm_client.generate_structured(
            messages=self.conversation_history,
            response_format=GeneratedCode
        )

        # Add to history
        self.conversation_history.append({
            "role": "assistant",
            "content": f"Fixed code:\n{result.code}\n\nExplanation: {result.explanation}"
        })

        logger.info("Code fixed")
        return result

    def fix_code_after_validation(
        self,
        original_code: str,
        validation_result: ValidationResult,
        task_content: str
    ) -> GeneratedCode:
        """
        Fix code after validation failure.

        Args:
            original_code: The code that passed execution but failed validation
            validation_result: Validation result with issues
            task_content: Original task specification

        Returns:
            GeneratedCode with fixed code
        """
        logger.info(f"Fixing code after validation failure: {validation_result.issues}")

        # Format validation fix prompt
        validation_prompt = format_validation_fix_prompt(
            original_code=original_code,
            expected_slides=validation_result.expected_slide_count or 0,
            actual_slides=validation_result.slide_count,
            has_titles=validation_result.has_titles,
            issues=validation_result.issues,
            task_content=task_content
        )

        # Add to conversation
        self.conversation_history.append({
            "role": "user",
            "content": validation_prompt
        })

        # Get fixed code
        result = self.llm_client.generate_structured(
            messages=self.conversation_history,
            response_format=GeneratedCode
        )

        # Add to history
        self.conversation_history.append({
            "role": "assistant",
            "content": f"Fixed code:\n{result.code}\n\nExplanation: {result.explanation}"
        })

        logger.info("Code fixed after validation")
        return result
