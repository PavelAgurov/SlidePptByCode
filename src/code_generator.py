"""Code generation and error fixing using LLM."""

import logging
from typing import Any

from .llm_client import LLMClient
from .models import (
    ChunkedSectionGeneratedCode,
    CodeExecutionResult,
    GeneratedCode,
    IncrementalLlmScriptCode,
    ValidationResult,
)
from .prompts import (
    SYSTEM_PROMPT,
    SYSTEM_PROMPT_INCREMENTAL_H1,
    SYSTEM_PROMPT_INCREMENTAL_H2,
    format_chunk_h2_slide_user_message,
    format_chunk_title_deck_user_message,
    format_error_fix_prompt,
    format_incremental_h1_user_message,
    format_incremental_h1_validation_fix_prompt,
    format_incremental_h2_user_message,
    format_incremental_h2_validation_fix_prompt,
    format_validation_fix_prompt,
)
from .task_chunker import TaskChunk, extract_deck_title, section_function_name
from .code_merger import merge_chunked_modules

logger = logging.getLogger(__name__)


def _incremental_raw_to_generated(raw: IncrementalLlmScriptCode) -> GeneratedCode:
    """Build ``GeneratedCode`` without re-running import validators (incremental LLM output)."""
    return GeneratedCode.model_construct(
        code=raw.code.strip(),
        explanation=raw.explanation,
        expected_output_filename=raw.expected_output_filename,
    )


class CodeGenerator:
    """Generates and fixes Python code using LLM."""

    def __init__(self, llm_client: LLMClient):
        """Initialize code generator with LLM client."""
        self.llm_client = llm_client
        self.conversation_history: list[dict[str, Any]] = []
        self.style_content: str | None = None

    def generate_initial_code(self, task_content: str, style_content: str | None = None, language: str | None = None, output_filename: str | None = None) -> GeneratedCode:
        """
        Generate initial Python code for presentation task.

        Args:
            task_content: Text description of the presentation to create
            style_content: Optional style guidelines to apply
            language: Optional language for presentation content
            output_filename: Optional output filename for the presentation

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

        # Inject language requirement if provided
        if language:
            user_message += f"\n\n## LANGUAGE REQUIREMENT:\n\nAll text content in the presentation MUST be in {language}. Translate all slide titles, bullet points, and text content to {language}."
            logger.info(f"Language requirement included: {language}")

        # Inject output filename requirement if provided
        if output_filename:
            user_message += f"\n\n## OUTPUT FILENAME:\n\nYou MUST save the presentation to: {output_filename}\n\nUse this exact path in your code. Do not generate a different filename."
            logger.info(f"Output filename specified: {output_filename}")

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

    def reset_conversation_with_merged_code(self, code: str, explanation: str = "") -> None:
        """Replace history so retries operate on merged script only (chunked pipeline)."""
        self.conversation_history = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "assistant",
                "content": f"Code:\n{code}\n\nExplanation: {explanation}",
            },
        ]

    def generate_chunked_code(
        self,
        chunks: list[TaskChunk],
        full_task_content: str,
        style_content: str | None = None,
        language: str | None = None,
        output_filename: str | None = None,
    ) -> GeneratedCode:
        """
        Generate script via one LLM call per chunk (preamble + each H2 section), then merge.

        Caller must ensure the task was validated (H1 before first H2, at least one ``##``).
        """
        logger.info("Chunked generation: %d chunk(s)", len(chunks))
        self.style_content = style_content

        preamble = next(c for c in chunks if c.kind == "preamble")
        h2_chunks = [c for c in chunks if c.kind == "h2"]
        if not h2_chunks:
            raise ValueError("generate_chunked_code requires at least one H2 chunk")

        deck_title = extract_deck_title(full_task_content)
        if deck_title is None:
            raise ValueError(
                "generate_chunked_code requires an H1 title in the task (validate_chunked_task_markdown first)"
            )
        total_parts = len(chunks)

        preamble_user = format_chunk_title_deck_user_message(
            title_block_markdown=preamble.body,
            deck_title=deck_title,
            total_h2_slides=len(h2_chunks),
            style_content=style_content,
            language=language,
            output_filename=output_filename,
        )
        messages0 = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": preamble_user},
        ]
        part0 = self.llm_client.generate_structured(messages0, GeneratedCode)
        logger.info("Chunked part 0 (title deck) received")

        section_codes: list[str] = []
        explanation_parts: list[str] = [part0.explanation]
        for ord1, c in enumerate(h2_chunks, start=1):
            fn = section_function_name(c.index)
            section_user = format_chunk_h2_slide_user_message(
                section_markdown=c.body,
                function_name=fn,
                part_number=c.index + 1,
                total_parts=total_parts,
                section_ordinal=ord1,
                num_h2_slides=len(h2_chunks),
                deck_title=deck_title,
                style_content=style_content,
                language=language,
            )
            messages_i = [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": section_user},
            ]
            sec = self.llm_client.generate_structured(
                messages_i, ChunkedSectionGeneratedCode
            )
            section_codes.append(sec.code)
            explanation_parts.append(sec.explanation)
            logger.info("Chunked part %d/%d (%s) received", c.index + 1, total_parts, fn)

        merged = merge_chunked_modules(part0.code, section_codes)
        tail = " | ".join(explanation_parts[:5])
        if len(explanation_parts) > 5:
            tail += " ..."
        combined_explanation = (
            f"Chunked merge ({len(h2_chunks)} H2 section(s)): {tail}"
        )

        result = GeneratedCode(
            code=merged,
            explanation=combined_explanation,
            expected_output_filename=part0.expected_output_filename,
        )
        logger.info("Chunked merge complete, expected output: %s", result.expected_output_filename)
        return result

    def fix_code_after_error(
        self,
        original_code: str,
        error_result: CodeExecutionResult,
        *,
        incremental: bool = False,
    ) -> GeneratedCode:
        """
        Fix code after execution error.

        Args:
            original_code: The code that failed
            error_result: Execution result with error details

        Args:
            incremental: If True, parse LLM output as ``IncrementalLlmScriptCode`` (no import
                guard on parse); still returned as ``GeneratedCode`` via ``model_construct``.

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

        # Get fixed code (separate parse types so Pyright narrows correctly)
        if incremental:
            raw_inc = self.llm_client.generate_structured(
                messages=self.conversation_history,
                response_format=IncrementalLlmScriptCode,
            )
            result = _incremental_raw_to_generated(raw_inc)
        else:
            result = self.llm_client.generate_structured(
                messages=self.conversation_history,
                response_format=GeneratedCode,
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

    def generate_incremental_h1(
        self,
        preamble_markdown: str,
        deck_title: str,
        style_content: str | None = None,
        language: str | None = None,
        deck_from_template: bool = False,
    ) -> GeneratedCode:
        """Generate H1 script: fill deck stub and write ``shared.py``."""
        logger.info("Incremental H1 generation")
        self.style_content = style_content
        user = format_incremental_h1_user_message(
            preamble_markdown=preamble_markdown,
            deck_title=deck_title,
            style_content=style_content,
            language=language,
            deck_from_template=deck_from_template,
        )
        self.conversation_history = [
            {"role": "system", "content": SYSTEM_PROMPT_INCREMENTAL_H1},
            {"role": "user", "content": user},
        ]
        raw = self.llm_client.generate_structured(
            messages=self.conversation_history,
            response_format=IncrementalLlmScriptCode,
        )
        result = _incremental_raw_to_generated(raw)
        self.conversation_history.append(
            {
                "role": "assistant",
                "content": f"Code:\n{result.code}\n\nExplanation: {result.explanation}",
            }
        )
        return result

    def fix_incremental_h1_validation(
        self,
        original_code: str,
        issues: list[str],
        preamble_markdown: str,
    ) -> GeneratedCode:
        """Fix H1 script after incremental validation failure."""
        logger.info("Fixing incremental H1 after validation: %s", issues)
        prompt = format_incremental_h1_validation_fix_prompt(
            original_code=original_code,
            issues=issues,
            preamble_markdown=preamble_markdown,
        )
        self.conversation_history.append({"role": "user", "content": prompt})
        raw = self.llm_client.generate_structured(
            messages=self.conversation_history,
            response_format=IncrementalLlmScriptCode,
        )
        result = _incremental_raw_to_generated(raw)
        self.conversation_history.append(
            {
                "role": "assistant",
                "content": f"Fixed code:\n{result.code}\n\nExplanation: {result.explanation}",
            }
        )
        return result

    def generate_incremental_slide(
        self,
        section_markdown: str,
        deck_title: str,
        section_ordinal: int,
        num_h2_slides: int,
        style_content: str | None = None,
        language: str | None = None,
    ) -> GeneratedCode:
        """Generate one H2 slide-append script (fresh conversation)."""
        logger.info(
            "Incremental H2 slide generation (%d/%d)",
            section_ordinal,
            num_h2_slides,
        )
        self.style_content = style_content
        user = format_incremental_h2_user_message(
            section_markdown=section_markdown,
            deck_title=deck_title,
            section_ordinal=section_ordinal,
            num_h2_slides=num_h2_slides,
            style_content=style_content,
            language=language,
        )
        self.conversation_history = [
            {"role": "system", "content": SYSTEM_PROMPT_INCREMENTAL_H2},
            {"role": "user", "content": user},
        ]
        raw = self.llm_client.generate_structured(
            messages=self.conversation_history,
            response_format=IncrementalLlmScriptCode,
        )
        result = _incremental_raw_to_generated(raw)
        self.conversation_history.append(
            {
                "role": "assistant",
                "content": f"Code:\n{result.code}\n\nExplanation: {result.explanation}",
            }
        )
        return result

    def fix_incremental_h2_validation(
        self,
        original_code: str,
        issues: list[str],
        deck_title: str,
        section_markdown: str,
    ) -> GeneratedCode:
        """Fix one H2 slide script after incremental validation failure."""
        logger.info("Fixing incremental H2 after validation: %s", issues)
        prompt = format_incremental_h2_validation_fix_prompt(
            original_code=original_code,
            issues=issues,
            deck_title=deck_title,
            section_markdown=section_markdown,
        )
        self.conversation_history.append({"role": "user", "content": prompt})
        raw = self.llm_client.generate_structured(
            messages=self.conversation_history,
            response_format=IncrementalLlmScriptCode,
        )
        result = _incremental_raw_to_generated(raw)
        self.conversation_history.append(
            {
                "role": "assistant",
                "content": f"Fixed code:\n{result.code}\n\nExplanation: {result.explanation}",
            }
        )
        return result
