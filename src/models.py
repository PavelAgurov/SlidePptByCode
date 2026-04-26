"""Pydantic models for structured data and LLM responses."""

from pathlib import Path
from typing import Optional
from pydantic import BaseModel, Field, field_validator


class GeneratedCode(BaseModel):
    """Structure for LLM to return generated Python code."""

    code: str = Field(
        description="Complete Python code to generate the presentation"
    )
    explanation: str = Field(
        description="Brief explanation of the approach and key features"
    )
    expected_output_filename: str = Field(
        description="Expected .pptx filename (e.g., 'report.pptx')"
    )

    @field_validator('code')
    @classmethod
    def validate_code_has_imports(cls, v: str) -> str:
        """Ensure code imports python-pptx."""
        if 'from pptx import' not in v and 'import pptx' not in v:
            raise ValueError("Code must import python-pptx")
        return v


class IncrementalLlmScriptCode(BaseModel):
    """
    Structured LLM output for incremental H1 / per-slide H2 scripts.

    Unlike ``GeneratedCode``, this does **not** require a pptx import in the
    validator — the API parser must not reject odd model ordering. Runtime
    execution still requires valid python-pptx usage.
    """

    code: str = Field(
        description=(
            "One complete Python source file. Must include "
            "`from pptx import Presentation` or `import pptx`. "
            "For H1: embed shared module source as a string passed to "
            "SHARED_PY_PATH.write_text(...); never return only the raw shared.py "
            "file as this field without pptx imports and deck logic."
        )
    )
    explanation: str = Field(default="", description="Brief note on the approach")
    expected_output_filename: str = Field(
        default="presentation.pptx",
        description="Legacy field; incremental runner uses injected deck paths.",
    )


class ChunkedSectionGeneratedCode(BaseModel):
    """Single H2-section fragment for chunked generation (no full script)."""

    code: str = Field(
        description="Python defining exactly one def add_section_NNN(prs): ..."
    )
    explanation: str = Field(
        default="",
        description="Brief note on the slide content",
    )


class LayoutDescription(BaseModel):
    """Structured output for one slide layout: short intent summary for later selection."""

    description: str = Field(
        description="One short sentence (<=160 chars) describing the layout intent / typical use"
    )


class LayoutSelection(BaseModel):
    """Structured output for choosing a slide layout from a template."""

    explanation: str = Field(description="Why this layout fits the slide content")
    selected_layout_index: int = Field(
        ge=0,
        description="Index in prs.slide_layouts; MUST be one of the listed allowed indices",
    )


class CodeExecutionResult(BaseModel):
    """Result of executing generated code."""

    success: bool
    output_file: Optional[Path] = None
    error_message: Optional[str] = None
    stderr: Optional[str] = None
    stdout: Optional[str] = None
    traceback: Optional[str] = None


class ValidationResult(BaseModel):
    """Result of validating a generated presentation."""

    valid: bool
    slide_count: int
    expected_slide_count: Optional[int] = None
    has_titles: bool
    issues: list[str] = Field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        """Check if validation passed."""
        return self.valid and len(self.issues) == 0
