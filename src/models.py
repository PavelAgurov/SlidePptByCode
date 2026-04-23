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


class ChunkedSectionGeneratedCode(BaseModel):
    """Single H2-section fragment for chunked generation (no full script)."""

    code: str = Field(
        description="Python defining exactly one def add_section_NNN(prs): ..."
    )
    explanation: str = Field(
        default="",
        description="Brief note on the slide content",
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
