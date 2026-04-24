"""Safe code execution in subprocess."""

import subprocess
import logging
import traceback
from pathlib import Path
from datetime import datetime
from .config import Settings
from .models import CodeExecutionResult

logger = logging.getLogger(__name__)


class CodeExecutor:
    """Executes generated Python code safely in subprocess."""

    def __init__(self, config: Settings):
        """Initialize code executor with configuration."""
        self.generated_dir = config.generated_dir
        self.output_dir = config.output_dir
        self.execution_timeout = config.execution_timeout

        # Determine Python executable path
        self.python_exe = self._find_python_executable()
        self.project_root = Path.cwd()

        logger.info(f"Initialized code executor with Python: {self.python_exe}")

    def _find_python_executable(self) -> Path:
        """Find the Python executable in venv."""
        # Try Windows path first
        venv_python = Path(".venv/Scripts/python.exe")
        if venv_python.exists():
            return venv_python

        # Try Unix path
        venv_python = Path(".venv/bin/python")
        if venv_python.exists():
            return venv_python

        # Fallback to system python
        logger.warning("venv Python not found, using system python")
        return Path("python")

    def _extract_output_path_from_stdout(self, stdout: str) -> Path | None:
        """
        Extract output file path from stdout.

        Looks for patterns like "Presentation saved to <path>"

        Args:
            stdout: Standard output from code execution

        Returns:
            Path to output file or None if not found
        """
        import re

        if not stdout:
            return None

        # Look for "Presentation saved to <path>" or "saved to <path>"
        patterns = [
            r'Presentation saved to (.+\.pptx)',
            r'saved to (.+\.pptx)',
            r'Saved to (.+\.pptx)',
            r'Created (.+\.pptx)',
            r'Output: (.+\.pptx)',
        ]

        for pattern in patterns:
            match = re.search(pattern, stdout, re.IGNORECASE)
            if match:
                path_str = match.group(1).strip()
                output_path = Path(path_str)
                logger.debug(f"Extracted output path from stdout: {output_path}")
                return output_path

        return None

    def save_code(self, code: str, filename: str) -> Path:
        """
        Save generated code to .generated directory.

        Args:
            code: Python code to save
            filename: Base filename (will add timestamp)

        Returns:
            Path to saved code file
        """
        # Add timestamp to filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename_with_timestamp = f"{timestamp}_{filename}"

        code_path = self.generated_dir / filename_with_timestamp
        code_path.write_text(code, encoding='utf-8')

        logger.info(f"Saved code to {code_path}")
        return code_path

    def execute(
        self,
        code_file: Path,
        *,
        expected_output_pptx: Path | None = None,
    ) -> CodeExecutionResult:
        """
        Execute Python code file in subprocess.

        Args:
            code_file: Path to Python file to execute
            expected_output_pptx: If set and the process exits 0, treat this path as
                the output deck when it exists (incremental pipeline; no stdout parsing).

        Returns:
            CodeExecutionResult with execution status and details
        """
        logger.info(f"Executing code: {code_file}")

        # Record execution start time
        import time
        execution_start_time = time.time()

        try:
            # Run subprocess
            result = subprocess.run(
                [str(self.python_exe), str(code_file)],
                cwd=str(self.project_root),
                capture_output=True,
                text=True,
                timeout=self.execution_timeout,
                check=False,
            )

            # Check return code
            if result.returncode == 0:
                if expected_output_pptx is not None:
                    exp = expected_output_pptx.resolve()
                    if exp.exists():
                        logger.info(
                            "Code executed successfully, using expected output: %s",
                            exp,
                        )
                        return CodeExecutionResult(
                            success=True,
                            output_file=exp,
                            stdout=result.stdout,
                            stderr=result.stderr,
                        )
                    logger.warning(
                        "Process exited 0 but expected pptx missing: %s", exp
                    )
                    return CodeExecutionResult(
                        success=False,
                        error_message=f"Expected output not found: {exp}",
                        stdout=result.stdout,
                        stderr=result.stderr,
                    )

                # Try to extract output path from stdout (code prints "Presentation saved to <path>")
                output_file = self._extract_output_path_from_stdout(result.stdout)

                if output_file and output_file.exists():
                    logger.info(f"Code executed successfully, output: {output_file}")
                    return CodeExecutionResult(
                        success=True,
                        output_file=output_file,
                        stdout=result.stdout,
                        stderr=result.stderr
                    )

                # Fallback: Check for .pptx files modified after execution started in .output/
                pptx_files = list(self.output_dir.glob("*.pptx"))
                modified_files = [
                    f for f in pptx_files
                    if f.stat().st_mtime >= execution_start_time
                ]

                if modified_files:
                    # Get the most recently modified file
                    output_file = max(modified_files, key=lambda p: p.stat().st_mtime)
                    logger.info(f"Code executed successfully, output: {output_file}")

                    return CodeExecutionResult(
                        success=True,
                        output_file=output_file,
                        stdout=result.stdout,
                        stderr=result.stderr
                    )
                else:
                    logger.warning("Code executed but no .pptx file found")
                    return CodeExecutionResult(
                        success=False,
                        error_message="No .pptx file created",
                        stdout=result.stdout,
                        stderr=result.stderr
                    )
            else:
                logger.error(f"Code execution failed with return code {result.returncode}")
                return CodeExecutionResult(
                    success=False,
                    error_message=f"Execution failed with return code {result.returncode}",
                    stderr=result.stderr,
                    stdout=result.stdout,
                    traceback=result.stderr  # stderr often contains traceback
                )

        except subprocess.TimeoutExpired as e:
            logger.error(f"Code execution timed out after {self.execution_timeout}s")
            return CodeExecutionResult(
                success=False,
                error_message=f"Execution timed out after {self.execution_timeout} seconds",
                stdout=e.stdout.decode('utf-8') if e.stdout else None,
                stderr=e.stderr.decode('utf-8') if e.stderr else None
            )

        except Exception as e:
            logger.error(f"Unexpected error during execution: {e}")
            return CodeExecutionResult(
                success=False,
                error_message=f"Unexpected error: {str(e)}",
                traceback=traceback.format_exc()
            )
