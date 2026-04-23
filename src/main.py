"""Main CLI entry point for presentation generator."""

import argparse
import logging
import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.config import load_settings
from src.llm_client import LLMClient
from src.code_generator import CodeGenerator
from src.code_executor import CodeExecutor
from src.validator import validate_presentation
from src.task_chunker import split_into_chunks, validate_chunked_task_markdown

logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False) -> None:
    """Configure logging."""
    level = logging.DEBUG if verbose else logging.INFO

    # Create logs directory
    logs_dir = Path('.logs')
    logs_dir.mkdir(exist_ok=True)

    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(logs_dir / 'agent.log', encoding='utf-8')
        ]
    )


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Generate PowerPoint presentations using LLM-generated code"
    )
    parser.add_argument(
        "input_file",
        help="Path to task specification file (markdown or text)"
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=None,
        help="Maximum number of retry attempts (default: from config)"
    )
    parser.add_argument(
        "--model",
        default=None,
        help=(
            "LLM model to use (default: from config). "
            "On PowerShell, quote values with hyphens, e.g. --model \"gpt-4.1\", or use --model=gpt-4.1"
        ),
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=None,
        help="Timeout for code execution in seconds (default: from config)"
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable verbose logging"
    )
    parser.add_argument(
        "--style",
        default=None,
        help="Path to style guidelines file (markdown or text)"
    )
    parser.add_argument(
        "--lang",
        default=None,
        help="Language for the presentation content (e.g., 'Russian', 'English', 'Spanish')"
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output filename for the presentation (e.g., 'my_presentation.pptx' or 'my_presentation'). If directory not specified, uses '.output/'"
    )
    return parser.parse_args()


def read_task_file(file_path: str) -> str:
    """Read task content from file."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Task file not found: {file_path}")

    content = path.read_text(encoding='utf-8')
    logger.info(f"Read task file: {file_path} ({len(content)} characters)")
    return content


def normalize_output_filename(output_arg: str | None) -> str | None:
    """
    Normalize output filename argument.

    - If None, returns None (let LLM generate name)
    - If no extension, adds .pptx
    - If no directory, prepends .output/

    Args:
        output_arg: User-provided output argument

    Returns:
        Normalized output path or None
    """
    if output_arg is None:
        return None

    path = Path(output_arg)

    # Add .pptx extension if missing
    if path.suffix.lower() != '.pptx':
        path = Path(str(path) + '.pptx')

    # If no directory specified, use .output/
    if not path.parent or path.parent == Path('.'):
        path = Path('.output') / path.name

    result = str(path)
    logger.info(f"Normalized output filename: {result}")
    return result


def execute_with_retry(
    task_content: str,
    style_content: str | None,
    language: str | None,
    output_filename: str | None,
    generator: CodeGenerator,
    executor: CodeExecutor,
    max_retries: int,
    initial_code: str | None = None,
    initial_explanation: str | None = None,
) -> tuple[bool, Path | None, list[str]]:
    """
    Execute generation and validation with retry loop.

    Args:
        task_content: Task specification content
        style_content: Optional style guidelines content
        language: Optional language for presentation content
        output_filename: Optional output filename for the presentation
        generator: Code generator instance
        executor: Code executor instance
        max_retries: Maximum retry attempts
        initial_code: If set, skip initial generation and use this merged script (chunked mode)
        initial_explanation: Optional short note stored in conversation for retries

    Returns:
        (success, output_file_path, error_log)
    """
    error_log = []
    current_code = None

    # Initial generation (or pre-built merged code from chunked pipeline)
    try:
        if initial_code is not None:
            logger.info("=== Using merged chunked code (skipping single-pass generate) ===")
            current_code = initial_code
            generator.reset_conversation_with_merged_code(
                initial_code, initial_explanation or "Merged chunked output"
            )
        else:
            logger.info("=== Generating initial code ===")
            generated = generator.generate_initial_code(
                task_content, style_content, language, output_filename
            )
            current_code = generated.code
    except Exception as e:
        error_log.append(f"Initial generation failed: {str(e)}")
        return False, None, error_log

    # Retry loop
    for attempt in range(max_retries):
        logger.info(f"=== Attempt {attempt + 1}/{max_retries} ===")

        # Save and execute code
        try:
            code_file = executor.save_code(current_code, f"attempt_{attempt + 1}.py")
            result = executor.execute(code_file)

            if result.success and result.output_file:
                # Execution succeeded, validate presentation
                logger.info("Execution succeeded, validating presentation")
                validation = validate_presentation(result.output_file, task_content)

                if validation.is_valid:
                    logger.info("✓ Validation passed!")
                    return True, result.output_file, error_log
                else:
                    # Validation failed
                    error_msg = f"Validation failed: {', '.join(validation.issues)}"
                    error_log.append(f"Attempt {attempt + 1}: {error_msg}")
                    logger.warning(error_msg)

                    if attempt < max_retries - 1:
                        # Request fix
                        logger.info("Requesting fix for validation issues")
                        generated = generator.fix_code_after_validation(
                            current_code,
                            validation,
                            task_content
                        )
                        current_code = generated.code
            else:
                # Execution failed
                error_msg = result.error_message or "Unknown execution error"
                error_log.append(f"Attempt {attempt + 1}: {error_msg}")
                logger.error(error_msg)

                if attempt < max_retries - 1:
                    # Request fix
                    logger.info("Requesting fix for execution error")
                    generated = generator.fix_code_after_error(current_code, result)
                    current_code = generated.code

        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            error_log.append(f"Attempt {attempt + 1}: {error_msg}")
            logger.error(error_msg)

            if attempt >= max_retries - 1:
                break

    return False, None, error_log


def main() -> int:
    """Main entry point."""
    args = parse_args()

    # Setup logging
    setup_logging(args.verbose)

    logger.info("=== Starting presentation generator ===")

    try:
        # Load configuration
        config = load_settings()

        # Override config with CLI args
        if args.max_retries is not None:
            config.max_retries = args.max_retries
        if args.model is not None:
            config.model_id = args.model
        if args.timeout is not None:
            config.execution_timeout = args.timeout

        # Create directories
        config.create_directories()
        logger.info(f"Using model: {config.model_id}, max retries: {config.max_retries}")

        # Read task file
        task_content = read_task_file(args.input_file)
        validate_chunked_task_markdown(task_content)

        # Read style file if provided
        style_content = None
        if args.style:
            style_path = Path(args.style)
            if not style_path.exists():
                raise FileNotFoundError(f"Style file not found: {args.style}")
            style_content = style_path.read_text(encoding='utf-8')
            logger.info(f"Read style file: {args.style} ({len(style_content)} characters)")

        # Initialize components
        llm_client = LLMClient(config)
        generator = CodeGenerator(llm_client)
        executor = CodeExecutor(config)

        # Normalize output filename
        output_filename = normalize_output_filename(args.output)

        chunks = split_into_chunks(task_content)
        logger.info(
            "Chunked generation: %d part(s) (1 title block + %d H2 slide(s))",
            len(chunks),
            len(chunks) - 1,
        )
        chunked = generator.generate_chunked_code(
            chunks,
            task_content,
            style_content,
            args.lang,
            output_filename,
        )
        initial_code = chunked.code
        initial_explanation = chunked.explanation

        # Execute with retry
        success, output_file, error_log = execute_with_retry(
            task_content,
            style_content,
            args.lang,
            output_filename,
            generator,
            executor,
            config.max_retries,
            initial_code=initial_code,
            initial_explanation=initial_explanation,
        )

        # Report results
        print("\n" + "=" * 60)
        if success:
            print(f"SUCCESS: Presentation created at {output_file}")
            print("=" * 60)
            return 0
        else:
            print("FAILED: Could not generate valid presentation")
            print("\nError log:")
            for error in error_log:
                print(f"  - {error}")
            print("=" * 60)
            return 1

    except Exception as e:
        logger.exception("Fatal error")
        print(f"\nFATAL ERROR: {str(e)}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
