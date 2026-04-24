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
from src.task_chunker import validate_chunked_task_markdown
from src.incremental_runner import run_incremental_pipeline

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

        success, output_file, error_log = run_incremental_pipeline(
            task_content=task_content,
            style_content=style_content,
            language=args.lang,
            output_filename=output_filename,
            generator=generator,
            executor=executor,
            config=config,
            max_retries=config.max_retries,
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
