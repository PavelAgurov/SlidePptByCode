"""Configuration management using pydantic-settings."""

from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings loaded from .env and environment variables."""

    model_config = SettingsConfigDict(
        env_file='.env',
        env_file_encoding='utf-8',
        case_sensitive=False,
        extra='ignore'
    )

    # API Configuration
    openrouter_api_key: str
    model_id: str = "gpt-4.1"
    base_url: str = "https://openrouter.ai/api/v1"
    temperature: float = 0
    # Cap completion budget so providers (e.g. OpenRouter) do not default to very large
    # max_tokens, which can fail with 402 when credits only cover a smaller reservation.
    llm_max_tokens: int = 16_384

    # Execution Configuration
    max_retries: int = 3
    execution_timeout: int = 60

    # Directory Configuration
    generated_dir: Path = Path(".generated")
    output_dir: Path = Path(".output")

    # Logging
    log_level: str = "INFO"

    def create_directories(self) -> None:
        """Create necessary directories if they don't exist."""
        self.generated_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)


def load_settings() -> Settings:
    """Load and return application settings."""
    # Fields such as ``openrouter_api_key`` are populated from env / ``.env`` via pydantic-settings.
    return Settings()  # type: ignore[call-arg]
