import logging
import os
from pathlib import Path
from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

BASE_DIR = Path(__file__).resolve().parents[1]
log = logging.getLogger('autoform')


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=BASE_DIR / '.env', extra='ignore')

    # ---- Storage -------------------------------------------------------------------
    # Local default is SQLite. For hosting, set DATABASE_URL to a Postgres URL
    # (Neon, Supabase, Render Postgres...). `postgres://` URLs are accepted as-is.
    database_url: str = f"sqlite:///{BASE_DIR / 'data' / 'app.db'}"
    model_dir: Path = BASE_DIR / 'models'

    # ---- Answer generation (LLM) -----------------------------------------------------
    # 'ollama'  -> a local / self-hosted Ollama server (OLLAMA_BASE_URL)
    # 'openai'  -> any OpenAI-compatible API: Groq, OpenRouter, Together, Gemini, OpenAI...
    # 'auto'    -> 'openai' when LLM_API_KEY is set, otherwise 'ollama'
    llm_provider: str = 'auto'
    llm_base_url: str = 'https://api.groq.com/openai/v1'
    llm_api_key: str = ''
    llm_model: str = 'llama-3.1-8b-instant'
    llm_timeout_seconds: float = 60
    llm_reasoning_effort: str = ''
    llm_max_tokens_extra: int = Field(default=0, ge=0, le=2000)

    ollama_base_url: str = 'http://127.0.0.1:11434'
    ollama_model: str = 'qwen3:4b'
    ollama_fallback_models: str = 'llama3.1:8b,llama3:8b,llama3:latest,qwen2.5:7b,mistral:7b,mistral:latest,gemma2:9b,llama3.2:3b,qwen3:8b,qwen3:4b,qwen3:1.7b,phi3:medium'

    answer_context_chars: int = Field(default=6000, ge=1000, le=20000)
    auto_fill_threshold: float = Field(default=0.90, ge=0, le=1)
    suggest_threshold: float = Field(default=0.60, ge=0, le=1)
    max_cv_bytes: int = 10 * 1024 * 1024

    # ---- Abuse protection (per process, in memory) ----------------------------------
    register_limit_per_hour: int = 20        # new anonymous accounts per IP
    generate_limit_per_hour: int = 30        # LLM drafts per user
    cv_upload_limit_per_hour: int = 20       # CV parses per user
    max_fields_per_scan: int = 400

    # Extra browser origins allowed by CORS (comma separated). Extension origins
    # (chrome-extension://, moz-extension://) are always allowed.
    cors_origins: str = ''

    @field_validator('database_url')
    @classmethod
    def _normalise_db_url(cls, value: str) -> str:
        # Hosting providers hand out postgres:// or postgresql:// URLs; SQLAlchemy
        # needs the driver spelled out to use psycopg 3.
        for prefix in ('postgres://', 'postgresql://'):
            if value.startswith(prefix):
                return 'postgresql+psycopg://' + value[len(prefix):]
        return value

    @property
    def provider(self) -> str:
        choice = self.llm_provider.strip().lower()
        if choice in ('ollama', 'openai'):
            return choice
        return 'openai' if self.llm_api_key else 'ollama'

    @property
    def is_sqlite(self) -> bool:
        return self.database_url.startswith('sqlite')


settings = Settings()
settings.model_dir.mkdir(parents=True, exist_ok=True)
if settings.is_sqlite:
    Path(settings.database_url.removeprefix('sqlite:///')).parent.mkdir(parents=True, exist_ok=True)
    if os.environ.get('RENDER') or os.environ.get('RAILWAY_ENVIRONMENT') or os.environ.get('FLY_APP_NAME'):
        log.warning('DATABASE_URL is not set: using SQLite on an ephemeral disk. '
                    'All user data will be lost on every restart or deploy. Set DATABASE_URL to a Postgres database.')
