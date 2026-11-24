from pathlib import Path
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict

BASE_DIR = Path(__file__).resolve().parents[1]

class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=BASE_DIR / '.env', extra='ignore')
    database_url: str = f"sqlite:///{BASE_DIR / 'data' / 'app.db'}"
    deployment_mode: str = 'local'
    upload_dir: Path = BASE_DIR / 'data' / 'uploads' / 'cv'
    model_dir: Path = BASE_DIR / 'models'
    ollama_base_url: str = 'http://127.0.0.1:11434'
    ollama_model: str = 'qwen3:4b'
    ollama_fallback_models: str = 'llama3.1:8b,llama3:8b,llama3:latest,qwen2.5:7b,mistral:7b,mistral:latest,gemma2:9b,llama3.2:3b,qwen3:8b,qwen3:4b,qwen3:1.7b,phi3:medium'
    answer_context_chars: int = Field(default=6000, ge=1000, le=20000)
    auto_fill_threshold: float = Field(default=0.90, ge=0, le=1)
    suggest_threshold: float = Field(default=0.60, ge=0, le=1)
    max_cv_bytes: int = 10 * 1024 * 1024

    @property
    def public_demo(self) -> bool:
        return self.deployment_mode.strip().lower() in {'public', 'public-demo', 'demo'}

settings = Settings()
settings.upload_dir.mkdir(parents=True, exist_ok=True)
settings.model_dir.mkdir(parents=True, exist_ok=True)
