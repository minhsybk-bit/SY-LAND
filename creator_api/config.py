from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env.creator",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    environment: str = "development"
    public_api_url: str = "http://localhost:8000"
    allowed_origins: str = "http://localhost:5173"

    supabase_url: str
    supabase_anon_key: str
    supabase_service_role_key: str

    openai_api_key: str = ""
    openai_translation_model: str = "gpt-5.6-terra"
    gemini_api_key: str = ""
    gemini_translation_model: str = "gemini-3.1-flash-lite"

    redis_url: str = "redis://redis:6379/0"
    creator_storage_dir: Path = Path("/data/creator")
    download_signing_secret: str = Field(min_length=32)
    signed_url_ttl_seconds: int = Field(default=900, ge=60, le=86400)
    max_upload_bytes: int = Field(default=500 * 1024 * 1024, ge=1024)
    max_video_seconds: int = Field(default=900, ge=10, le=7200)
    yt_dlp_socket_timeout: int = Field(default=30, ge=5, le=120)
    output_retention_hours: int = Field(default=24, ge=1, le=720)

    @field_validator("public_api_url")
    @classmethod
    def normalize_api_url(cls, value: str) -> str:
        return value.strip().rstrip("/")

    @property
    def origins(self) -> list[str]:
        return [item.strip().rstrip("/") for item in self.allowed_origins.split(",") if item.strip()]

    def ensure_directories(self) -> None:
        self.creator_storage_dir.mkdir(parents=True, exist_ok=True)


@lru_cache
def get_settings() -> Settings:
    settings = Settings()
    settings.ensure_directories()
    return settings
