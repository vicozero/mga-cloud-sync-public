from functools import lru_cache
from typing import List

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    app_name: str = "Disponibilidad API"
    app_env: str = "development"
    app_host: str = "0.0.0.0"
    app_port: int = 8000
    secret_key: str = "change-me"
    access_token_expire_minutes: int = 480
    api_key: str = "MTTOPROV"
    default_admin_username: str = "admin"
    default_admin_password: str = "admin123"
    default_admin_name: str = "Administrador"
    database_url: str = "postgresql+psycopg://postgres:postgres@localhost:5432/disponibilidad"
    cors_origins: str = ""
    upload_dir: str = ""
    reports_dir: str = ""
    max_upload_mb: int = 8
    supabase_url: str = ""
    supabase_service_role_key: str = ""
    supabase_bucket: str = "disponibilidad"

    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    @property
    def cors_origin_list(self) -> List[str]:
        if not self.cors_origins.strip():
            return []
        return [item.strip() for item in self.cors_origins.split(",") if item.strip()]


@lru_cache
def get_settings() -> Settings:
    return Settings()
