"""应用运行配置。

只读取操作系统环境变量，不读取或写入 .env 文件。
"""

from dataclasses import dataclass
from os import environ
from pathlib import Path
from typing import Mapping


DEFAULT_JWT_SECRET = "autopayroll-pro-dev-secret-change-in-prod-9f8a7b6c5d4e3f2a1"
DEFAULT_CORS_ORIGINS = (
    "http://localhost:3000",
    "http://127.0.0.1:3000",
    "http://localhost:3001",
    "http://127.0.0.1:3001",
    "http://localhost:3002",
    "http://127.0.0.1:3002",
    "http://localhost:3010",
    "http://127.0.0.1:3010",
    "http://localhost:13000",
    "http://127.0.0.1:13000",
)
DEFAULT_DATA_DIR = Path(__file__).resolve().parent / "data"


def data_dir_from_mapping(values: Mapping[str, str]) -> Path:
    """Resolve the writable runtime data directory for web and desktop builds."""
    configured_path = values.get("PAYROLL_DATA_DIR", "").strip()
    return Path(configured_path) if configured_path else DEFAULT_DATA_DIR


DEFAULT_DATABASE_PATH = data_dir_from_mapping(environ) / "payroll.db"
DEFAULT_DATABASE_URL = f"sqlite:///{DEFAULT_DATABASE_PATH.as_posix()}"


@dataclass(frozen=True)
class Settings:
    environment: str
    database_url: str
    jwt_secret: str
    cors_origins: tuple[str, ...]

    @classmethod
    def from_mapping(cls, values: Mapping[str, str]) -> "Settings":
        raw_origins = values.get("PAYROLL_CORS_ORIGINS", "")
        cors_origins = tuple(
            origin.strip()
            for origin in raw_origins.split(",")
            if origin.strip()
        ) or DEFAULT_CORS_ORIGINS

        return cls(
            environment=values.get("PAYROLL_ENV", "development").strip().lower(),
            database_url=values.get("PAYROLL_DATABASE_URL", DEFAULT_DATABASE_URL),
            jwt_secret=values.get("PAYROLL_JWT_SECRET", DEFAULT_JWT_SECRET),
            cors_origins=cors_origins,
        )

    def validate(self) -> None:
        """拒绝明显不安全的生产配置。"""
        if self.environment == "production" and self.jwt_secret == DEFAULT_JWT_SECRET:
            raise RuntimeError(
                "生产环境必须设置独立的 PAYROLL_JWT_SECRET，不能使用开发默认值"
            )


settings = Settings.from_mapping(environ)
