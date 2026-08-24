"""应用基础框架测试，不读写真实薪资数据。"""

from pathlib import Path

import pytest

from backend.config import DEFAULT_JWT_SECRET, Settings, data_dir_from_mapping
from backend.main import api_health, app, health
from backend.database import engine_options


def test_health_reports_service_ready() -> None:
    assert health() == {
        "status": "ok",
        "service": "AutoPayroll-Pro API",
    }


def test_api_health_reports_runtime_environment() -> None:
    result = api_health()

    assert result["status"] == "ok"
    assert result["service"] == "AutoPayroll-Pro API"
    assert result["environment"] in {"development", "test", "production"}


def test_app_registers_core_routes() -> None:
    paths = set(app.openapi()["paths"])

    assert "/" in paths
    assert "/api/auth/login" in paths
    assert "/api/projects" in paths
    assert "/api/pipeline/{project_id}/base-db/load" in paths
    assert "/api/projects/{project_id}/files/batch-auto" in paths
    assert "/api/projects/{project_id}/files/{file_id}/analysis" in paths
    assert "/api/pipeline/{project_id}/integrate" in paths


def test_sqlite_engine_disables_thread_check() -> None:
    assert engine_options("sqlite:///local.db") == {
        "connect_args": {"check_same_thread": False},
        "echo": False,
    }


def test_postgresql_engine_does_not_receive_sqlite_options() -> None:
    assert engine_options("postgresql://localhost/payroll") == {
        "echo": False,
    }


def test_production_rejects_default_jwt_secret() -> None:
    settings = Settings.from_mapping(
        {
            "PAYROLL_ENV": "production",
            "PAYROLL_JWT_SECRET": DEFAULT_JWT_SECRET,
        }
    )

    with pytest.raises(RuntimeError, match="PAYROLL_JWT_SECRET"):
        settings.validate()


def test_cors_origins_are_parsed_from_environment() -> None:
    settings = Settings.from_mapping(
        {
            "PAYROLL_CORS_ORIGINS": "https://payroll.example.com, https://admin.example.com",
        }
    )

    assert settings.cors_origins == (
        "https://payroll.example.com",
        "https://admin.example.com",
    )


def test_default_cors_origins_cover_web_and_desktop_frontends() -> None:
    settings = Settings.from_mapping({})

    assert "http://127.0.0.1:3000" in settings.cors_origins
    assert "http://127.0.0.1:13000" in settings.cors_origins


def test_data_dir_uses_desktop_runtime_override(tmp_path: Path) -> None:
    expected = tmp_path / "外服账单系统" / "data"

    assert data_dir_from_mapping({"PAYROLL_DATA_DIR": str(expected)}) == expected
