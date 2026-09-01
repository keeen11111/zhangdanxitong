from backend.config import DEFAULT_CORS_ORIGINS, Settings


def test_local_dev_cors_allows_supported_verification_ports() -> None:
    assert "http://127.0.0.1:3001" in DEFAULT_CORS_ORIGINS
    assert "http://localhost:3002" in DEFAULT_CORS_ORIGINS
    assert "http://127.0.0.1:3010" in DEFAULT_CORS_ORIGINS


def test_explicit_cors_configuration_still_replaces_defaults() -> None:
    settings = Settings.from_mapping({"PAYROLL_CORS_ORIGINS": "http://example.test"})

    assert settings.cors_origins == ("http://example.test",)
