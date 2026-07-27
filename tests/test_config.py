from pathlib import Path

import pytest

from pm_analyzer.config import Settings


def test_settings_are_loaded_from_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PM_ANALYZER_ENVIRONMENT", "test")
    monkeypatch.setenv("PM_ANALYZER_LOG_LEVEL", "debug")
    monkeypatch.setenv("PM_ANALYZER_DATA_DIR", "~/pm-analyzer-data")

    settings = Settings.from_environment()

    assert settings.environment == "test"
    assert settings.log_level == "DEBUG"
    assert settings.data_dir == Path("~/pm-analyzer-data").expanduser()


def test_invalid_log_level_is_rejected() -> None:
    with pytest.raises(ValueError, match="PM_ANALYZER_LOG_LEVEL"):
        Settings(log_level="VERBOSE").validated()


def test_empty_environment_is_rejected() -> None:
    with pytest.raises(ValueError, match="PM_ANALYZER_ENVIRONMENT"):
        Settings(environment="").validated()
