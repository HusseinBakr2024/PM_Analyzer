import pytest

from pm_analyzer.cli import main


def test_version_command(capsys: pytest.CaptureFixture[str]) -> None:
    main(["version"])

    assert capsys.readouterr().out == "0.1.0\n"


def test_check_config_uses_defaults(
    capsys: pytest.CaptureFixture[str], monkeypatch: pytest.MonkeyPatch
) -> None:
    for name in (
        "PM_ANALYZER_ENVIRONMENT",
        "PM_ANALYZER_LOG_LEVEL",
        "PM_ANALYZER_DATA_DIR",
    ):
        monkeypatch.delenv(name, raising=False)

    main(["check-config"])

    assert capsys.readouterr().out.splitlines() == [
        "environment=development",
        "log_level=INFO",
        "data_dir=data",
    ]
