"""Runtime configuration loaded from environment variables."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

_VALID_LOG_LEVELS = frozenset({"CRITICAL", "ERROR", "WARNING", "INFO", "DEBUG"})


@dataclass(frozen=True, slots=True)
class Settings:
    """Validated application settings."""

    environment: str = "development"
    log_level: str = "INFO"
    data_dir: Path = Path("data")

    @classmethod
    def from_environment(cls) -> Settings:
        """Build settings from PM Analyzer environment variables."""
        return cls(
            environment=os.getenv("PM_ANALYZER_ENVIRONMENT", "development").strip(),
            log_level=os.getenv("PM_ANALYZER_LOG_LEVEL", "INFO").strip().upper(),
            data_dir=Path(os.getenv("PM_ANALYZER_DATA_DIR", "data")).expanduser(),
        ).validated()

    def validated(self) -> Settings:
        """Return these settings after checking user-controlled values."""
        if not self.environment:
            raise ValueError("PM_ANALYZER_ENVIRONMENT must not be empty")
        if self.log_level not in _VALID_LOG_LEVELS:
            allowed = ", ".join(sorted(_VALID_LOG_LEVELS))
            raise ValueError(f"PM_ANALYZER_LOG_LEVEL must be one of: {allowed}")
        return self

