"""Runtime configuration loaded from environment variables."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path

_VALID_LOG_LEVELS = frozenset({"CRITICAL", "ERROR", "WARNING", "INFO", "DEBUG"})

_PREFERENCES_PATH = Path.home() / ".pm_analyzer" / "settings.json"


@dataclass(frozen=True, slots=True)
class UserPreferences:
    """User-owned maintenance policy; no analytical defaults are embedded in code."""

    interval_km: int | None = None
    idle_equivalent_km: float | None = None
    due_soon_percent: int | None = None

    @classmethod
    def load(cls, path: Path = _PREFERENCES_PATH) -> UserPreferences:
        if not path.exists():
            return cls()
        payload = json.loads(path.read_text(encoding="utf-8"))
        return cls(
            interval_km=int(payload["interval_km"]),
            idle_equivalent_km=float(payload["idle_equivalent_km"]),
            due_soon_percent=int(payload["due_soon_percent"]),
        ).validated()

    def save(self, path: Path = _PREFERENCES_PATH) -> None:
        values = self.validated()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(
                {
                    "interval_km": values.interval_km,
                    "idle_equivalent_km": values.idle_equivalent_km,
                    "due_soon_percent": values.due_soon_percent,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

    def validated(self) -> UserPreferences:
        if self.interval_km is None or self.interval_km <= 0:
            raise ValueError("يجب إدخال فترة صيانة أكبر من صفر")
        if self.idle_equivalent_km is None or self.idle_equivalent_km < 0:
            raise ValueError("يجب إدخال معامل تشغيل ساكن صالح")
        if self.due_soon_percent is None or not 0 < self.due_soon_percent < 100:
            raise ValueError("يجب أن تكون نسبة الصيانة القريبة بين 1 و99")
        return self
# Preventive-maintenance policy. These values are intentionally simple to edit.
PM_INTERVAL_KM = 10_000
DUE_SOON_PERCENT = 80
IDLE_HOUR_EQUIVALENT_KM = 30


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
