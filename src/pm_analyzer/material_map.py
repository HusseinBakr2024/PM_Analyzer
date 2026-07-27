"""Configurable preventive-maintenance material knowledge map and scoring."""

from __future__ import annotations

import json
import unicodedata
from dataclasses import dataclass
from importlib.resources import files
from pathlib import Path
from typing import Any


@dataclass(frozen=True, slots=True)
class MaterialRule:
    name: str
    weight: int
    keywords: tuple[str, ...]
    enabled: bool = True


@dataclass(frozen=True, slots=True)
class PMMaterialMap:
    rules: tuple[MaterialRule, ...]
    pm_score_threshold: int
    excluded_phrases: tuple[str, ...]
    breakdown_indicators: tuple[str, ...]
    version: int

    @classmethod
    def load(cls, path: Path | None = None) -> PMMaterialMap:
        """Load the bundled map or a company-specific JSON override."""
        if path is None:
            source = files("pm_analyzer").joinpath("data/pm_material_map.json")
            payload = json.loads(source.read_text(encoding="utf-8"))
        else:
            payload = json.loads(path.read_text(encoding="utf-8"))
        return cls.from_dict(payload)

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> PMMaterialMap:
        rules = tuple(
            MaterialRule(
                name=str(item["name"]),
                weight=int(item["weight"]),
                keywords=tuple(normalize_phrase(keyword) for keyword in item["keywords"]),
                enabled=bool(item.get("enabled", True)),
            )
            for item in payload["categories"]
        )
        threshold = int(payload["pm_score_threshold"])
        if threshold <= 0 or not rules:
            raise ValueError("PM material map must contain rules and a positive threshold")
        return cls(
            rules=rules,
            pm_score_threshold=threshold,
            excluded_phrases=tuple(
                normalize_phrase(value) for value in payload.get("excluded_phrases", [])
            ),
            breakdown_indicators=tuple(
                normalize_phrase(value) for value in payload.get("breakdown_indicators", [])
            ),
            version=int(payload.get("version", 1)),
        )

    def match(self, description: str) -> tuple[str, int, str] | None:
        """Return category, weight, and matched keyword for a material description."""
        normalized = normalize_phrase(description)
        if any(phrase in normalized for phrase in self.excluded_phrases):
            return None
        for rule in self.rules:
            if not rule.enabled:
                continue
            keyword = next((word for word in rule.keywords if word in normalized), None)
            if keyword is not None:
                return rule.name, rule.weight, keyword
        return None

    def has_breakdown_indicator(self, descriptions: list[str]) -> bool:
        combined = " | ".join(normalize_phrase(value) for value in descriptions)
        return any(keyword in combined for keyword in self.breakdown_indicators)


def normalize_phrase(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value)).replace("\u200b", "").upper()
    text = "".join("-" if unicodedata.category(character) == "Pd" else character for character in text)
    return " ".join(text.split())
