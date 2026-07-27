"""Central logging configuration."""

from __future__ import annotations

import logging
import time


def configure_logging(level: str) -> None:
    """Configure application logging with a consistent, UTC-friendly format."""
    logging.basicConfig(
        level=level,
        format="%(asctime)sZ %(levelname)s %(name)s %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S",
        force=True,
    )
    logging.Formatter.converter = time.gmtime
