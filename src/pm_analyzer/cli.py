"""Command-line interface for PM Analyzer."""

from __future__ import annotations

import argparse
from collections.abc import Sequence

from pm_analyzer import __version__
from pm_analyzer.config import Settings
from pm_analyzer.logging import configure_logging


def build_parser() -> argparse.ArgumentParser:
    """Create the application argument parser."""
    parser = argparse.ArgumentParser(
        prog="pm-analyzer",
        description="Analyze SAP maintenance and fleet GPS data.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("version", help="Display the application version.")
    subparsers.add_parser("check-config", help="Validate and display runtime configuration.")
    return parser


def main(argv: Sequence[str] | None = None) -> None:
    """Run the requested PM Analyzer command."""
    args = build_parser().parse_args(argv)

    if args.command == "version":
        print(__version__)
        return

    settings = Settings.from_environment()
    configure_logging(settings.log_level)
    print(f"environment={settings.environment}")
    print(f"log_level={settings.log_level}")
    print(f"data_dir={settings.data_dir}")

