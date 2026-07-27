"""Command-line interface for PM Analyzer."""

from __future__ import annotations

import argparse
from collections.abc import Sequence
from pathlib import Path

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
    analyze_parser = subparsers.add_parser("analyze", help="Create the preventive-maintenance report.")
    analyze_parser.add_argument("--maintenance", type=Path, required=True)
    analyze_parser.add_argument("--materials", type=Path, required=True)
    analyze_parser.add_argument("--gps", type=Path, nargs="+", required=True)
    analyze_parser.add_argument("--output", type=Path, required=True)
    return parser


def main(argv: Sequence[str] | None = None) -> None:
    """Run the requested PM Analyzer command."""
    args = build_parser().parse_args(argv)

    if args.command == "version":
        print(__version__)
        return

    if args.command == "analyze":
        from pm_analyzer.engine import analyze, export_report

        result = analyze(args.maintenance, args.materials, args.gps)
        export_report(result, args.output)
        print(f"report={args.output}")
        print(f"assets={len(result.analysis)}")
        return

    settings = Settings.from_environment()
    configure_logging(settings.log_level)
    print(f"environment={settings.environment}")
    print(f"log_level={settings.log_level}")
    print(f"data_dir={settings.data_dir}")
