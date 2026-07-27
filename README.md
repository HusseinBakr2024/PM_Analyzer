# PM Analyzer

PM Analyzer is a Python project for joining SAP maintenance exports with GPS fleet data and
producing traceable maintenance, utilization, and cost analytics.

The project is being implemented incrementally. The current milestone provides only the
project foundation: packaging, configuration, logging, a command-line entry point, and tests.
Excel ingestion and business analytics will be added in later, separately reviewed milestones.

## Requirements

- Python 3.11 or newer

## Local setup

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -e '.[dev]'
```

## Commands

Display the installed version:

```bash
pm-analyzer version
```

Check the resolved runtime configuration without reading any source data:

```bash
pm-analyzer check-config
```

The same commands can be run directly from the source tree:

```bash
PYTHONPATH=src python -m pm_analyzer version
```

## Development checks

```bash
pytest
ruff check .
mypy
```

## Data safety

Source Excel workbooks are treated as immutable inputs. Generated raw copies, quarantined
records, and reports are excluded from Git. Secrets and environment-specific configuration
belong in a local `.env` file and must not be committed.
