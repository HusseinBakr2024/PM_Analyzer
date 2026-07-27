# PM Analyzer

PM Analyzer is a Python project for joining SAP maintenance exports with GPS fleet data and
producing traceable maintenance, utilization, and cost analytics.

The application provides an Arabic desktop workflow for selecting one to seven GPS workbooks,
the SAP maintenance-order workbook, and the SAP material-document workbook. It calculates
preventive-maintenance status and exports a ten-sheet, right-to-left Excel report.

## Requirements

- Python 3.11 or newer

## Local setup

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -e '.[dev]'
```

## Commands

### Arabic desktop application (recommended)

On Windows, double-click `تشغيل_PM_Analyzer.bat`. Alternatively run:

```bash
PYTHONPATH=src python -m pm_analyzer.gui
```

Select the files in this order:

1. One to seven GPS files.
2. `Maintenance Notifications and Orders` as the authoritative asset/order source.
3. `Material Documents` for preventive-maintenance materials and posting dates.
4. Choose the output location and press the report button.

The maintenance interval and due-soon percentage can be edited in the application. Their
defaults, together with the 30-km idle-hour equivalent, are defined in `config.py`. The main
window exposes all three policy inputs so the user can change them before every report:

- preventive-maintenance interval in kilometers;
- equivalent kilometers for one idle hour;
- due-soon percentage.

### Command-line report

```bash
pm-analyzer analyze \
  --maintenance "Maintenance Notifications and Orders.xlsx" \
  --materials "Material Documents.xlsx" \
  --gps GPS-1.xlsx GPS-2.xlsx \
  --interval-km 10000 \
  --idle-equivalent-km 30 \
  --due-soon-percent 80 \
  --output "preventive-maintenance-report.xlsx"
```

The generated workbook contains the dashboard, due analysis, status-specific lists, latest
maintenance, merged GPS detail, preventive-material detail, and data-quality findings.

### Diagnostic commands

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
