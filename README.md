
# Automated Client Report Pipeline

## Overview

A portfolio project demonstrating an automated institutional client-reporting workflow using SQLite, Python, Excel, and PowerPoint.

The pipeline retrieves fund holdings and performance records, validates portfolio data, generates reporting files, and records execution outcomes in an audit table.

**Data notice:** All client, fund, holdings, benchmark, and performance data used in this project are fictional sample data. This is an independent portfolio project, not an official Franklin Templeton system.

## Business Problem

Institutional reporting teams need repeatable reporting processes, reliable data checks, consistent outputs, and an audit trail. Manual preparation can introduce data-quality errors and inconsistencies between reports.

## Workflow

1. Retrieve holdings and performance data from SQLite.
2. Validate required records, security identifiers, market values, and portfolio weights.
3. Block report generation when validation fails.
4. Generate an Excel report.
5. Generate a four-slide PowerPoint report.
6. Record pipeline status and processing counts in SQLite.

## Technology Stack

- Python
- SQLite
- openpyxl
- python-pptx
- pytest

## Project Outputs

The pipeline generates:

- Excel report: Executive Summary, Holdings, and Validation Checks worksheets.
- PowerPoint report: four-slide presentation.
- Audit records: pipeline status, timestamps, processed-row counts, and error details.

## Validation and Testing

The automated tests cover valid portfolios, incorrect portfolio weights, negative market values, duplicate securities, missing performance records, Excel report structure, PowerPoint slide count, report timestamp consistency, and successful-run audit logging.

Latest recorded test result: **9 passed**.

## Run the Pipeline

From the project root in PowerShell:

```powershell
.\.venv\Scripts\python.exe .\src\run_pipeline.py
```

Run the test suite:

```powershell
.\.venv\Scripts\python.exe -m pytest -v
```

## Limitations

The current implementation uses sample data and a fixed default fund and reporting period. Production deployment would require configurable report parameters, stronger integration and reconciliation controls, access controls, operational monitoring, and testing of failure scenarios.

## Project Status

Core pipeline and report generation implemented. Automated baseline tests passing. Further production-grade controls remain future enhancements.
