# Automated Client Report Pipeline

An independent portfolio project demonstrating an automated institutional client-reporting workflow using Python, SQLite, Excel, and PowerPoint.

> **Data notice:** All client, fund, holdings, benchmark, and performance data are fictional sample data. This is not an official Franklin Templeton system and is not affiliated with Franklin Templeton.

## Business problem

Institutional reporting teams need repeatable reporting processes, reliable data checks, consistent report outputs, and an audit trail. Manual preparation can introduce data-quality errors and inconsistencies between reports.

This project demonstrates a small, reproducible workflow that validates reporting data before creating deliverables and records whether each run succeeded, was blocked by validation, or failed.

## Workflow

1. Read holdings and performance data from a local SQLite database.
2. Validate required records, security identifiers, market values, returns, and portfolio weights.
3. Block report generation when portfolio validation fails.
4. Generate an Excel report with **Executive Summary**, **Holdings**, and **Validation Checks** worksheets.
5. Generate a four-slide PowerPoint report.
6. Ensure the Excel and PowerPoint output filenames share a timestamp.
7. Record pipeline status, timestamps, processed-row counts, and error details in the SQLite audit table.

## Technology

- Python
- SQLite
- `openpyxl`
- `python-pptx`
- `pytest`

## Repository structure

```text
Automated-Client-Report-Pipeline/
├── src/
│   ├── database.py
│   ├── generate_excel.py
│   ├── generate_powerpoint.py
│   ├── load_sample_data.py
│   ├── query_reporting_data.py
│   ├── run_pipeline.py
│   └── validation.py
├── tests/
│   ├── test_pipeline_failures.py
│   ├── test_pipeline_outputs.py
│   └── test_validation.py
├── data/       # Local SQLite database; generated locally and excluded from Git
├── outputs/    # Generated Excel and PowerPoint reports; excluded from Git
├── .gitignore
├── README.md
└── requirements.txt
```

## Setup and execution (Windows PowerShell)

### 1. Clone the repository

```powershell
git clone https://github.com/arjavianup523/Automated-Client-Report-Pipeline.git
cd Automated-Client-Report-Pipeline
```

### 2. Create and activate a virtual environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, use this for the current terminal session only:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### 4. Create the sample database

The local database is not stored in Git. Recreate it from the sample-data loader:

```powershell
python .\src\load_sample_data.py
```

Run this only when creating a fresh local database or intentionally rebuilding the sample data.

### 5. Run the pipeline

```powershell
python .\src\run_pipeline.py
```

The script should validate the data, generate one Excel report and one PowerPoint report in `outputs/`, and record the outcome in the SQLite audit table. The Excel and PowerPoint filenames for a successful run should share the same timestamp.

### 6. Run automated tests

```powershell
python -m pytest -v
```

The development verification run completed with **15 tests passed**. This is a record of the tested development environment, not a guarantee that every environment or production scenario is covered.

## Outputs

The pipeline creates:

- **Excel report** — Executive Summary, Holdings, and Validation Checks worksheets.
- **PowerPoint report** — four slides with summary metrics and a fictional-data disclaimer.
- **Audit record** — status, timestamps, rows processed, and error information in SQLite.

Generated reports and the local database are excluded from Git. Run the pipeline locally to create them.

## Test coverage

The current automated suite covers:

- Valid portfolio data.
- Incorrect portfolio weights.
- Negative market values.
- Duplicate securities.
- Missing performance records.
- Invalid data blocking report generation and recording an audit outcome.
- Report-generation failure recording a failed audit.
- Excel worksheet structure and database reconciliation for summary and holdings.
- Excel validation-check results.
- PowerPoint slide count, fictional-data disclaimer, and report figures.
- Matching Excel and PowerPoint report timestamps.
- Successful pipeline audit logging.

## Limitations and next improvements

This is a portfolio demonstration using fictional data and a fixed default client, fund, and reporting period. It is **not a production reporting platform**.

Potential next improvements include:

- Parameterized client, fund, and reporting-period selection.
- Stronger reconciliation controls and more edge-case tests.
- Atomic report publication and cleanup/recovery if only one output is generated.
- Dependency/version pinning and continuous integration.
- Structured logging, operational monitoring, and clearer report versioning.
- Access controls and protection appropriate for confidential client data.

## Project status

The core pipeline and report generation have been exercised locally, the current automated test suite passed **15/15 tests**, and the code changes were pushed to the `main` branch. Production readiness and every possible failure scenario have not been certified.
