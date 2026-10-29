
# Automated Client Report Pipeline

## Overview

An independent portfolio project demonstrating an automated institutional client-reporting workflow using Python, SQLite, Excel, and PowerPoint.

The pipeline retrieves fund holdings and performance records, validates portfolio data, generates reporting files, and records execution outcomes in an audit table.

**Data notice:** All client, fund, holdings, benchmark, and performance data used in this project are fictional sample data. This is not an official Franklin Templeton system.

## Business Problem

Institutional reporting teams need repeatable reporting processes, reliable data checks, consistent outputs, and an audit trail. Manual preparation can introduce data-quality errors and inconsistencies between reports.

## Workflow

1. Retrieve holdings and performance data from SQLite.
2. Validate required records, security identifiers, market values, and portfolio weights.
3. Block report generation when validation fails.
4. Generate an Excel report.
5. Generate a four-slide PowerPoint report.
6. Record pipeline status, timestamps, and processing counts in SQLite.

## Technology Stack

- Python
- SQLite
- openpyxl
- python-pptx
- pytest

## Project Structure

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
│   ├── test_pipeline_outputs.py
│   └── test_validation.py
├── data/                 # Generated locally; excluded from Git
├── outputs/              # Generated reports; excluded from Git
├── .gitignore
└── README.md
```

## Setup and Execution

### 1. Clone the repository

```powershell
git clone https://github.com/arjavianup523/Automated-Client-Report-Pipeline.git
cd Automated-Client-Report-Pipeline
```

### 2. Create a virtual environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
python -m pip install openpyxl python-pptx pytest
```

### 4. Create the sample database

The local SQLite database is excluded from Git. Run the sample-data loader to recreate it:

```powershell
python .\src\load_sample_data.py
```

### 5. Run the reporting pipeline

```powershell
python .\src\run_pipeline.py
```

The pipeline validates the data before generating reports and records its execution outcome in the database.

### 6. Run the tests

```powershell
python -m pytest -v
```

Latest recorded test result during development: **9 passed**.

## Project Outputs

The pipeline generates:

- **Excel report:** Executive Summary, Holdings, and Validation Checks worksheets.
- **PowerPoint report:** Four-slide presentation.
- **Audit records:** Pipeline status, timestamps, processed-row counts, and error details.

Generated reports are saved in the local `outputs/` directory, which is excluded from Git.

## Validation and Testing

The automated tests cover:

- Valid portfolio data.
- Incorrect portfolio weights.
- Negative market values.
- Duplicate securities.
- Missing performance records.
- Excel worksheet structure and holdings rows.
- PowerPoint slide count.
- Matching Excel and PowerPoint report timestamps.
- Successful pipeline audit logging.

## Limitations and Future Improvements

The current implementation uses fictional sample data and a fixed default fund and reporting period. Further improvements could include configurable report parameters, stronger reconciliation controls, failure-path testing, access controls, and operational monitoring.

## Project Status

Core pipeline execution, data validation, Excel and PowerPoint generation, audit logging, and baseline automated tests are implemented. Further production-grade controls remain future enhancements.
