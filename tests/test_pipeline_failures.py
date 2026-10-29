
"""Failure-path tests using isolated temporary databases."""

import sqlite3
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import run_pipeline


def prepare_test_database(path):
    """Create an isolated pipeline audit database."""
    with sqlite3.connect(path) as connection:
        connection.execute(
            """
            CREATE TABLE pipeline_runs (
                run_id INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at TEXT NOT NULL,
                completed_at TEXT,
                status TEXT NOT NULL CHECK (
                    status IN ('RUNNING', 'SUCCEEDED', 'FAILED', 'BLOCKED')
                ),
                rows_processed INTEGER NOT NULL DEFAULT 0,
                error_message TEXT
            )
            """
        )


def read_latest_audit(path):
    with sqlite3.connect(path) as connection:
        return connection.execute(
            """
            SELECT status, rows_processed, error_message
            FROM pipeline_runs
            ORDER BY run_id DESC
            LIMIT 1
            """
        ).fetchone()


@pytest.fixture
def isolated_pipeline(tmp_path, monkeypatch):
    """Configure the pipeline to use temporary resources."""
    database_path = tmp_path / "reporting.db"
    output_path = tmp_path / "outputs"

    prepare_test_database(database_path)
    output_path.mkdir()

    monkeypatch.setattr(run_pipeline, "DB_PATH", database_path)
    monkeypatch.setattr(run_pipeline, "OUTPUT_DIR", output_path)

    return database_path, output_path


def test_invalid_data_blocks_reports_and_records_audit(
    isolated_pipeline, monkeypatch
):
    database_path, output_path = isolated_pipeline

    monkeypatch.setattr(
        run_pipeline,
        "load_reporting_data",
        lambda: ([], []),
    )

    monkeypatch.setattr(
        run_pipeline,
        "generate_excel_report",
        lambda: pytest.fail("Excel must not be generated for invalid data."),
    )

    monkeypatch.setattr(
        run_pipeline,
        "generate_powerpoint",
        lambda: pytest.fail("PowerPoint must not be generated for invalid data."),
    )

    result = run_pipeline.run_pipeline()

    assert result is False

    status, rows_processed, error_message = read_latest_audit(database_path)

    assert status == "BLOCKED"
    assert rows_processed == 0
    assert "No holdings found" in error_message
    assert not list(output_path.iterdir())


def test_report_generation_error_records_failed_audit(
    isolated_pipeline, monkeypatch
):
    database_path, _ = isolated_pipeline

    holdings = [
        {
            "security_id": "TEST001",
            "market_value": 100.0,
            "weight": 1.0,
        }
    ]

    performance_records = [
        {
            "fund_return": 0.05,
            "benchmark_return": 0.04,
        }
    ]

    monkeypatch.setattr(
        run_pipeline,
        "load_reporting_data",
        lambda: (holdings, performance_records),
    )

    def fail_excel_generation():
        raise RuntimeError("Simulated Excel generation failure")

    monkeypatch.setattr(
        run_pipeline,
        "generate_excel_report",
        fail_excel_generation,
    )

    result = run_pipeline.run_pipeline()

    assert result is False

    status, rows_processed, error_message = read_latest_audit(database_path)

    assert status == "FAILED"
    assert rows_processed == 1
    assert "Simulated Excel generation failure" in error_message
