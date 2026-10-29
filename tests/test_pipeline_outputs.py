
"""End-to-end checks for the automated client reporting pipeline."""

import sqlite3
import sys
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DB_PATH = PROJECT_ROOT / "data" / "reporting.db"
OUTPUT_DIR = PROJECT_ROOT / "outputs"


def latest_file(extension):
    files = list(OUTPUT_DIR.glob(f"client_report_*.{extension}"))
    assert files, f"No .{extension} report found in {OUTPUT_DIR}"
    return max(files, key=lambda path: path.stat().st_mtime)


def test_excel_report_has_expected_worksheets():
    excel_path = latest_file("xlsx")

    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        assert workbook.sheetnames == [
            "Executive Summary",
            "Holdings",
            "Validation Checks",
        ]
        assert workbook["Holdings"].max_row >= 6
    finally:
        workbook.close()


def test_powerpoint_report_has_four_slides():
    powerpoint_path = latest_file("pptx")
    presentation = Presentation(powerpoint_path)

    assert len(presentation.slides) == 4


def test_excel_and_powerpoint_share_the_same_run_timestamp():
    excel_path = latest_file("xlsx")
    powerpoint_path = latest_file("pptx")

    assert excel_path.stem == powerpoint_path.stem, (
        "Excel and PowerPoint reports do not share the same run timestamp."
    )


def test_successful_pipeline_run_is_audited():
    with sqlite3.connect(DB_PATH) as connection:
        row = connection.execute(
            """
            SELECT status, rows_processed, error_message
            FROM pipeline_runs
            WHERE status = 'SUCCEEDED'
            ORDER BY run_id DESC
            LIMIT 1
            """
        ).fetchone()

    assert row is not None, "No successful pipeline run was recorded."
    assert row[0] == "SUCCEEDED"
    assert row[1] == 5
    assert row[2] is None
