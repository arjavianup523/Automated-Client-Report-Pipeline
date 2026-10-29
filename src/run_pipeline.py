
"""Automated institutional client reporting pipeline."""

import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_PATH = PROJECT_ROOT / "data" / "reporting.db"
OUTPUT_DIR = PROJECT_ROOT / "outputs"

SRC_DIR = Path(__file__).resolve().parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from validation import validate_portfolio
from generate_excel import generate_excel_report
from generate_powerpoint import generate_powerpoint


def utc_now():
    """Return the current UTC timestamp."""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def load_reporting_data(fund_id=1, period_id=1):
    """Load holdings and performance using the database's actual schema."""
    with sqlite3.connect(DB_PATH) as connection:
        connection.row_factory = sqlite3.Row

        holdings = [
            dict(row)
            for row in connection.execute(
                """
                SELECT holding_id, fund_id, period_id, security_id,
                       security_name, sector, market_value, weight
                FROM holdings
                WHERE fund_id = ? AND period_id = ?
                ORDER BY security_id
                """,
                (fund_id, period_id),
            ).fetchall()
        ]

        performance_records = [
            dict(row)
            for row in connection.execute(
                """
                SELECT performance_id, fund_id, period_id,
                       fund_return, benchmark_return
                FROM performance
                WHERE fund_id = ? AND period_id = ?
                """,
                (fund_id, period_id),
            ).fetchall()
        ]

    return holdings, performance_records


def create_audit_run():
    """Create an audit record and return SQLite's integer ID."""
    with sqlite3.connect(DB_PATH) as connection:
        cursor = connection.execute(
            """
            INSERT INTO pipeline_runs
                (started_at, status, rows_processed, error_message)
            VALUES (?, ?, ?, ?)
            """,
            (utc_now(), "RUNNING", 0, None),
        )
        return cursor.lastrowid


def update_audit_run(run_id, status, rows_processed, error_message=None):
    """Save the final pipeline outcome to the audit table."""
    allowed_statuses = {"RUNNING", "SUCCEEDED", "FAILED", "BLOCKED"}

    if status not in allowed_statuses:
        raise ValueError(f"Unsupported audit status: {status}")

    with sqlite3.connect(DB_PATH) as connection:
        cursor = connection.execute(
            """
            UPDATE pipeline_runs
            SET completed_at = ?,
                status = ?,
                rows_processed = ?,
                error_message = ?
            WHERE run_id = ?
            """,
            (utc_now(), status, rows_processed, error_message, run_id),
        )

        if cursor.rowcount != 1:
            raise RuntimeError(
                f"Could not update audit record {run_id}."
            )


def run_pipeline():
    """Validate data, generate reports, and record the outcome."""
    print("=" * 60)
    print("AUTOMATED CLIENT REPORT PIPELINE")
    print("=" * 60)

    if not DB_PATH.exists():
        print(f"ERROR: Database not found: {DB_PATH}")
        return False

    run_id = None
    rows_processed = 0
    status = "FAILED"
    error_message = None

    try:
        run_id = create_audit_run()
        print(f"Run ID: {run_id}")
        print(f"Database: {DB_PATH}")
        print()

        print("STEP 1: Loading reporting data...")
        holdings, performance_records = load_reporting_data()
        rows_processed = len(holdings)

        print(f"Holdings loaded: {len(holdings)}")
        print(f"Performance records loaded: {len(performance_records)}")
        print()

        print("STEP 2: Validating reporting data...")
        validation = validate_portfolio(holdings, performance_records)

        if not validation["passed"]:
            status = "BLOCKED"
            error_message = "; ".join(validation["errors"])

            print("VALIDATION FAILED — report generation blocked.")
            for error in validation["errors"]:
                print(f"  - {error}")

            return False

        print("Validation passed.")
        print()

        print("STEP 3: Generating Excel report...")
        excel_result = generate_excel_report()
        print(f"Excel result: {excel_result}")
        print()

        print("STEP 4: Generating PowerPoint report...")
        powerpoint_result = generate_powerpoint()
        print(f"PowerPoint result: {powerpoint_result}")
        print()

        status = "SUCCEEDED"
        print("PIPELINE COMPLETED SUCCESSFULLY.")
        print(f"Holdings processed: {rows_processed}")
        print(f"Reports directory: {OUTPUT_DIR}")
        return True

    except Exception as exc:
        status = "FAILED"
        error_message = f"{type(exc).__name__}: {exc}"
        print()
        print("PIPELINE FAILED.")
        print(error_message)
        return False

    finally:
        if run_id is not None:
            try:
                update_audit_run(
                    run_id=run_id,
                    status=status,
                    rows_processed=rows_processed,
                    error_message=error_message,
                )
                print(f"\nAudit record {run_id}: {status}")
            except Exception as audit_exc:
                print(
                    "AUDIT UPDATE FAILED: "
                    f"{type(audit_exc).__name__}: {audit_exc}"
                )


if __name__ == "__main__":
    success = run_pipeline()
    sys.exit(0 if success else 1)
