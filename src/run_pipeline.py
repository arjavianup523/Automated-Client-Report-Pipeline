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


def report_timestamp():
    """Return a filename-safe UTC timestamp for this pipeline run."""
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def load_reporting_data(fund_id=1, period_id=1):
    """Load holdings and performance for the selected fund and period."""
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
    """Create an audit record and return its ID."""
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
    """Save the final pipeline outcome."""
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
            raise RuntimeError(f"Could not update audit record {run_id}.")


def align_report_timestamps(excel_result, powerpoint_result, timestamp):
    """Rename both generated reports to use the same run timestamp."""
    excel_path = Path(excel_result)
    powerpoint_path = Path(powerpoint_result)

    if not excel_path.is_file():
        raise FileNotFoundError(
            f"Excel generator did not create the expected file: {excel_path}"
        )

    if not powerpoint_path.is_file():
        raise FileNotFoundError(
            "PowerPoint generator did not create the expected file: "
            f"{powerpoint_path}"
        )

    excel_target = excel_path.with_name(f"client_report_{timestamp}.xlsx")
    powerpoint_target = powerpoint_path.with_name(
        f"client_report_{timestamp}.pptx"
    )

    # Avoid overwriting an existing report unexpectedly.
    for target, source in (
        (excel_target, excel_path),
        (powerpoint_target, powerpoint_path),
    ):
        if target.exists() and target.resolve() != source.resolve():
            raise FileExistsError(
                f"Cannot safely rename report because this file already exists: "
                f"{target}"
            )

    if excel_path.resolve() != excel_target.resolve():
        excel_path.replace(excel_target)

    if powerpoint_path.resolve() != powerpoint_target.resolve():
        powerpoint_path.replace(powerpoint_target)

    print(f"Shared report timestamp: {timestamp}")
    print(f"Excel report: {excel_target}")
    print(f"PowerPoint report: {powerpoint_target}")

    return excel_target, powerpoint_target


def run_pipeline(client_id=1, fund_id=1, period_id=1):
    """Validate data, generate reports, and audit the result."""
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
        print(f"Client ID: {client_id}")
        print(f"Fund ID: {fund_id}")
        print(f"Reporting period ID: {period_id}")
        print()

        print("STEP 1: Loading reporting data...")

        if fund_id == 1 and period_id == 1:
            holdings, performance_records = load_reporting_data()
        else:
            holdings, performance_records = load_reporting_data(
                fund_id=fund_id,
                period_id=period_id,
            )

        rows_processed = len(holdings)

        print(f"Holdings loaded: {len(holdings)}")
        print(f"Performance records loaded: {len(performance_records)}")
        print()

        print("STEP 2: Validating reporting data...")
        validation = validate_portfolio(holdings, performance_records)

        if not validation["passed"]:
            status = "BLOCKED"
            error_message = "; ".join(validation["errors"])

            print("VALIDATION FAILED - report generation blocked.")
            for error in validation["errors"]:
                print(f"  - {error}")

            return False

        print("Validation passed.")
        print()

        use_default_arguments = (
            client_id == 1 and fund_id == 1 and period_id == 1
        )

        # Capture one timestamp for the pair of reports.
        timestamp = report_timestamp()

        print("STEP 3: Generating Excel report...")
        if use_default_arguments:
            excel_result = generate_excel_report()
        else:
            excel_result = generate_excel_report(
                client_id=client_id,
                fund_id=fund_id,
                period_id=period_id,
            )

        print(f"Excel result: {excel_result}")
        print()

        print("STEP 4: Generating PowerPoint report...")
        if use_default_arguments:
            powerpoint_result = generate_powerpoint()
        else:
            powerpoint_result = generate_powerpoint(
                client_id=client_id,
                fund_id=fund_id,
                period_id=period_id,
            )

        print(f"PowerPoint result: {powerpoint_result}")
        print()

        align_report_timestamps(
            excel_result=excel_result,
            powerpoint_result=powerpoint_result,
            timestamp=timestamp,
        )

        status = "SUCCEEDED"
        print()
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
