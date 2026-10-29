from pathlib import Path
import sqlite3

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
DB_PATH = DATA_DIR / "reporting.db"


SCHEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS clients (
    client_id INTEGER PRIMARY KEY,
    client_name TEXT NOT NULL UNIQUE,
    reporting_currency TEXT NOT NULL DEFAULT 'INR',
    is_active INTEGER NOT NULL DEFAULT 1
        CHECK (is_active IN (0, 1))
);

CREATE TABLE IF NOT EXISTS funds (
    fund_id INTEGER PRIMARY KEY,
    fund_name TEXT NOT NULL UNIQUE,
    benchmark_name TEXT NOT NULL,
    base_currency TEXT NOT NULL DEFAULT 'INR',
    is_active INTEGER NOT NULL DEFAULT 1
        CHECK (is_active IN (0, 1))
);

CREATE TABLE IF NOT EXISTS reporting_periods (
    period_id INTEGER PRIMARY KEY,
    period_start TEXT NOT NULL,
    period_end TEXT NOT NULL,
    frequency TEXT NOT NULL
        CHECK (frequency IN ('MONTHLY', 'QUARTERLY', 'ANNUAL')),
    status TEXT NOT NULL DEFAULT 'OPEN'
        CHECK (status IN ('OPEN', 'VALIDATED', 'RELEASED', 'BLOCKED')),
    CHECK (period_start <= period_end),
    UNIQUE (period_start, period_end, frequency)
);

CREATE TABLE IF NOT EXISTS holdings (
    holding_id INTEGER PRIMARY KEY,
    fund_id INTEGER NOT NULL,
    period_id INTEGER NOT NULL,
    security_id TEXT NOT NULL,
    security_name TEXT NOT NULL,
    sector TEXT NOT NULL,
    market_value REAL NOT NULL CHECK (market_value >= 0),
    weight REAL NOT NULL CHECK (weight >= 0 AND weight <= 1),
    FOREIGN KEY (fund_id) REFERENCES funds(fund_id),
    FOREIGN KEY (period_id) REFERENCES reporting_periods(period_id),
    UNIQUE (fund_id, period_id, security_id)
);

CREATE TABLE IF NOT EXISTS performance (
    performance_id INTEGER PRIMARY KEY,
    fund_id INTEGER NOT NULL,
    period_id INTEGER NOT NULL,
    fund_return REAL NOT NULL,
    benchmark_return REAL NOT NULL,
    FOREIGN KEY (fund_id) REFERENCES funds(fund_id),
    FOREIGN KEY (period_id) REFERENCES reporting_periods(period_id),
    UNIQUE (fund_id, period_id)
);

CREATE TABLE IF NOT EXISTS pipeline_runs (
    run_id INTEGER PRIMARY KEY,
    started_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    completed_at TEXT,
    status TEXT NOT NULL
        CHECK (status IN ('RUNNING', 'SUCCEEDED', 'FAILED', 'BLOCKED')),
    rows_processed INTEGER NOT NULL DEFAULT 0,
    error_message TEXT
);
"""


def initialize_database():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(DB_PATH) as connection:
        connection.executescript(SCHEMA)

        tables = connection.execute("""
            SELECT name
            FROM sqlite_master
            WHERE type = 'table'
              AND name NOT LIKE 'sqlite_%'
            ORDER BY name
        """).fetchall()

    print(f"Database: {DB_PATH}")
    print("Tables created:")
    for (table_name,) in tables:
        print(f" - {table_name}")


if __name__ == "__main__":
    initialize_database()
