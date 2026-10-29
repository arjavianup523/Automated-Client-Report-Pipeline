import sqlite3
from pathlib import Path
from database import initialize_database

initialize_database()

DB_PATH = Path(__file__).resolve().parents[1] / "data" / "reporting.db"

with sqlite3.connect(DB_PATH) as con:
    con.execute("PRAGMA foreign_keys = ON")

    con.execute("""
        INSERT OR IGNORE INTO clients
            (client_name, reporting_currency)
        VALUES (?, ?)
    """, ("Sample Institutional Client", "INR"))

    con.execute("""
        INSERT OR IGNORE INTO funds
            (fund_name, benchmark_name, base_currency)
        VALUES (?, ?, ?)
    """, ("Sample Balanced Fund", "Sample Composite Benchmark", "INR"))

    con.execute("""
        INSERT OR IGNORE INTO reporting_periods
            (period_start, period_end, frequency)
        VALUES (?, ?, ?)
    """, ("2026-01-01", "2026-03-31", "QUARTERLY"))

    client_id = con.execute("""
        SELECT client_id FROM clients
        WHERE client_name = ?
    """, ("Sample Institutional Client",)).fetchone()[0]

    fund_id = con.execute("""
        SELECT fund_id FROM funds
        WHERE fund_name = ?
    """, ("Sample Balanced Fund",)).fetchone()[0]

    period_id = con.execute("""
        SELECT period_id FROM reporting_periods
        WHERE period_start = ? AND period_end = ?
          AND frequency = ?
    """, ("2026-01-01", "2026-03-31", "QUARTERLY")).fetchone()[0]

    holdings = [
        (fund_id, period_id, "SEC001", "Sample Bank Ltd", "Financials", 3000000, 0.30),
        (fund_id, period_id, "SEC002", "Sample Technology Ltd", "Technology", 2500000, 0.25),
        (fund_id, period_id, "SEC003", "Sample Consumer Ltd", "Consumer", 2000000, 0.20),
        (fund_id, period_id, "SEC004", "Sample Energy Ltd", "Energy", 1500000, 0.15),
        (fund_id, period_id, "SEC005", "Sample Healthcare Ltd", "Healthcare", 1000000, 0.10),
    ]

    con.executemany("""
        INSERT OR IGNORE INTO holdings
            (fund_id, period_id, security_id, security_name,
             sector, market_value, weight)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, holdings)

    con.execute("""
        INSERT OR IGNORE INTO performance
            (fund_id, period_id, fund_return, benchmark_return)
        VALUES (?, ?, ?, ?)
    """, (fund_id, period_id, 0.042, 0.035))

    print(f"Client ID: {client_id}")
    print(f"Fund ID: {fund_id}")
    print(f"Reporting period ID: {period_id}")
    print("Holdings:", con.execute("SELECT COUNT(*) FROM holdings").fetchone()[0])
    print("Performance records:", con.execute("SELECT COUNT(*) FROM performance").fetchone()[0])
    print("Sample data loaded successfully.")

