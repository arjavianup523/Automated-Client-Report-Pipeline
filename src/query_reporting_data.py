import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).resolve().parents[1] / "data" / "reporting.db"

SQL = """
SELECT
    c.client_name,
    f.fund_name,
    f.benchmark_name,
    rp.period_start,
    rp.period_end,
    COUNT(h.holding_id) AS number_of_holdings,
    SUM(h.market_value) AS total_market_value,
    SUM(h.weight) AS total_weight,
    p.fund_return,
    p.benchmark_return,
    p.fund_return - p.benchmark_return AS active_return
FROM clients AS c
JOIN funds AS f
    ON f.is_active IN (0, 1)
JOIN reporting_periods AS rp
    ON rp.period_start = ?
   AND rp.period_end = ?
JOIN holdings AS h
    ON h.fund_id = f.fund_id
   AND h.period_id = rp.period_id
JOIN performance AS p
    ON p.fund_id = f.fund_id
   AND p.period_id = rp.period_id
WHERE c.client_name = ?
  AND f.fund_name = ?
GROUP BY
    c.client_name,
    f.fund_name,
    f.benchmark_name,
    rp.period_start,
    rp.period_end,
    p.fund_return,
    p.benchmark_return;
"""

with sqlite3.connect(DB_PATH) as con:
    row = con.execute(
        SQL,
        (
            "2026-01-01",
            "2026-03-31",
            "Sample Institutional Client",
            "Sample Balanced Fund",
        ),
    ).fetchone()

if row is None:
    raise RuntimeError("No reporting data found for the selected period.")

columns = [
    "Client", "Fund", "Benchmark", "Period Start", "Period End",
    "Number of Holdings", "Total Market Value", "Total Weight",
    "Fund Return", "Benchmark Return", "Active Return",
]

print("\nQUARTERLY REPORTING SUMMARY")
print("-" * 60)

for name, value in zip(columns, row):
    if name in ("Fund Return", "Benchmark Return", "Active Return", "Total Weight"):
        print(f"{name:24}: {value:.2%}")
    elif name == "Total Market Value":
        print(f"{name:24}: ₹{value:,.2f}")
    else:
        print(f"{name:24}: {value}")

if abs(row[7] - 1.0) > 0.000001:
    print("\nVALIDATION FAILED: holding weights do not total 100%.")
else:
    print("\nVALIDATION PASSED: holding weights total 100%.")
