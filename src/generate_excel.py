
import sqlite3
from pathlib import Path
from datetime import datetime, timezone

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
DB_PATH = ROOT / "data" / "reporting.db"
OUTPUT_DIR = ROOT / "outputs"


def generate_excel_report():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(DB_PATH) as connection:
        connection.row_factory = sqlite3.Row

        summary = connection.execute("""
            SELECT
                c.client_name AS client,
                f.fund_name AS fund,
                f.benchmark_name AS benchmark,
                rp.period_start AS period_start,
                rp.period_end AS period_end,
                COUNT(h.holding_id) AS number_of_holdings,
                SUM(h.market_value) AS total_market_value,
                SUM(h.weight) AS total_weight,
                p.fund_return AS fund_return,
                p.benchmark_return AS benchmark_return,
                p.fund_return - p.benchmark_return AS active_return
            FROM clients c
            JOIN funds f ON f.fund_id = 1
            JOIN reporting_periods rp ON rp.period_id = 1
            JOIN holdings h
                ON h.fund_id = f.fund_id
               AND h.period_id = rp.period_id
            JOIN performance p
                ON p.fund_id = f.fund_id
               AND p.period_id = rp.period_id
            WHERE c.client_id = 1
            GROUP BY
                c.client_name, f.fund_name, f.benchmark_name,
                rp.period_start, rp.period_end,
                p.fund_return, p.benchmark_return
        """).fetchone()

        holdings = connection.execute("""
            SELECT security_id, security_name, sector,
                   market_value, weight
            FROM holdings
            WHERE fund_id = 1 AND period_id = 1
            ORDER BY market_value DESC
        """).fetchall()

    if summary is None or not holdings:
        raise ValueError("Required report data is missing.")

    total_weight = sum(row["weight"] for row in holdings)

    if abs(total_weight - 1.0) > 0.000001:
        raise ValueError("Report blocked: portfolio weights do not total 100%.")

    filename = (
        "client_report_"
        + datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        + ".xlsx"
    )
    output_path = OUTPUT_DIR / filename

    workbook = Workbook()

    # Sheet 1: Executive summary
    ws = workbook.active
    ws.title = "Executive Summary"

    ws.append(["Metric", "Value"])
    summary_rows = [
        ("Client", summary["client"]),
        ("Fund", summary["fund"]),
        ("Benchmark", summary["benchmark"]),
        ("Period Start", summary["period_start"]),
        ("Period End", summary["period_end"]),
        ("Number of Holdings", summary["number_of_holdings"]),
        ("Total Market Value (INR)", summary["total_market_value"]),
        ("Total Portfolio Weight", summary["total_weight"]),
        ("Fund Return", summary["fund_return"]),
        ("Benchmark Return", summary["benchmark_return"]),
        ("Active Return", summary["active_return"]),
    ]

    for row in summary_rows:
        ws.append(list(row))

    # Sheet 2: Holdings
    holdings_ws = workbook.create_sheet("Holdings")
    holdings_ws.append([
        "Security ID", "Security Name", "Sector",
        "Market Value (INR)", "Portfolio Weight"
    ])

    for holding in holdings:
        holdings_ws.append([
            holding["security_id"],
            holding["security_name"],
            holding["sector"],
            holding["market_value"],
            holding["weight"],
        ])

    # Sheet 3: Validation checks
    checks_ws = workbook.create_sheet("Validation Checks")
    checks_ws.append(["Control", "Result", "Details"])
    checks_ws.append([
        "Holdings present", "PASS",
        f"{len(holdings)} holdings loaded"
    ])
    checks_ws.append([
        "Portfolio weights", "PASS",
        f"Total weight: {total_weight:.2%}"
    ])
    checks_ws.append([
        "Performance record", "PASS",
        "One performance record loaded"
    ])

    # Format sheets
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                fill_type="solid", fgColor="1F4E78"
            )
            cell.alignment = Alignment(horizontal="center")

        for column_cells in sheet.columns:
            max_length = max(
                len(str(cell.value or "")) for cell in column_cells
            )
            letter = get_column_letter(column_cells[0].column)
            sheet.column_dimensions[letter].width = min(max_length + 3, 40)

    # Format financial values and percentages
    for row in range(2, ws.max_row + 1):
        metric = ws.cell(row, 1).value
        value_cell = ws.cell(row, 2)

        if metric == "Total Market Value (INR)":
            value_cell.number_format = '"₹"#,##0.00'
        elif metric in (
            "Total Portfolio Weight", "Fund Return",
            "Benchmark Return", "Active Return"
        ):
            value_cell.number_format = "0.00%"

    for row in range(2, holdings_ws.max_row + 1):
        holdings_ws.cell(row, 4).number_format = '"₹"#,##0.00'
        holdings_ws.cell(row, 5).number_format = "0.00%"

    workbook.save(output_path)

    print(f"Excel report created: {output_path}")
    print("Worksheets:", ", ".join(workbook.sheetnames))

    return output_path


if __name__ == "__main__":
    generate_excel_report()
