"""Generate an institutional portfolio PowerPoint report."""

import sqlite3
from pathlib import Path
from datetime import datetime, timezone

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


ROOT = Path(__file__).resolve().parents[1]
DB_PATH = ROOT / "data" / "reporting.db"
OUTPUT_DIR = ROOT / "outputs"

NAVY = RGBColor(31, 78, 121)
BLUE = RGBColor(68, 114, 196)
LIGHT_BLUE = RGBColor(221, 235, 247)
DARK = RGBColor(45, 45, 45)
WHITE = RGBColor(255, 255, 255)
GREEN = RGBColor(0, 120, 80)
RED = RGBColor(180, 50, 50)


def fetch_report_data(client_id=1, fund_id=1, period_id=1):
    """Fetch and validate report data for the selected records."""
    if not DB_PATH.exists():
        raise FileNotFoundError(
            f"Reporting database not found: {DB_PATH}"
        )

    with sqlite3.connect(DB_PATH) as connection:
        connection.row_factory = sqlite3.Row

        summary = connection.execute(
            """
            SELECT
                c.client_name AS client,
                f.fund_name AS fund,
                f.benchmark_name AS benchmark,
                f.base_currency AS currency,
                rp.period_start,
                rp.period_end,
                COUNT(h.holding_id) AS number_of_holdings,
                SUM(h.market_value) AS total_market_value,
                SUM(h.weight) AS total_weight,
                p.fund_return,
                p.benchmark_return,
                p.fund_return - p.benchmark_return AS active_return
            FROM clients c
            JOIN funds f ON f.fund_id = ?
            JOIN reporting_periods rp ON rp.period_id = ?
            JOIN holdings h
                ON h.fund_id = f.fund_id
               AND h.period_id = rp.period_id
            JOIN performance p
                ON p.fund_id = f.fund_id
               AND p.period_id = rp.period_id
            WHERE c.client_id = ?
            GROUP BY
                c.client_name,
                f.fund_name,
                f.benchmark_name,
                f.base_currency,
                rp.period_start,
                rp.period_end,
                p.fund_return,
                p.benchmark_return
            """,
            (fund_id, period_id, client_id),
        ).fetchone()

        holdings = connection.execute(
            """
            SELECT
                security_id,
                security_name,
                sector,
                market_value,
                weight
            FROM holdings
            WHERE fund_id = ? AND period_id = ?
            ORDER BY market_value DESC
            """,
            (fund_id, period_id),
        ).fetchall()

        performance_count = connection.execute(
            """
            SELECT COUNT(*)
            FROM performance
            WHERE fund_id = ? AND period_id = ?
            """,
            (fund_id, period_id),
        ).fetchone()[0]

    if summary is None or not holdings:
        raise ValueError(
            "Report generation blocked: required data is missing."
        )

    if performance_count != 1:
        raise ValueError(
            "Report generation blocked: expected exactly one performance "
            "record for the selected fund and period."
        )

    for holding in holdings:
        market_value = holding["market_value"]
        weight = holding["weight"]

        if (
            market_value is None
            or isinstance(market_value, bool)
            or market_value < 0
        ):
            raise ValueError(
                f"Invalid market value for {holding['security_id']}."
            )

        if (
            weight is None
            or not isinstance(weight, (int, float))
            or isinstance(weight, bool)
            or not 0 <= weight <= 1
        ):
            raise ValueError(
                f"Invalid portfolio weight for {holding['security_id']}."
            )

    total_weight = sum(row["weight"] for row in holdings)

    if abs(total_weight - 1.0) > 0.000001:
        raise ValueError(
            f"Report generation blocked: portfolio weights total "
            f"{total_weight:.2%}, not 100%."
        )

    if summary["fund_return"] is None or summary["benchmark_return"] is None:
        raise ValueError("Report generation blocked: performance data is missing.")

    return dict(summary), [dict(row) for row in holdings]


def set_cell_text(cell, text, bold=False, color=DARK, size=12):
    """Apply consistent formatting to a PowerPoint table cell."""
    cell.text = str(text)
    cell.vertical_anchor = 1

    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Aptos"
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = color


def add_footer(slide, slide_number):
    """Add a fictional-data disclaimer and slide number."""
    box = slide.shapes.add_textbox(
        Inches(0.5),
        Inches(7.12),
        Inches(12.2),
        Inches(0.2),
    )

    paragraph = box.text_frame.paragraphs[0]
    paragraph.text = (
        "CONFIDENTIALITY: Illustrative sample report — fictional data"
        f"                                                   {slide_number}"
    )
    paragraph.font.size = Pt(8)
    paragraph.font.color.rgb = RGBColor(100, 100, 100)


def add_title(slide, title, subtitle=None):
    """Add a slide title and optional subtitle."""
    box = slide.shapes.add_textbox(
        Inches(0.6),
        Inches(0.35),
        Inches(12),
        Inches(0.6),
    )

    paragraph = box.text_frame.paragraphs[0]
    paragraph.text = title
    paragraph.font.name = "Aptos Display"
    paragraph.font.size = Pt(27)
    paragraph.font.bold = True
    paragraph.font.color.rgb = NAVY

    if subtitle:
        sub = slide.shapes.add_textbox(
            Inches(0.65),
            Inches(0.98),
            Inches(12),
            Inches(0.35),
        )

        paragraph = sub.text_frame.paragraphs[0]
        paragraph.text = subtitle
        paragraph.font.size = Pt(11)
        paragraph.font.color.rgb = DARK


def add_metric_card(slide, x, y, width, label, value):
    """Add a KPI card to a slide."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(width),
        Inches(1.15),
    )

    shape.fill.solid()
    shape.fill.fore_color.rgb = LIGHT_BLUE
    shape.line.color.rgb = BLUE

    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.margin_left = Inches(0.12)
    text_frame.margin_right = Inches(0.12)
    text_frame.vertical_anchor = 3

    paragraph = text_frame.paragraphs[0]
    paragraph.text = label
    paragraph.alignment = PP_ALIGN.CENTER
    paragraph.font.size = Pt(11)
    paragraph.font.color.rgb = DARK

    paragraph = text_frame.add_paragraph()
    paragraph.text = value
    paragraph.alignment = PP_ALIGN.CENTER
    paragraph.font.size = Pt(19)
    paragraph.font.bold = True
    paragraph.font.color.rgb = NAVY


def generate_powerpoint(client_id=1, fund_id=1, period_id=1):
    """Generate a four-slide portfolio report."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    summary, holdings = fetch_report_data(
        client_id=client_id,
        fund_id=fund_id,
        period_id=period_id,
    )

    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)

    # Slide 1: Cover
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = NAVY

    title_box = slide.shapes.add_textbox(
        Inches(0.9),
        Inches(1.6),
        Inches(11.5),
        Inches(1.2),
    )

    paragraph = title_box.text_frame.paragraphs[0]
    paragraph.text = "Institutional Portfolio Report"
    paragraph.font.name = "Aptos Display"
    paragraph.font.size = Pt(34)
    paragraph.font.bold = True
    paragraph.font.color.rgb = WHITE

    details = slide.shapes.add_textbox(
        Inches(0.95),
        Inches(3.0),
        Inches(11),
        Inches(1.4),
    )

    text_frame = details.text_frame
    text_frame.text = f"Client: {summary['client']}"
    text_frame.paragraphs[0].font.size = Pt(19)
    text_frame.paragraphs[0].font.color.rgb = WHITE

    paragraph = text_frame.add_paragraph()
    paragraph.text = f"Fund: {summary['fund']}"
    paragraph.font.size = Pt(17)
    paragraph.font.color.rgb = WHITE

    paragraph = text_frame.add_paragraph()
    paragraph.text = (
        f"Reporting period: {summary['period_start']} to "
        f"{summary['period_end']}"
    )
    paragraph.font.size = Pt(15)
    paragraph.font.color.rgb = WHITE

    disclaimer = slide.shapes.add_textbox(
        Inches(0.95),
        Inches(6.65),
        Inches(11.5),
        Inches(0.3),
    )

    paragraph = disclaimer.text_frame.paragraphs[0]
    paragraph.text = (
        "Illustrative portfolio data — not actual client or fund reporting"
    )
    paragraph.font.size = Pt(11)
    paragraph.font.italic = True
    paragraph.font.color.rgb = WHITE

    # Slide 2: Executive summary
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    add_title(slide, "Executive Summary", "Portfolio overview and performance")

    currency = summary["currency"] or "INR"
    market_value = summary["total_market_value"] or 0

    add_metric_card(
        slide, 0.7, 1.7, 3.7, "PORTFOLIO VALUE",
        f"{currency} {market_value:,.0f}",
    )
    add_metric_card(
        slide, 4.8, 1.7, 3.7, "HOLDINGS",
        str(summary["number_of_holdings"]),
    )
    add_metric_card(
        slide, 8.9, 1.7, 3.7, "PORTFOLIO WEIGHT",
        f"{summary['total_weight']:.2%}",
    )

    table = slide.shapes.add_table(
        4, 2, Inches(1.5), Inches(3.5), Inches(10.2), Inches(2.0)
    ).table

    table.columns[0].width = Inches(4.4)
    table.columns[1].width = Inches(5.8)

    rows = [
        ("Reporting benchmark", summary["benchmark"]),
        ("Fund return", f"{summary['fund_return']:.2%}"),
        ("Benchmark return", f"{summary['benchmark_return']:.2%}"),
        ("Active return", f"{summary['active_return']:.2%}"),
    ]

    for row_index, (label, value) in enumerate(rows):
        set_cell_text(table.cell(row_index, 0), label, bold=True, size=13)
        set_cell_text(table.cell(row_index, 1), value, size=13)
        table.cell(row_index, 0).fill.solid()
        table.cell(row_index, 0).fill.fore_color.rgb = LIGHT_BLUE

    add_footer(slide, 2)

    # Slide 3: Holdings breakdown
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    add_title(slide, "Portfolio Holdings", "Security-level exposure")

    table = slide.shapes.add_table(
        len(holdings) + 1,
        4,
        Inches(0.55),
        Inches(1.6),
        Inches(12.2),
        Inches(4.7),
    ).table

    widths = [1.5, 4.3, 2.8, 1.7]
    for index, width in enumerate(widths):
        table.columns[index].width = Inches(width)

    headers = ["Security ID", "Security Name", "Sector", "Weight"]

    for column, header in enumerate(headers):
        set_cell_text(
            table.cell(0, column),
            header,
            bold=True,
            color=WHITE,
            size=11,
        )
        table.cell(0, column).fill.solid()
        table.cell(0, column).fill.fore_color.rgb = NAVY

    for row_index, holding in enumerate(holdings, start=1):
        values = [
            holding["security_id"],
            holding["security_name"],
            holding["sector"],
            f"{holding['weight']:.2%}",
        ]

        for column, value in enumerate(values):
            set_cell_text(table.cell(row_index, column), value, size=10)

            if row_index % 2 == 0:
                table.cell(row_index, column).fill.solid()
                table.cell(row_index, column).fill.fore_color.rgb = LIGHT_BLUE

    note = slide.shapes.add_textbox(
        Inches(0.65),
        Inches(6.45),
        Inches(12),
        Inches(0.35),
    )

    paragraph = note.text_frame.paragraphs[0]
    paragraph.text = (
        "Weights are checked against a 100% portfolio total before report creation."
    )
    paragraph.font.size = Pt(10)
    paragraph.font.italic = True

    add_footer(slide, 3)

    # Slide 4: Performance interpretation
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    add_title(slide, "Performance Snapshot", "Fund return compared with benchmark")

    fund_return = summary["fund_return"]
    benchmark_return = summary["benchmark_return"]
    active_return = summary["active_return"]

    add_metric_card(
        slide, 0.8, 1.8, 3.6, "FUND RETURN", f"{fund_return:.2%}"
    )
    add_metric_card(
        slide, 4.85, 1.8, 3.6, "BENCHMARK RETURN",
        f"{benchmark_return:.2%}",
    )
    add_metric_card(
        slide, 8.9, 1.8, 3.6, "ACTIVE RETURN", f"{active_return:.2%}"
    )

    interpretation = slide.shapes.add_textbox(
        Inches(1.0),
        Inches(3.6),
        Inches(11.2),
        Inches(1.5),
    )

    text_frame = interpretation.text_frame
    text_frame.word_wrap = True
    paragraph = text_frame.paragraphs[0]

    if active_return > 0:
        paragraph.text = (
            f"The fund outperformed its benchmark by {active_return:.2%} "
            "over the selected reporting period."
        )
    elif active_return < 0:
        paragraph.text = (
            f"The fund underperformed its benchmark by {abs(active_return):.2%} "
            "over the selected reporting period."
        )
    else:
        paragraph.text = (
            "The fund return matched the benchmark over the selected period."
        )

    paragraph.font.size = Pt(20)
    paragraph.font.color.rgb = GREEN if active_return >= 0 else RED

    caveat = text_frame.add_paragraph()
    caveat.text = (
        "Active return is calculated as fund return minus benchmark return. "
        "It is not a measure of risk-adjusted performance."
    )
    caveat.font.size = Pt(12)
    caveat.font.color.rgb = DARK

    add_footer(slide, 4)

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    output_path = OUTPUT_DIR / f"client_report_{timestamp}.pptx"

    presentation.save(output_path)

    print(f"PowerPoint report created: {output_path}")
    print(f"Slides: {len(presentation.slides)}")
    print("Report data passed the basic portfolio-weight checks.")

    return output_path


if __name__ == "__main__":
    generate_powerpoint()