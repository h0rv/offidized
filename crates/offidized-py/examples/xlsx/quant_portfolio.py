# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""
Production-grade hedge fund portfolio workbook.

Generates a multi-sheet Excel file resembling what a quantitative trading desk
would use for portfolio management, risk analytics, and performance attribution.
All data is deterministically generated (seeded RNG) — no external dependencies.

Run from the offidized-py crate directory:
    cd crates/offidized-py && uv run python examples/xlsx/quant_portfolio.py
"""

import math
import os
import random
import tempfile
from datetime import date, timedelta

from offidized import (
    Workbook,
    XlsxAlignment,
    XlsxBorder,
    XlsxChart,
    XlsxChartDataRef,
    XlsxChartSeries,
    XlsxComment,
    XlsxConditionalFormatting,
    XlsxDataValidation,
    XlsxFill,
    XlsxFont,
    XlsxPageMargins,
    XlsxPageSetup,
    XlsxPrintHeaderFooter,
    XlsxSheetProtection,
    XlsxSheetViewOptions,
    XlsxSparkline,
    XlsxSparklineGroup,
    XlsxStyle,
)

random.seed(42)

# ── palette ──────────────────────────────────────────────────────────────────
DARK = "0B1117"
NAVY = "1B2A4A"
STEEL = "2C3E5A"
ACCENT = "3B82F6"  # bright blue
GREEN = "22C55E"
RED = "EF4444"
AMBER = "F59E0B"
WHITE = "FFFFFF"
LIGHT_GRAY = "F1F5F9"
MID_GRAY = "94A3B8"
BORDER_COLOR = "334155"

# ── universe ─────────────────────────────────────────────────────────────────
HOLDINGS = [
    ("AAPL", "Apple Inc", "Technology", 185.50, 0.95),
    ("MSFT", "Microsoft Corp", "Technology", 378.20, 1.12),
    ("NVDA", "NVIDIA Corp", "Technology", 495.80, 1.65),
    ("GOOGL", "Alphabet Inc", "Technology", 141.30, 1.08),
    ("AMZN", "Amazon.com Inc", "Consumer Disc", 153.40, 1.22),
    ("META", "Meta Platforms", "Technology", 355.60, 1.35),
    ("JPM", "JPMorgan Chase", "Financials", 172.90, 1.15),
    ("V", "Visa Inc", "Financials", 264.80, 0.88),
    ("JNJ", "Johnson & Johnson", "Healthcare", 158.30, 0.62),
    ("UNH", "UnitedHealth Group", "Healthcare", 527.10, 0.78),
    ("PG", "Procter & Gamble", "Consumer Stap", 152.40, 0.55),
    ("HD", "Home Depot Inc", "Consumer Disc", 342.70, 1.05),
    ("MA", "Mastercard Inc", "Financials", 418.60, 0.92),
    ("XOM", "Exxon Mobil Corp", "Energy", 104.20, 0.85),
    ("CVX", "Chevron Corp", "Energy", 155.80, 0.90),
    ("LLY", "Eli Lilly & Co", "Healthcare", 582.30, 0.72),
    ("ABBV", "AbbVie Inc", "Healthcare", 154.60, 0.68),
    ("PFE", "Pfizer Inc", "Healthcare", 28.40, 0.75),
    ("COST", "Costco Wholesale", "Consumer Stap", 572.10, 0.82),
    ("MRK", "Merck & Co", "Healthcare", 108.90, 0.65),
    ("AVGO", "Broadcom Inc", "Technology", 892.40, 1.45),
    ("TMO", "Thermo Fisher", "Healthcare", 532.80, 0.88),
    ("ACN", "Accenture plc", "Technology", 338.20, 1.02),
    ("CRM", "Salesforce Inc", "Technology", 258.40, 1.18),
    ("NEE", "NextEra Energy", "Utilities", 62.80, 0.48),
    ("LIN", "Linde plc", "Materials", 412.30, 0.78),
    ("GS", "Goldman Sachs", "Financials", 382.60, 1.32),
    ("BLK", "BlackRock Inc", "Financials", 748.20, 1.18),
    ("ISRG", "Intuitive Surgical", "Healthcare", 342.10, 0.95),
    ("CAT", "Caterpillar Inc", "Industrials", 278.40, 1.08),
]


def gen_shares():
    """Random position size between 500 and 25000 shares, round lots."""
    return random.randint(5, 250) * 100


def gen_return(base_price, beta, days=60):
    """Generate a plausible price path."""
    prices = [base_price]
    for _ in range(days):
        daily_r = random.gauss(0.0003 * beta, 0.012 * math.sqrt(abs(beta)))
        prices.append(prices[-1] * (1 + daily_r))
    return prices


# ── build workbook ───────────────────────────────────────────────────────────
wb = Workbook()

# ── styles ───────────────────────────────────────────────────────────────────


def make_style(
    font_name="Inter",
    font_size="11",
    bold=False,
    italic=False,
    fg_color=None,
    font_color=None,
    num_fmt=None,
    h_align=None,
    v_align="center",
    border_bottom=None,
    border_top=None,
    wrap=False,
):
    s = XlsxStyle()
    f = XlsxFont()
    f.name = font_name
    f.size = font_size
    f.bold = bold
    f.italic = italic
    if font_color:
        f.color = font_color
    s.set_font(f)

    if fg_color:
        fl = XlsxFill()
        fl.pattern = "solid"
        fl.foreground_color = fg_color
        s.set_fill(fl)

    if num_fmt:
        s.set_number_format(num_fmt)

    a = XlsxAlignment()
    if h_align:
        a.horizontal = h_align
    if v_align:
        a.vertical = v_align
    if wrap:
        a.wrap_text = True
    s.set_alignment(a)

    if border_bottom or border_top:
        b = XlsxBorder()
        if border_bottom:
            b.bottom_style = border_bottom[0]
            b.bottom_color = border_bottom[1]
        if border_top:
            b.top_style = border_top[0]
            b.top_color = border_top[1]
        s.set_border(b)

    return wb.add_style(s)


# Register styles
S_TITLE = make_style(font_size="16", bold=True, fg_color=NAVY, font_color=WHITE)
S_SUBTITLE = make_style(font_size="11", italic=True, fg_color=NAVY, font_color=MID_GRAY)
S_HDR = make_style(
    font_size="10",
    bold=True,
    fg_color=STEEL,
    font_color=WHITE,
    h_align="center",
    border_bottom=("thin", BORDER_COLOR),
)
S_HDR_LEFT = make_style(
    font_size="10",
    bold=True,
    fg_color=STEEL,
    font_color=WHITE,
    border_bottom=("thin", BORDER_COLOR),
)
S_BODY = make_style(font_size="10")
S_BODY_ALT = make_style(font_size="10", fg_color=LIGHT_GRAY)
S_BODY_CENTER = make_style(font_size="10", h_align="center")
S_BODY_CENTER_ALT = make_style(font_size="10", h_align="center", fg_color=LIGHT_GRAY)
S_USD = make_style(font_size="10", num_fmt="#,##0.00", h_align="right")
S_USD_ALT = make_style(
    font_size="10", num_fmt="#,##0.00", h_align="right", fg_color=LIGHT_GRAY
)
S_USD_BOLD = make_style(
    font_size="10",
    num_fmt="#,##0.00",
    h_align="right",
    bold=True,
    border_top=("double", BORDER_COLOR),
)
S_INT = make_style(font_size="10", num_fmt="#,##0", h_align="right")
S_INT_ALT = make_style(
    font_size="10", num_fmt="#,##0", h_align="right", fg_color=LIGHT_GRAY
)
S_PCT = make_style(font_size="10", num_fmt="0.00%", h_align="right")
S_PCT_ALT = make_style(
    font_size="10", num_fmt="0.00%", h_align="right", fg_color=LIGHT_GRAY
)
S_PCT_BOLD = make_style(
    font_size="10",
    num_fmt="0.00%",
    h_align="right",
    bold=True,
    border_top=("double", BORDER_COLOR),
)
S_BPS = make_style(font_size="10", num_fmt="0.0", h_align="right")
S_BPS_ALT = make_style(
    font_size="10", num_fmt="0.0", h_align="right", fg_color=LIGHT_GRAY
)
S_DATE = make_style(font_size="10", num_fmt="yyyy-mm-dd", h_align="center")
S_DATE_ALT = make_style(
    font_size="10", num_fmt="yyyy-mm-dd", h_align="center", fg_color=LIGHT_GRAY
)
S_KPI_LABEL = make_style(
    font_size="10", fg_color=NAVY, font_color=MID_GRAY, h_align="right"
)
S_KPI_VALUE = make_style(
    font_size="14", bold=True, fg_color=NAVY, font_color=WHITE, h_align="left"
)
S_KPI_VALUE_PCT = make_style(
    font_size="14",
    bold=True,
    fg_color=NAVY,
    font_color=WHITE,
    h_align="left",
    num_fmt="0.00%",
)
S_KPI_VALUE_USD = make_style(
    font_size="14",
    bold=True,
    fg_color=NAVY,
    font_color=WHITE,
    h_align="left",
    num_fmt="$#,##0",
)
S_KPI_VALUE_X = make_style(
    font_size="14",
    bold=True,
    fg_color=NAVY,
    font_color=WHITE,
    h_align="left",
    num_fmt='0.00"x"',
)
S_SECTION = make_style(
    font_size="11",
    bold=True,
    fg_color=NAVY,
    font_color=ACCENT,
    border_bottom=("thin", ACCENT),
)
S_TOTAL_LABEL = make_style(
    font_size="10", bold=True, border_top=("double", BORDER_COLOR)
)
S_TOTAL_INT = make_style(
    font_size="10",
    num_fmt="#,##0",
    h_align="right",
    bold=True,
    border_top=("double", BORDER_COLOR),
)


def col(c):
    """Column letter(s) from 0-based index."""
    if c < 26:
        return chr(65 + c)
    return chr(64 + c // 26) + chr(65 + c % 26)


def excel_serial_to_iso(serial):
    """Convert an Excel date serial into an ISO date string."""
    return (date(1899, 12, 30) + timedelta(days=int(serial))).isoformat()


# ═════════════════════════════════════════════════════════════════════════════
#  SHEET 1 — HOLDINGS
# ═════════════════════════════════════════════════════════════════════════════

print("Building Sheet 1: Holdings...")
ws = wb.add_sheet("Holdings")
ws.set_tab_color(ACCENT)

# Generate position data
positions = []
for ticker, name, sector, base_price, beta in HOLDINGS:
    shares = gen_shares()
    prices = gen_return(base_price, beta)
    entry_price = prices[0]
    current_price = prices[-1]
    mkt_val = shares * current_price
    unrealized_pnl = shares * (current_price - entry_price)
    pnl_pct = (current_price - entry_price) / entry_price
    vol_30d = (
        sum(
            abs((prices[i] - prices[i - 1]) / prices[i - 1])
            for i in range(len(prices) - 30, len(prices))
        )
        / 30
        * math.sqrt(252)
    )
    positions.append(
        (
            ticker,
            name,
            sector,
            shares,
            entry_price,
            current_price,
            mkt_val,
            unrealized_pnl,
            pnl_pct,
            beta,
            vol_30d,
            prices,
        )
    )

total_mkt_val = sum(p[6] for p in positions)

# Title row
ws.add_merged_range("A1:L1")
ws.set_cell_value("A1", "PORTFOLIO HOLDINGS — Meridian Capital Partners")
ws.cell("A1").set_style_id(S_TITLE)
ws.row(1).set_height(32.0)

ws.add_merged_range("A2:L2")
ws.set_cell_value(
    "A2",
    "As of 2026-02-27  |  Fund: Meridian Systematic Alpha  |  NAV: $847.2M  |  30 positions",
)
ws.cell("A2").set_style_id(S_SUBTITLE)

# Headers
hdrs = [
    "Ticker",
    "Name",
    "Sector",
    "Shares",
    "Entry Px",
    "Current Px",
    "Market Value",
    "Unreal P&L",
    "P&L %",
    "Weight %",
    "Beta",
    "30d Vol",
]
widths = [10, 24, 16, 12, 13, 13, 17, 17, 11, 11, 9, 11]
for i, (h, w) in enumerate(zip(hdrs, widths)):
    ref = f"{col(i)}4"
    ws.set_cell_value(ref, h)
    ws.cell(ref).set_style_id(S_HDR if i > 2 else S_HDR_LEFT)
    ws.column(i + 1).set_width(float(w))

ws.row(4).set_height(22.0)

# Data rows
for r, (
    ticker,
    name,
    sector,
    shares,
    entry_px,
    cur_px,
    mkt_val,
    unreal_pnl,
    pnl_pct,
    beta,
    vol30,
    _prices,
) in enumerate(positions):
    row = r + 5  # 1-based row
    alt = r % 2 == 1

    ws.set_cell_value(f"A{row}", ticker)
    ws.cell(f"A{row}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws.set_cell_value(f"B{row}", name)
    ws.cell(f"B{row}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws.set_cell_value(f"C{row}", sector)
    ws.cell(f"C{row}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws.set_cell_value(f"D{row}", shares)
    ws.cell(f"D{row}").set_style_id(S_INT_ALT if alt else S_INT)

    ws.set_cell_value(f"E{row}", round(entry_px, 2))
    ws.cell(f"E{row}").set_style_id(S_USD_ALT if alt else S_USD)

    ws.set_cell_value(f"F{row}", round(cur_px, 2))
    ws.cell(f"F{row}").set_style_id(S_USD_ALT if alt else S_USD)

    # Market value = shares * current price (formula)
    ws.set_cell_formula(f"G{row}", f"D{row}*F{row}")
    ws.cell(f"G{row}").set_style_id(S_USD_ALT if alt else S_USD)

    # Unrealized P&L = shares * (current - entry)
    ws.set_cell_formula(f"H{row}", f"D{row}*(F{row}-E{row})")
    ws.cell(f"H{row}").set_style_id(S_USD_ALT if alt else S_USD)

    # P&L % = (current - entry) / entry
    ws.set_cell_formula(f"I{row}", f"(F{row}-E{row})/E{row}")
    ws.cell(f"I{row}").set_style_id(S_PCT_ALT if alt else S_PCT)

    # Weight = market value / total (formula referencing totals row)
    ws.set_cell_formula(f"J{row}", f"G{row}/G35")
    ws.cell(f"J{row}").set_style_id(S_PCT_ALT if alt else S_PCT)

    ws.set_cell_value(f"K{row}", round(beta, 2))
    ws.cell(f"K{row}").set_style_id(S_BPS_ALT if alt else S_BPS)

    ws.set_cell_value(f"L{row}", round(vol30, 4))
    ws.cell(f"L{row}").set_style_id(S_PCT_ALT if alt else S_PCT)

# Totals row (row 35)
total_row = 5 + len(positions)
ws.set_cell_value(f"A{total_row}", "TOTAL")
ws.cell(f"A{total_row}").set_style_id(S_TOTAL_LABEL)
for c in range(1, 3):
    ws.set_cell_value(f"{col(c)}{total_row}", "")
    ws.cell(f"{col(c)}{total_row}").set_style_id(S_TOTAL_LABEL)

ws.set_cell_formula(f"D{total_row}", f"SUM(D5:D{total_row - 1})")
ws.cell(f"D{total_row}").set_style_id(S_TOTAL_INT)

for c_letter in ["E", "F"]:
    ws.set_cell_value(f"{c_letter}{total_row}", "")
    ws.cell(f"{c_letter}{total_row}").set_style_id(S_USD_BOLD)

ws.set_cell_formula(f"G{total_row}", f"SUM(G5:G{total_row - 1})")
ws.cell(f"G{total_row}").set_style_id(S_USD_BOLD)

ws.set_cell_formula(f"H{total_row}", f"SUM(H5:H{total_row - 1})")
ws.cell(f"H{total_row}").set_style_id(S_USD_BOLD)

# Weighted average P&L %
ws.set_cell_formula(
    f"I{total_row}", f"SUMPRODUCT(I5:I{total_row - 1},G5:G{total_row - 1})/G{total_row}"
)
ws.cell(f"I{total_row}").set_style_id(S_PCT_BOLD)

ws.set_cell_value(f"J{total_row}", 1.0)
ws.cell(f"J{total_row}").set_style_id(S_PCT_BOLD)

# Weighted beta
ws.set_cell_formula(
    f"K{total_row}", f"SUMPRODUCT(K5:K{total_row - 1},G5:G{total_row - 1})/G{total_row}"
)
ws.cell(f"K{total_row}").set_style_id(S_USD_BOLD)

ws.set_cell_value(f"L{total_row}", "")
ws.cell(f"L{total_row}").set_style_id(S_USD_BOLD)

# Conditional formatting on P&L %
cf_gain = XlsxConditionalFormatting("expression", [f"I5:I{total_row - 1}"], ["I5>0"])
ws.add_conditional_formatting(cf_gain)
cf_loss = XlsxConditionalFormatting("expression", [f"I5:I{total_row - 1}"], ["I5<0"])
ws.add_conditional_formatting(cf_loss)

# Freeze panes below headers
ws.freeze_panes(0, 5)

# Sheet view
view = XlsxSheetViewOptions()
view.show_gridlines = False
view.zoom_scale = 90
ws.set_sheet_view_options(view)


# ═════════════════════════════════════════════════════════════════════════════
#  SHEET 2 — DAILY RETURNS
# ═════════════════════════════════════════════════════════════════════════════

print("Building Sheet 2: Daily Returns...")
ws2 = wb.add_sheet("Daily Returns")
ws2.set_tab_color(GREEN)

# Generate 60 trading days of returns
DAYS = 60
dates = []
base_date_serial = 46078  # ~ 2026-02-27 minus 60 trading days
port_returns = []
sp500_returns = []
ndx_returns = []
rut_returns = []
cumul_port_values = []
cumul_sp500_values = []

for d in range(DAYS):
    dates.append(base_date_serial + d)
    # Correlated daily returns
    mkt = random.gauss(0.0004, 0.011)
    port_r = mkt * 1.15 + random.gauss(0.0002, 0.005)  # alpha + beta*market
    sp_r = mkt + random.gauss(0, 0.002)
    ndx_r = mkt * 1.1 + random.gauss(0, 0.003)
    rut_r = mkt * 0.9 + random.gauss(0, 0.004)
    port_returns.append(port_r)
    sp500_returns.append(sp_r)
    ndx_returns.append(ndx_r)
    rut_returns.append(rut_r)
    if d == 0:
        cumul_port_values.append(1 + port_r)
        cumul_sp500_values.append(1 + sp_r)
    else:
        cumul_port_values.append(cumul_port_values[-1] * (1 + port_r))
        cumul_sp500_values.append(cumul_sp500_values[-1] * (1 + sp_r))

# Title
ws2.add_merged_range("A1:I1")
ws2.set_cell_value("A1", "DAILY RETURNS — 60 Trading Day Window")
ws2.cell("A1").set_style_id(S_TITLE)
ws2.row(1).set_height(32.0)

# Headers
ret_hdrs = [
    "Date",
    "Portfolio",
    "S&P 500",
    "Nasdaq",
    "Russell 2000",
    "Alpha",
    "Cumul Port",
    "Cumul S&P",
    "Spread",
]
ret_widths = [13, 12, 12, 12, 14, 12, 13, 13, 13]
for i, (h, w) in enumerate(zip(ret_hdrs, ret_widths)):
    ref = f"{col(i)}3"
    ws2.set_cell_value(ref, h)
    ws2.cell(ref).set_style_id(S_HDR)
    ws2.column(i + 1).set_width(float(w))

# Data
for d in range(DAYS):
    r = d + 4  # 1-based row
    alt = d % 2 == 1

    ws2.set_cell_value(f"A{r}", dates[d])
    ws2.cell(f"A{r}").set_style_id(S_DATE_ALT if alt else S_DATE)

    ws2.set_cell_value(f"B{r}", round(port_returns[d], 6))
    ws2.cell(f"B{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    ws2.set_cell_value(f"C{r}", round(sp500_returns[d], 6))
    ws2.cell(f"C{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    ws2.set_cell_value(f"D{r}", round(ndx_returns[d], 6))
    ws2.cell(f"D{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    ws2.set_cell_value(f"E{r}", round(rut_returns[d], 6))
    ws2.cell(f"E{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    # Alpha = Portfolio - Beta*SP500 (assume beta 1.15)
    ws2.set_cell_formula(f"F{r}", f"B{r}-1.15*C{r}")
    ws2.cell(f"F{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    # Cumulative returns (compound)
    if d == 0:
        ws2.set_cell_formula(f"G{r}", f"1+B{r}")
        ws2.set_cell_formula(f"H{r}", f"1+C{r}")
    else:
        ws2.set_cell_formula(f"G{r}", f"G{r - 1}*(1+B{r})")
        ws2.set_cell_formula(f"H{r}", f"H{r - 1}*(1+C{r})")

    ws2.cell(f"G{r}").set_style_id(S_USD_ALT if alt else S_USD)
    ws2.cell(f"H{r}").set_style_id(S_USD_ALT if alt else S_USD)

    # Spread = cumul port - cumul sp500
    ws2.set_cell_formula(f"I{r}", f"G{r}-H{r}")
    ws2.cell(f"I{r}").set_style_id(S_USD_ALT if alt else S_USD)

last_data_row = 3 + DAYS

# Summary stats below data
summary_row = last_data_row + 2
labels = ["Total Return", "Annualized", "Daily Sharpe", "Max Drawdown", "Win Rate"]
for i, lbl in enumerate(labels):
    r = summary_row + i
    ws2.set_cell_value(f"A{r}", lbl)
    ws2.cell(f"A{r}").set_style_id(S_KPI_LABEL)

# Total return = last cumulative - 1
ws2.set_cell_formula(f"B{summary_row}", f"G{last_data_row}-1")
ws2.cell(f"B{summary_row}").set_style_id(S_KPI_VALUE_PCT)

# Annualized = (1+total)^(252/days) - 1
ws2.set_cell_formula(f"B{summary_row + 1}", f"(1+B{summary_row})^(252/{DAYS})-1")
ws2.cell(f"B{summary_row + 1}").set_style_id(S_KPI_VALUE_PCT)

# Daily Sharpe = avg(daily)/stdev(daily)*sqrt(252)
ws2.set_cell_formula(
    f"B{summary_row + 2}",
    f"AVERAGE(B4:B{last_data_row})/STDEV(B4:B{last_data_row})*SQRT(252)",
)
ws2.cell(f"B{summary_row + 2}").set_style_id(S_KPI_VALUE_X)

# Max drawdown (approximate with MIN of daily returns * sqrt(5) as proxy)
ws2.set_cell_formula(f"B{summary_row + 3}", f"MIN(B4:B{last_data_row})*SQRT(5)")
ws2.cell(f"B{summary_row + 3}").set_style_id(S_KPI_VALUE_PCT)

# Win rate
ws2.set_cell_formula(f"B{summary_row + 4}", f'COUNTIF(B4:B{last_data_row},">0")/{DAYS}')
ws2.cell(f"B{summary_row + 4}").set_style_id(S_KPI_VALUE_PCT)

# Chart — cumulative returns
chart = XlsxChart("line")
chart.title = "Cumulative Returns: Portfolio vs S&P 500"
chart.set_anchor(0, summary_row + 6, 9, summary_row + 26)

s1 = XlsxChartSeries(0, 0)
s1.name = "Portfolio"
date_ref = XlsxChartDataRef.from_formula(f"'Daily Returns'!$A$4:$A${last_data_row}")
date_ref.str_values = [excel_serial_to_iso(serial) for serial in dates]
s1.set_categories(date_ref)

portfolio_ref = XlsxChartDataRef.from_formula(
    f"'Daily Returns'!$G$4:$G${last_data_row}"
)
portfolio_ref.num_values = cumul_port_values
s1.set_values(portfolio_ref)
chart.add_series(s1)

s2 = XlsxChartSeries(1, 1)
s2.name = "S&P 500"
sp500_ref = XlsxChartDataRef.from_formula(f"'Daily Returns'!$H$4:$H${last_data_row}")
sp500_ref.num_values = cumul_sp500_values
s2.set_values(sp500_ref)
chart.add_series(s2)

ws2.add_chart(chart)

# Conditional formatting — highlight negative alpha days
cf_neg_alpha = XlsxConditionalFormatting(
    "expression", [f"F4:F{last_data_row}"], ["F4<0"]
)
ws2.add_conditional_formatting(cf_neg_alpha)

ws2.freeze_panes(0, 4)

view2 = XlsxSheetViewOptions()
view2.show_gridlines = False
view2.zoom_scale = 90
ws2.set_sheet_view_options(view2)


# ═════════════════════════════════════════════════════════════════════════════
#  SHEET 3 — RISK ANALYTICS
# ═════════════════════════════════════════════════════════════════════════════

print("Building Sheet 3: Risk Analytics...")
ws3 = wb.add_sheet("Risk Analytics")
ws3.set_tab_color(RED)

ws3.add_merged_range("A1:H1")
ws3.set_cell_value("A1", "RISK ANALYTICS — Meridian Systematic Alpha")
ws3.cell("A1").set_style_id(S_TITLE)
ws3.row(1).set_height(32.0)

# ── Risk Metrics Panel ──
ws3.add_merged_range("A3:D3")
ws3.set_cell_value("A3", "KEY RISK METRICS")
ws3.cell("A3").set_style_id(S_SECTION)

risk_metrics = [
    (
        "VaR (95%, 1-day)",
        f"=-PERCENTILE('Daily Returns'!$B$4:$B${last_data_row},0.05)",
        S_KPI_VALUE_PCT,
    ),
    (
        "VaR (99%, 1-day)",
        f"=-PERCENTILE('Daily Returns'!$B$4:$B${last_data_row},0.01)",
        S_KPI_VALUE_PCT,
    ),
    ("CVaR (95%)", None, S_KPI_VALUE_PCT),
    (
        "Portfolio Beta",
        "=SUMPRODUCT(Holdings!$K$5:$K$34,Holdings!$G$5:$G$34)/Holdings!$G$35",
        S_KPI_VALUE_X,
    ),
    (
        "Sharpe Ratio",
        f"=AVERAGE('Daily Returns'!$B$4:$B${last_data_row})/STDEV('Daily Returns'!$B$4:$B${last_data_row})*SQRT(252)",
        S_KPI_VALUE_X,
    ),
    (
        "Max Drawdown",
        f"=MIN('Daily Returns'!$B$4:$B${last_data_row})*SQRT(5)",
        S_KPI_VALUE_PCT,
    ),
    (
        "Tracking Error",
        f"=STDEV('Daily Returns'!$F$4:$F${last_data_row})*SQRT(252)",
        S_KPI_VALUE_PCT,
    ),
    (
        "Information Ratio",
        f"=AVERAGE('Daily Returns'!$F$4:$F${last_data_row})/STDEV('Daily Returns'!$F$4:$F${last_data_row})*SQRT(252)",
        S_KPI_VALUE_X,
    ),
]

for i, (label, formula, style_id) in enumerate(risk_metrics):
    r = 5 + i
    ws3.set_cell_value(f"A{r}", label)
    ws3.cell(f"A{r}").set_style_id(S_KPI_LABEL)
    ws3.column(1).set_width(20.0)
    if formula:
        ws3.set_cell_formula(f"B{r}", formula.lstrip("="))
    else:
        # CVaR approximation (average of worst 5% daily returns)
        worst_n = max(1, int(len(port_returns) * 0.05))
        cvar_pct = -sum(sorted(port_returns)[:worst_n]) / worst_n
        ws3.set_cell_value(f"B{r}", round(cvar_pct, 6))
    ws3.cell(f"B{r}").set_style_id(style_id)
    ws3.column(2).set_width(16.0)

# ── Sector Exposure ──
ws3.add_merged_range("A15:D15")
ws3.set_cell_value("A15", "SECTOR EXPOSURE")
ws3.cell("A15").set_style_id(S_SECTION)

sector_hdrs = ["Sector", "# Positions", "Gross Exposure", "Weight"]
for i, h in enumerate(sector_hdrs):
    ws3.set_cell_value(f"{col(i)}17", h)
    ws3.cell(f"{col(i)}17").set_style_id(S_HDR if i > 0 else S_HDR_LEFT)
ws3.column(3).set_width(16.0)
ws3.column(4).set_width(12.0)

# Aggregate by sector
sector_data = {}
for t, n, sec, sh, ep, cp, mv, upnl, pp, beta, vol, _ in positions:
    if sec not in sector_data:
        sector_data[sec] = {"count": 0, "exposure": 0.0}
    sector_data[sec]["count"] += 1
    sector_data[sec]["exposure"] += mv

for i, (sector, data) in enumerate(
    sorted(sector_data.items(), key=lambda x: -x[1]["exposure"])
):
    r = 18 + i
    alt = i % 2 == 1
    ws3.set_cell_value(f"A{r}", sector)
    ws3.cell(f"A{r}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws3.set_cell_value(f"B{r}", data["count"])
    ws3.cell(f"B{r}").set_style_id(S_BODY_CENTER_ALT if alt else S_BODY_CENTER)

    ws3.set_cell_value(f"C{r}", round(data["exposure"], 2))
    ws3.cell(f"C{r}").set_style_id(S_USD_ALT if alt else S_USD)

    ws3.set_cell_value(f"D{r}", round(data["exposure"] / total_mkt_val, 4))
    ws3.cell(f"D{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

sector_end = 18 + len(sector_data)
sector_labels = [
    sector
    for sector, _data in sorted(sector_data.items(), key=lambda kv: -kv[1]["exposure"])
]
sector_weight_values = [
    round(sector_data[sector]["exposure"] / total_mkt_val, 4)
    for sector in sector_labels
]

# Sector pie chart
sector_chart = XlsxChart("pie")
sector_chart.title = "Sector Allocation"
sector_chart.set_anchor(5, 15, 12, 32)
sec_series = XlsxChartSeries(0, 0)
sector_cat_ref = XlsxChartDataRef.from_formula(
    f"'Risk Analytics'!$A$18:$A${sector_end - 1}"
)
sector_cat_ref.str_values = sector_labels
sec_series.set_categories(sector_cat_ref)

sector_val_ref = XlsxChartDataRef.from_formula(
    f"'Risk Analytics'!$D$18:$D${sector_end - 1}"
)
sector_val_ref.num_values = sector_weight_values
sec_series.set_values(sector_val_ref)
sector_chart.add_series(sec_series)
ws3.add_chart(sector_chart)

# ── Top Winners / Losers ──
ws3.add_merged_range("A34:D34")
ws3.set_cell_value("A34", "TOP 5 WINNERS / BOTTOM 5 LOSERS")
ws3.cell("A34").set_style_id(S_SECTION)

winner_hdrs = ["Ticker", "Name", "P&L", "P&L %"]
for i, h in enumerate(winner_hdrs):
    ws3.set_cell_value(f"{col(i)}36", h)
    ws3.cell(f"{col(i)}36").set_style_id(S_HDR if i > 1 else S_HDR_LEFT)

sorted_by_pnl = sorted(positions, key=lambda p: p[7], reverse=True)
top5 = sorted_by_pnl[:5]
bottom5 = sorted_by_pnl[-5:][::-1]

for i, pos in enumerate(top5 + bottom5):
    r = 37 + i
    alt = i % 2 == 1
    ws3.set_cell_value(f"A{r}", pos[0])
    ws3.cell(f"A{r}").set_style_id(S_BODY_ALT if alt else S_BODY)
    ws3.set_cell_value(f"B{r}", pos[1])
    ws3.cell(f"B{r}").set_style_id(S_BODY_ALT if alt else S_BODY)
    ws3.set_cell_value(f"C{r}", round(pos[7], 2))
    ws3.cell(f"C{r}").set_style_id(S_USD_ALT if alt else S_USD)
    ws3.set_cell_value(f"D{r}", round(pos[8], 4))
    ws3.cell(f"D{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

# Add a comment on the risk sheet
cmt = XlsxComment(
    "A1",
    "Risk Team",
    "Updated automatically from live portfolio feed. VaR uses historical simulation.",
    False,
)
ws3.add_comment(cmt)

view3 = XlsxSheetViewOptions()
view3.show_gridlines = False
view3.zoom_scale = 90
ws3.set_sheet_view_options(view3)


# ═════════════════════════════════════════════════════════════════════════════
#  SHEET 4 — TRADE BLOTTER
# ═════════════════════════════════════════════════════════════════════════════

print("Building Sheet 4: Trade Blotter...")
ws4 = wb.add_sheet("Trade Blotter")
ws4.set_tab_color(AMBER)

ws4.add_merged_range("A1:J1")
ws4.set_cell_value("A1", "TRADE BLOTTER — Last 30 Days")
ws4.cell("A1").set_style_id(S_TITLE)
ws4.row(1).set_height(32.0)

trade_hdrs = [
    "Date",
    "Ticker",
    "Side",
    "Qty",
    "Price",
    "Notional",
    "Commission",
    "Strategy",
    "Broker",
    "Status",
]
trade_widths = [13, 10, 8, 11, 12, 15, 13, 16, 14, 12]
for i, (h, w) in enumerate(zip(trade_hdrs, trade_widths)):
    ref = f"{col(i)}3"
    ws4.set_cell_value(ref, h)
    ws4.cell(ref).set_style_id(S_HDR if i > 1 else S_HDR_LEFT)
    ws4.column(i + 1).set_width(float(w))

strategies = [
    "Momentum",
    "Mean Reversion",
    "Stat Arb",
    "Event Driven",
    "Vol Arb",
    "Sector Rotation",
]
brokers = ["Goldman Sachs", "Morgan Stanley", "JP Morgan", "Citadel Sec", "Virtu"]
sides = ["BUY", "SELL", "BUY", "BUY", "SELL"]  # slight long bias

trades = []
for t in range(40):
    pos = random.choice(positions)
    side = random.choice(sides)
    qty = random.randint(1, 50) * 100
    px = pos[5] * (1 + random.gauss(0, 0.005))
    notional = qty * px
    comm = notional * 0.00015
    strategy = random.choice(strategies)
    broker = random.choice(brokers)
    date_serial = 46078 + random.randint(0, 59)
    trades.append(
        (date_serial, pos[0], side, qty, px, notional, comm, strategy, broker, "Filled")
    )

trades.sort(key=lambda x: x[0], reverse=True)

for i, (dt, ticker, side, qty, px, notional, comm, strat, broker, status) in enumerate(
    trades
):
    r = 4 + i
    alt = i % 2 == 1

    ws4.set_cell_value(f"A{r}", dt)
    ws4.cell(f"A{r}").set_style_id(S_DATE_ALT if alt else S_DATE)

    ws4.set_cell_value(f"B{r}", ticker)
    ws4.cell(f"B{r}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws4.set_cell_value(f"C{r}", side)
    ws4.cell(f"C{r}").set_style_id(S_BODY_CENTER_ALT if alt else S_BODY_CENTER)

    ws4.set_cell_value(f"D{r}", qty)
    ws4.cell(f"D{r}").set_style_id(S_INT_ALT if alt else S_INT)

    ws4.set_cell_value(f"E{r}", round(px, 2))
    ws4.cell(f"E{r}").set_style_id(S_USD_ALT if alt else S_USD)

    ws4.set_cell_formula(f"F{r}", f"D{r}*E{r}")
    ws4.cell(f"F{r}").set_style_id(S_USD_ALT if alt else S_USD)

    ws4.set_cell_formula(f"G{r}", f"F{r}*0.00015")
    ws4.cell(f"G{r}").set_style_id(S_USD_ALT if alt else S_USD)

    ws4.set_cell_value(f"H{r}", strat)
    ws4.cell(f"H{r}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws4.set_cell_value(f"I{r}", broker)
    ws4.cell(f"I{r}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws4.set_cell_value(f"J{r}", status)
    ws4.cell(f"J{r}").set_style_id(S_BODY_CENTER_ALT if alt else S_BODY_CENTER)

trade_end = 3 + len(trades)

# Totals
ws4.set_cell_value(f"A{trade_end + 1}", "TOTAL")
ws4.cell(f"A{trade_end + 1}").set_style_id(S_TOTAL_LABEL)
for c in range(1, 5):
    ws4.set_cell_value(f"{col(c)}{trade_end + 1}", "")
    ws4.cell(f"{col(c)}{trade_end + 1}").set_style_id(S_TOTAL_LABEL)
ws4.set_cell_formula(f"F{trade_end + 1}", f"SUM(F4:F{trade_end})")
ws4.cell(f"F{trade_end + 1}").set_style_id(S_USD_BOLD)
ws4.set_cell_formula(f"G{trade_end + 1}", f"SUM(G4:G{trade_end})")
ws4.cell(f"G{trade_end + 1}").set_style_id(S_USD_BOLD)

# Data validation on Side column
dv_side = XlsxDataValidation.list([f"C4:C{trade_end}"], '"BUY,SELL,SHORT,COVER"')
dv_side.prompt_title = "Trade Side"
dv_side.prompt_message = "Select trade direction"
dv_side.show_input_message = True
ws4.add_data_validation(dv_side)

# Auto-filter
ws4.set_auto_filter(f"A3:J{trade_end}")

ws4.freeze_panes(0, 4)

view4 = XlsxSheetViewOptions()
view4.show_gridlines = False
view4.zoom_scale = 90
ws4.set_sheet_view_options(view4)


# ═════════════════════════════════════════════════════════════════════════════
#  SHEET 5 — EXECUTIVE DASHBOARD
# ═════════════════════════════════════════════════════════════════════════════

print("Building Sheet 5: Dashboard...")
ws5 = wb.add_sheet("Dashboard")
ws5.set_tab_color(NAVY)

ws5.add_merged_range("A1:L1")
ws5.set_cell_value("A1", "MERIDIAN CAPITAL PARTNERS — Executive Dashboard")
ws5.cell("A1").set_style_id(S_TITLE)
ws5.row(1).set_height(36.0)

ws5.add_merged_range("A2:L2")
ws5.set_cell_value(
    "A2", "Meridian Systematic Alpha Fund  |  February 2026  |  CONFIDENTIAL"
)
ws5.cell("A2").set_style_id(S_SUBTITLE)

# ── KPI Row ──
ws5.add_merged_range("A4:L4")
ws5.set_cell_value("A4", "PERFORMANCE SNAPSHOT")
ws5.cell("A4").set_style_id(S_SECTION)
ws5.column(1).set_width(16.0)
ws5.column(2).set_width(14.0)

kpis = [
    ("A", "B", "NAV", 847200000, S_KPI_VALUE_USD),
    ("C", "D", "MTD Return", round(sum(port_returns[-20:]), 4), S_KPI_VALUE_PCT),
    ("E", "F", "YTD Return", round(sum(port_returns), 4), S_KPI_VALUE_PCT),
    ("G", "H", "Sharpe (Ann)", None, S_KPI_VALUE_X),
    ("I", "J", "# Positions", 30, S_KPI_VALUE),
    ("K", "L", "Gross Exp", round(total_mkt_val, 0), S_KPI_VALUE_USD),
]
for i in range(12):
    # Value columns (B, D, F, H, J, L) need extra width for currency/numbers
    ws5.column(i + 1).set_width(20.0 if i % 2 == 1 else 14.0)

for lbl_col, val_col, label, value, style in kpis:
    ws5.set_cell_value(f"{lbl_col}6", label)
    ws5.cell(f"{lbl_col}6").set_style_id(S_KPI_LABEL)
    if value is not None:
        ws5.set_cell_value(f"{val_col}6", value)
    else:
        ws5.set_cell_formula(
            f"{val_col}6",
            f"AVERAGE('Daily Returns'!$B$4:$B${last_data_row})/STDEV('Daily Returns'!$B$4:$B${last_data_row})*SQRT(252)",
        )
    ws5.cell(f"{val_col}6").set_style_id(style)

# ── Monthly Attribution ──
ws5.add_merged_range("A9:F9")
ws5.set_cell_value("A9", "MONTHLY RETURN ATTRIBUTION BY SECTOR")
ws5.cell("A9").set_style_id(S_SECTION)

attr_hdrs = ["Sector", "Weight", "Return", "Contribution", "# Pos", "Trend"]
for i, h in enumerate(attr_hdrs):
    ws5.set_cell_value(f"{col(i)}11", h)
    ws5.cell(f"{col(i)}11").set_style_id(S_HDR if i > 0 else S_HDR_LEFT)

# Build monthly sector attribution data + sparkline source data
sector_order = sorted(sector_data.keys(), key=lambda s: -sector_data[s]["exposure"])
dashboard_sector_weights = []
for i, sector in enumerate(sector_order):
    r = 12 + i
    alt = i % 2 == 1
    data = sector_data[sector]
    weight = data["exposure"] / total_mkt_val
    dashboard_sector_weights.append(round(weight, 4))
    # Simulated sector return
    sec_return = random.gauss(0.02, 0.04)
    contribution = weight * sec_return

    ws5.set_cell_value(f"A{r}", sector)
    ws5.cell(f"A{r}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws5.set_cell_value(f"B{r}", round(weight, 4))
    ws5.cell(f"B{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    ws5.set_cell_value(f"C{r}", round(sec_return, 4))
    ws5.cell(f"C{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    ws5.set_cell_value(f"D{r}", round(contribution, 4))
    ws5.cell(f"D{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    ws5.set_cell_value(f"E{r}", data["count"])
    ws5.cell(f"E{r}").set_style_id(S_BODY_CENTER_ALT if alt else S_BODY_CENTER)

    # Sparkline source: 6 data points for trend in hidden columns
    for j in range(6):
        c_idx = 12 + j  # columns M through R (hidden)
        val = random.gauss(sec_return / 6, 0.01)
        ws5.set_cell_value(f"{col(c_idx)}{r}", round(val, 4))

# Sparklines for the trend column
spark = XlsxSparklineGroup()
spark.sparkline_type = "line"
spark.high_point = True
spark.low_point = True
spark.line_weight = 1.25

n_sectors = len(sector_order)
for i in range(n_sectors):
    r = 12 + i
    spark.add_sparkline(XlsxSparkline(f"F{r}", f"Dashboard!$M${r}:$R${r}"))
ws5.add_sparkline_group(spark)

# Hide sparkline source columns
for c in range(12, 18):
    ws5.column(c + 1).set_hidden(True)

attr_end = 12 + n_sectors

# ── Top Movers ──
ws5.add_merged_range(f"A{attr_end + 2}:F{attr_end + 2}")
ws5.set_cell_value(f"A{attr_end + 2}", "TOP MOVERS — Period P&L")
ws5.cell(f"A{attr_end + 2}").set_style_id(S_SECTION)

mover_hdrs = ["Ticker", "Name", "P&L ($)", "P&L (%)", "Weight", "Signal"]
for i, h in enumerate(mover_hdrs):
    r = attr_end + 4
    ws5.set_cell_value(f"{col(i)}{r}", h)
    ws5.cell(f"{col(i)}{r}").set_style_id(S_HDR if i > 1 else S_HDR_LEFT)

# Show top 10 by absolute P&L
top_movers = sorted(positions, key=lambda p: abs(p[7]), reverse=True)[:10]
for i, pos in enumerate(top_movers):
    r = attr_end + 5 + i
    alt = i % 2 == 1

    ws5.set_cell_value(f"A{r}", pos[0])
    ws5.cell(f"A{r}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws5.set_cell_value(f"B{r}", pos[1])
    ws5.cell(f"B{r}").set_style_id(S_BODY_ALT if alt else S_BODY)

    ws5.set_cell_value(f"C{r}", round(pos[7], 2))
    ws5.cell(f"C{r}").set_style_id(S_USD_ALT if alt else S_USD)

    ws5.set_cell_value(f"D{r}", round(pos[8], 4))
    ws5.cell(f"D{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    ws5.set_cell_value(f"E{r}", round(pos[6] / total_mkt_val, 4))
    ws5.cell(f"E{r}").set_style_id(S_PCT_ALT if alt else S_PCT)

    # Signal score
    signal = (
        "STRONG BUY"
        if pos[8] > 0.03
        else "BUY"
        if pos[8] > 0
        else "SELL"
        if pos[8] > -0.03
        else "STRONG SELL"
    )
    ws5.set_cell_value(f"F{r}", signal)
    ws5.cell(f"F{r}").set_style_id(S_BODY_CENTER_ALT if alt else S_BODY_CENTER)

# Conditional formatting on P&L column
movers_start = attr_end + 5
movers_end = movers_start + 9
cf_movers = XlsxConditionalFormatting(
    "expression", [f"C{movers_start}:C{movers_end}"], [f"C{movers_start}>0"]
)
ws5.add_conditional_formatting(cf_movers)

# ── Right side: charts ──
# Bar chart of sector weights
bar_chart = XlsxChart("bar")
bar_chart.title = "Sector Weight Distribution"
bar_chart.set_anchor(6, 9, 12, 24)
bar_s = XlsxChartSeries(0, 0)
dashboard_cat_ref = XlsxChartDataRef.from_formula(f"Dashboard!$A$12:$A${attr_end - 1}")
dashboard_cat_ref.str_values = sector_order
bar_s.set_categories(dashboard_cat_ref)

dashboard_val_ref = XlsxChartDataRef.from_formula(f"Dashboard!$B$12:$B${attr_end - 1}")
dashboard_val_ref.num_values = dashboard_sector_weights
bar_s.set_values(dashboard_val_ref)
bar_chart.add_series(bar_s)
ws5.add_chart(bar_chart)

# Page setup
setup = XlsxPageSetup()
setup.orientation = "landscape"
setup.paper_size = 1
setup.fit_to_width = 1
setup.fit_to_height = 1
ws5.set_page_setup(setup)

margins = XlsxPageMargins()
margins.top = 0.5
margins.bottom = 0.5
margins.left = 0.4
margins.right = 0.4
ws5.set_page_margins(margins)

hf = XlsxPrintHeaderFooter()
hf.odd_header = "&C&B MERIDIAN CAPITAL PARTNERS — CONFIDENTIAL"
hf.odd_footer = "&L&D&RPage &P of &N"
ws5.set_header_footer(hf)

view5 = XlsxSheetViewOptions()
view5.show_gridlines = False
view5.zoom_scale = 85
view5.tab_selected = True
ws5.set_sheet_view_options(view5)

# Sheet protection — read-only dashboard
prot = XlsxSheetProtection()
prot.sheet = True
prot.format_cells = True
prot.sort = False
prot.auto_filter = False
ws5.set_protection_detail(prot)


# ═════════════════════════════════════════════════════════════════════════════
#  DEFINED NAMES
# ═════════════════════════════════════════════════════════════════════════════
wb.add_defined_name("PortfolioNAV", "Dashboard!$B$6")
wb.add_defined_name("TotalExposure", f"Holdings!$G${total_row}")
wb.add_defined_name("DailyReturns", f"'Daily Returns'!$B$4:$B${last_data_row}")


# ═════════════════════════════════════════════════════════════════════════════
#  SAVE
# ═════════════════════════════════════════════════════════════════════════════

OUT = tempfile.mkdtemp(prefix="offidized_quant_")
path = os.path.join(OUT, "meridian_portfolio.xlsx")
wb.save(path)

print()
print(f"Saved: {path}")
print()
print(
    f"  5 sheets  |  {len(positions)} positions  |  {DAYS} days returns  |  {len(trades)} trades"
)
print(
    f"  Styles: {S_TOTAL_INT + 1} registered  |  Charts: 3  |  Sparklines: {n_sectors}"
)
print("  Features: formulas, conditional formatting, data validation,")
print("            sparklines, charts, freeze panes, merged cells,")
print("            comments, defined names, page setup, sheet protection")
