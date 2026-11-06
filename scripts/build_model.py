"""Build the Excel valuation model (outputs/Diagnostics_Valuation_Model.xlsx).

Every number on the analysis sheets is an Excel formula that points back to the
Inputs sheets, so the model can be audited cell by cell and updated by pasting
new data into the Inputs tabs.

Usage: python scripts/build_model.py
"""
import csv
import shutil
import subprocess
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = ROOT / "outputs" / "Diagnostics_Valuation_Model.xlsx"

# ---------- styles ----------
BLUE = Font(color="0000CC")                      # hard-coded input
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
WHITE_BOLD = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
KEY_FILL = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="999999")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

CR = '#,##0;(#,##0)'
CR1 = '#,##0.0;(#,##0.0)'
PCT = '0.0%;(0.0%)'
MULT = '0.0"x"'
RS = '"Rs "#,##0'

COMPANIES = ["Dr Lal PathLabs", "Metropolis Healthcare", "Thyrocare Technologies", "Vijaya Diagnostic Centre"]


def read_csv(name):
    with open(DATA / name, newline="") as f:
        return list(csv.reader(f))


def num(v):
    if v == "":
        return None
    try:
        return float(v)
    except ValueError:
        return v


def header_row(ws, row, values, start_col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = WHITE_BOLD
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def input_cell(cell, value, fmt=None):
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL
    cell.border = BOX
    if fmt:
        cell.number_format = fmt


def write_table(ws, rows, money_cols=()):
    """Write a CSV table (header + rows) as inputs, starting at A1."""
    header_row(ws, 1, rows[0])
    for r, row in enumerate(rows[1:], start=2):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=num(v))
            cell.font = BLUE
            cell.border = BOX
            if c in money_cols:
                cell.number_format = CR1 if isinstance(cell.value, float) and cell.value % 1 else CR
    ws.freeze_panes = "A2"
    return len(rows) - 1


# ---------- helpers that return SUMIFS formulas ----------
def pl(col, company_ref, year):
    """Value from Inputs_PL for a company/year. col is the column letter of the metric."""
    return f'SUMIFS(Inputs_PL!${col}$2:${col}$21,Inputs_PL!$A$2:$A$21,{company_ref},Inputs_PL!$B$2:$B$21,{year})'


def bs(col, company_ref, as_of):
    return f'SUMIFS(Inputs_BS!${col}$2:${col}$17,Inputs_BS!$A$2:$A$17,{company_ref},Inputs_BS!$B$2:$B$17,{as_of})'


def cf(col, company_ref, year):
    return f'SUMIFS(Inputs_CF!${col}$2:${col}$17,Inputs_CF!$A$2:$A$17,{company_ref},Inputs_CF!$B$2:$B$17,{year})'


def mkt(col, company_ref):
    return f'SUMIFS(Inputs_Market!${col}$2:${col}$5,Inputs_Market!$A$2:$A$5,{company_ref})'


# Column letters in the input sheets
PL_COL = dict(revenue="C", ebitda="D", other_income="E", interest="F", depreciation="G", pbt="H", net_profit="I", eps="J")
BS_COL = dict(equity_capital="C", reserves="D", borrowings="E", other_liabilities="F", fixed_assets="G",
              cwip="H", investments="I", other_assets="J", total_assets="K")
CF_COL = dict(cfo="C", cfi="D")
MK_COL = dict(ticker="B", price="C", mcap="D", price_date="E", ltm_revenue="F", ltm_ebitda="G",
              ltm_depreciation="H", ltm_net_profit="I")


def build_readme(wb):
    ws = wb.active
    ws.title = "README"
    lines = [
        ("Indian Diagnostics: Peer Benchmarking, Valuation & Deal Screening Model", TITLE),
        ("", None),
        ("Question this model answers", BOLD),
        ("How do India's four listed diagnostics chains compare on growth, margins, returns and cash conversion,", None),
        ("what are they worth on trading multiples and on a DCF, and what does that imply for valuing a private target?", None),
        ("", None),
        ("Sheets", BOLD),
        ("Inputs_PL / Inputs_BS / Inputs_CF / Inputs_Market : consolidated financials (Rs crore) and market data. Blue = hard-coded input.", None),
        ("Ratios : growth, margin, ROE, ROCE, leverage and cash conversion, all pulled from the inputs with SUMIFS.", None),
        ("Comps : enterprise value and EV/Revenue, EV/EBITDA, P/E on last-twelve-month numbers, plus an implied valuation for a private target.", None),
        ("Precedents : disclosed acquisitions of private diagnostics chains and the multiple paid where revenue was disclosed.", None),
        ("Screening : weighted scorecard that ranks the peers; change the weights in yellow to test different priorities.", None),
        ("DCF : 10-year free cash flow model. Pick a company from the drop-down in C3; all assumptions are in yellow.", None),
        ("", None),
        ("Conventions", BOLD),
        ("Units: Rs crore unless stated. FY26 = year ending 31 March 2026. LTM = last twelve months to Sep-2026.", None),
        ("Blue font on yellow = input you can change. Black font = formula. Green fill = key output.", None),
        ("Net debt = borrowings - investments (cash is not reported separately in the source, so investments are used as the cash proxy).", None),
        ("", None),
        ("Sources", BOLD),
        ("Financials and prices: screener.in consolidated pages for LALPATHLAB, METROPOLIS, THYROCARE, VIJAYA (prices as of 5-6 Nov 2026).", None),
        ("Deals: company press releases and news coverage, links in the Precedents sheet.", None),
        ("Risk-free rate: India 10-year G-sec yield, about 6.9% (Trading Economics, Aug 2026).", None),
        ("", None),
        ("Built by Kartik Kumar Boxi as a self-directed learning project. Not investment advice.", Font(italic=True)),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=text)
        if font:
            c.font = font
    ws.column_dimensions["A"].width = 130


def build_inputs(wb):
    ws = wb.create_sheet("Inputs_PL")
    write_table(ws, read_csv("income_statement.csv"), money_cols=range(3, 11))
    widths(ws, {"A": 26, "B": 11, **{c: 12 for c in "CDEFGHIJ"}})

    ws = wb.create_sheet("Inputs_BS")
    write_table(ws, read_csv("balance_sheet.csv"), money_cols=range(3, 12))
    widths(ws, {"A": 26, "B": 11, **{c: 13 for c in "CDEFGHIJK"}})

    ws = wb.create_sheet("Inputs_CF")
    write_table(ws, read_csv("cash_flow.csv"), money_cols=range(3, 5))
    widths(ws, {"A": 26, "B": 11, "C": 12, "D": 12})

    ws = wb.create_sheet("Inputs_Market")
    write_table(ws, read_csv("market_data.csv"), money_cols=[3, 4, 6, 7, 8, 9])
    widths(ws, {"A": 26, "B": 14, **{c: 13 for c in "CDEFGHI"}})


def build_ratios(wb):
    ws = wb.create_sheet("Ratios")
    ws["A1"] = "Peer ratios (FY26 unless stated)"
    ws["A1"].font = TITLE
    ws["A2"] = "Section A pulls raw values with SUMIFS; section B turns them into ratios."

    # Section A: pulled values
    ws["A4"] = "A. Pulled values (Rs crore)"
    ws["A4"].font = BOLD
    cols_a = ["Company", "Revenue FY22", "Revenue FY23", "Revenue FY26", "EBITDA FY26", "D&A FY26",
              "Net profit FY26", "Equity Mar-25", "Equity Mar-26", "Borrowings Mar-25", "Borrowings Mar-26",
              "Investments Mar-26", "CFO FY23-26", "EBITDA FY23-26"]
    header_row(ws, 5, cols_a)
    for i, co in enumerate(COMPANIES):
        r = 6 + i
        a = f"$A{r}"
        ws[f"A{r}"] = co
        ws[f"B{r}"] = "=" + pl(PL_COL["revenue"], a, '"FY22"')
        ws[f"C{r}"] = "=" + pl(PL_COL["revenue"], a, '"FY23"')
        ws[f"D{r}"] = "=" + pl(PL_COL["revenue"], a, '"FY26"')
        ws[f"E{r}"] = "=" + pl(PL_COL["ebitda"], a, '"FY26"')
        ws[f"F{r}"] = "=" + pl(PL_COL["depreciation"], a, '"FY26"')
        ws[f"G{r}"] = "=" + pl(PL_COL["net_profit"], a, '"FY26"')
        ws[f"H{r}"] = "=" + bs(BS_COL["equity_capital"], a, '"Mar-2025"') + "+" + bs(BS_COL["reserves"], a, '"Mar-2025"')
        ws[f"I{r}"] = "=" + bs(BS_COL["equity_capital"], a, '"Mar-2026"') + "+" + bs(BS_COL["reserves"], a, '"Mar-2026"')
        ws[f"J{r}"] = "=" + bs(BS_COL["borrowings"], a, '"Mar-2025"')
        ws[f"K{r}"] = "=" + bs(BS_COL["borrowings"], a, '"Mar-2026"')
        ws[f"L{r}"] = "=" + bs(BS_COL["investments"], a, '"Mar-2026"')
        ws[f"M{r}"] = f'=SUMIFS(Inputs_CF!$C$2:$C$17,Inputs_CF!$A$2:$A$17,{a})'
        ws[f"N{r}"] = (f'=SUMIFS(Inputs_PL!$D$2:$D$21,Inputs_PL!$A$2:$A$21,{a})-'
                       + pl(PL_COL["ebitda"], a, '"FY22"'))
        for col in "BCDEFGHIJKLMN":
            ws[f"{col}{r}"].number_format = CR

    # Section B: ratios
    ws["A12"] = "B. Ratios"
    ws["A12"].font = BOLD
    cols_b = ["Company", "Revenue CAGR FY22-26", "Revenue CAGR FY23-26", "EBITDA margin", "PAT margin",
              "ROE", "ROCE", "Net debt / EBITDA", "Cash conversion (CFO/EBITDA, FY23-26)"]
    header_row(ws, 13, cols_b)
    for i, co in enumerate(COMPANIES):
        r = 14 + i
        s = 6 + i  # matching row in section A
        ws[f"A{r}"] = f"=A{s}"
        ws[f"B{r}"] = f"=(D{s}/B{s})^(1/4)-1"
        ws[f"C{r}"] = f"=(D{s}/C{s})^(1/3)-1"
        ws[f"D{r}"] = f"=E{s}/D{s}"
        ws[f"E{r}"] = f"=G{s}/D{s}"
        ws[f"F{r}"] = f"=G{s}/AVERAGE(H{s},I{s})"
        ws[f"G{r}"] = f"=(E{s}-F{s})/AVERAGE(H{s}+J{s},I{s}+K{s})"
        ws[f"H{r}"] = f"=(K{s}-L{s})/E{s}"
        ws[f"I{r}"] = f"=M{s}/N{s}"
        for col in "BCDEFGI":
            ws[f"{col}{r}"].number_format = PCT
        ws[f"H{r}"].number_format = '0.00"x";(0.00"x")'
    r = 18
    ws[f"A{r}"] = "Peer median"
    ws[f"A{r}"].font = BOLD
    for col in "BCDEFGHI":
        ws[f"{col}{r}"] = f"=MEDIAN({col}14:{col}17)"
        ws[f"{col}{r}"].number_format = ws[f"{col}14"].number_format
        ws[f"{col}{r}"].font = BOLD
    ws["A20"] = "ROE = net profit / average equity. ROCE = (EBITDA - D&A) / average (equity + borrowings). Negative net debt = net cash."
    widths(ws, {"A": 26, **{c: 15 for c in "BCDEFGHIJKLMN"}})
    ws.row_dimensions[5].height = 32
    ws.row_dimensions[13].height = 45
    ws.freeze_panes = "B6"


def build_comps(wb):
    ws = wb.create_sheet("Comps")
    ws["A1"] = "Trading comparables (LTM to Sep-2026, prices 5-6 Nov 2026)"
    ws["A1"].font = TITLE
    cols = ["Company", "Price (Rs)", "Market cap", "Shares (cr)", "Borrowings Sep-26", "Investments Sep-26",
            "Net debt", "Enterprise value", "LTM revenue", "LTM EBITDA", "LTM net profit",
            "EV / Revenue", "EV / EBITDA", "P / E"]
    header_row(ws, 3, cols)
    for i, co in enumerate(COMPANIES):
        r = 4 + i
        a = f"$A{r}"
        ws[f"A{r}"] = co
        ws[f"B{r}"] = "=" + mkt(MK_COL["price"], a)
        ws[f"C{r}"] = "=" + mkt(MK_COL["mcap"], a)
        ws[f"D{r}"] = f"=C{r}/B{r}"
        ws[f"E{r}"] = "=" + bs(BS_COL["borrowings"], a, '"Sep-2026"')
        ws[f"F{r}"] = "=" + bs(BS_COL["investments"], a, '"Sep-2026"')
        ws[f"G{r}"] = f"=E{r}-F{r}"
        ws[f"H{r}"] = f"=C{r}+G{r}"
        ws[f"I{r}"] = "=" + mkt(MK_COL["ltm_revenue"], a)
        ws[f"J{r}"] = "=" + mkt(MK_COL["ltm_ebitda"], a)
        ws[f"K{r}"] = "=" + mkt(MK_COL["ltm_net_profit"], a)
        ws[f"L{r}"] = f"=H{r}/I{r}"
        ws[f"M{r}"] = f"=H{r}/J{r}"
        ws[f"N{r}"] = f"=C{r}/K{r}"
        for col in "BCEFGHIJK":
            ws[f"{col}{r}"].number_format = CR
        ws[f"D{r}"].number_format = '0.00'
        for col in "LMN":
            ws[f"{col}{r}"].number_format = MULT

    stats = [("Mean", "AVERAGE({c}4:{c}7)"), ("Median", "MEDIAN({c}4:{c}7)"),
             ("25th percentile", "QUARTILE({c}4:{c}7,1)"), ("75th percentile", "QUARTILE({c}4:{c}7,3)")]
    for j, (label, f) in enumerate(stats):
        r = 9 + j
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = BOLD
        for col in "LMN":
            ws[f"{col}{r}"] = "=" + f.format(c=col)
            ws[f"{col}{r}"].number_format = MULT
            ws[f"{col}{r}"].font = BOLD

    # Implied valuation for a private target
    ws["A15"] = "Implied valuation of a private target (illustrative)"
    ws["A15"].font = TITLE
    ws["A16"] = "Listed multiples carry a liquidity and scale premium, so a discount is applied for a smaller private business."
    rows = [
        ("Target LTM revenue (Rs cr)", 120, CR),
        ("Target EBITDA margin", 0.25, PCT),
        ("Private-company discount to listed multiples", 0.35, PCT),
    ]
    for j, (label, val, fmt) in enumerate(rows):
        r = 18 + j
        ws[f"A{r}"] = label
        input_cell(ws[f"B{r}"], val, fmt)
    ws["A21"] = "Target LTM EBITDA (Rs cr)"
    ws["B21"] = "=B18*B19"
    ws["B21"].number_format = CR1

    header_row(ws, 23, ["Method", "Multiple used", "Implied EV (Rs cr)"])
    methods = [
        ("EV/EBITDA, 25th percentile, discounted", "=M11*(1-$B$20)", "=B24*$B$21"),
        ("EV/EBITDA, median, discounted", "=M10*(1-$B$20)", "=B25*$B$21"),
        ("EV/EBITDA, 75th percentile, discounted", "=M12*(1-$B$20)", "=B26*$B$21"),
        ("EV/Revenue, precedent deal (Dr Ahujas', Apr-2025)", "=Precedents!H6", "=B27*$B$18"),
    ]
    for j, (label, mult, ev) in enumerate(methods):
        r = 24 + j
        ws[f"A{r}"] = label
        ws[f"B{r}"] = mult
        ws[f"B{r}"].number_format = MULT
        ws[f"C{r}"] = ev
        ws[f"C{r}"].number_format = CR
        ws[f"C{r}"].fill = KEY_FILL
    ws["A29"] = ("Read-across: small private chains have changed hands at about 3x revenue while listed peers trade at "
                 "about 7-17x, so the price a buyer should pay depends mostly on which benchmark is defensible.")
    widths(ws, {"A": 48, **{c: 14 for c in "BCDEFGHIJKLMN"}})
    ws.row_dimensions[3].height = 32
    ws.freeze_panes = "B4"


def build_precedents(wb):
    ws = wb.create_sheet("Precedents")
    ws["A1"] = "Precedent transactions: acquisitions of private diagnostics chains"
    ws["A1"].font = TITLE
    rows = read_csv("precedent_transactions.csv")
    header = ["Announced", "Acquirer", "Target", "City", "Deal value (Rs cr)", "Target revenue (Rs cr)",
              "Consideration", "EV / Revenue", "Source"]
    header_row(ws, 3, header)
    for j, row in enumerate(rows[1:]):
        r = 4 + j
        d = dict(zip(rows[0], row))
        vals = [d["announced"], d["acquirer"], d["target"], d["target_city"], num(d["deal_value_cr"]),
                num(d["target_revenue_cr"]), d["consideration"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BLUE
        ws[f"E{r}"].number_format = CR1
        ws[f"F{r}"].number_format = CR1
        ws[f"H{r}"] = f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r})),E{r}/F{r},"n/a")'
        ws[f"H{r}"].number_format = MULT
        ws[f"I{r}"] = d["source_url"]
    ws["A10"] = "Only deals that disclosed both price and target revenue produce a multiple; the rest are kept to show deal activity."
    widths(ws, {"A": 12, "B": 22, "C": 36, "D": 12, "E": 14, "F": 14, "G": 34, "H": 12, "I": 60})
    # Precedents must sit before Comps in formulas; order of sheets does not matter for references.


def build_screening(wb):
    ws = wb.create_sheet("Screening")
    ws["A1"] = "Peer screening scorecard"
    ws["A1"].font = TITLE
    ws["A2"] = "Each metric is rescaled 0-1 across the four peers (1 = best). Weighted score = SUMPRODUCT(scores, weights)."

    metrics = [  # label, source formula template (row r of Ratios/Comps), higher_is_better
        ("Revenue CAGR FY23-26", "Ratios!C{r}", True),
        ("EBITDA margin", "Ratios!D{r}", True),
        ("ROCE", "Ratios!G{r}", True),
        ("Cash conversion", "Ratios!I{r}", True),
        ("Net debt / EBITDA", "Ratios!H{r}", False),
        ("EV / EBITDA (LTM)", "Comps!M{c}", False),
    ]
    weights = [0.25, 0.15, 0.20, 0.10, 0.10, 0.20]

    ws["A4"] = "Weights"
    ws["A4"].font = BOLD
    header_row(ws, 5, [m[0] for m in metrics] + ["Total"], start_col=2)
    for j, w in enumerate(weights):
        input_cell(ws.cell(row=6, column=2 + j), w, PCT)
    ws["H6"] = "=SUM(B6:G6)"
    ws["H6"].number_format = PCT
    ws["I6"] = '=IF(ROUND(H6,4)=1,"OK","Weights must add to 100%")'

    ws["A8"] = "Raw metrics"
    ws["A8"].font = BOLD
    header_row(ws, 9, ["Company"] + [m[0] for m in metrics])
    for i, co in enumerate(COMPANIES):
        r = 10 + i
        ws[f"A{r}"] = co
        for j, (_, src, _) in enumerate(metrics):
            cell = ws.cell(row=r, column=2 + j, value="=" + src.format(r=14 + i, c=4 + i))
            cell.number_format = MULT if "EV" in metrics[j][0] or "Net debt" in metrics[j][0] else PCT

    ws["A15"] = "Scores (0 = worst, 1 = best)"
    ws["A15"].font = BOLD
    header_row(ws, 16, ["Company"] + [m[0] for m in metrics] + ["Weighted score", "Rank"])
    for i, co in enumerate(COMPANIES):
        r = 17 + i
        raw = 10 + i
        ws[f"A{r}"] = co
        for j, (_, _, higher) in enumerate(metrics):
            col = chr(ord("B") + j)
            rng = f"{col}$10:{col}$13"
            if higher:
                f = f"=({col}{raw}-MIN({rng}))/(MAX({rng})-MIN({rng}))"
            else:
                f = f"=(MAX({rng})-{col}{raw})/(MAX({rng})-MIN({rng}))"
            ws[f"{col}{r}"] = f
            ws[f"{col}{r}"].number_format = "0.00"
        ws[f"H{r}"] = f"=SUMPRODUCT(B{r}:G{r},$B$6:$G$6)"
        ws[f"H{r}"].number_format = "0.00"
        ws[f"H{r}"].fill = KEY_FILL
        ws[f"I{r}"] = f"=RANK(H{r},$H$17:$H$20)"
    ws["A22"] = ("Default weights favour growth, returns and price paid. Shift weight to EBITDA margin or cash "
                 "conversion to see how the ranking changes for a buyer focused on profitability.")
    widths(ws, {"A": 26, **{c: 15 for c in "BCDEFGHI"}})
    for r in (5, 9, 16):
        ws.row_dimensions[r].height = 32


def build_dcf(wb):
    ws = wb.create_sheet("DCF")
    ws["A1"] = "Discounted cash flow valuation (10-year explicit forecast)"
    ws["A1"].font = TITLE

    ws["B3"] = "Company (pick from list)"
    ws["B3"].font = BOLD
    input_cell(ws["C3"], "Vijaya Diagnostic Centre")
    dv = DataValidation(type="list", formula1="=Inputs_Market!$A$2:$A$5", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C3")

    co = "$C$3"
    ws["B5"] = "Assumptions"
    ws["B5"].font = BOLD
    assumptions = [
        # row, label, value or formula, format, is_input
        (6, "Base year", "FY26", None, True),
        (7, "Tax rate (Indian corporate rate, new regime)", 0.2517, PCT, True),
        (8, "Risk-free rate (10-yr G-sec)", 0.069, PCT, True),
        (9, "Equity beta (assumed; diagnostics are defensive)", 0.85, "0.00", True),
        (10, "Equity risk premium (India)", 0.065, PCT, True),
        (11, "Cost of equity  = Rf + beta x ERP", "=C8+C9*C10", PCT, False),
        (12, "Pre-tax cost of debt", 0.09, PCT, True),
        (13, "Debt weight  = borrowings / (borrowings + market cap)",
         "=" + bs(BS_COL["borrowings"], co, '"Sep-2026"') + "/(" + bs(BS_COL["borrowings"], co, '"Sep-2026"')
         + "+" + mkt(MK_COL["mcap"], co) + ")", PCT, False),
        (14, "WACC", "=C11*(1-C13)+C12*(1-C7)*C13", PCT, False),
        (15, "Terminal growth rate", 0.05, PCT, True),
        (16, "Year-1 revenue growth  = FY23-26 CAGR (fades to terminal by year 10)",
         "=(" + pl(PL_COL["revenue"], co, "$C$6") + "/" + pl(PL_COL["revenue"], co, '"FY23"') + ")^(1/3)-1", PCT, False),
        (17, "EBITDA margin change per year (percentage points)", 0.0, PCT, True),
        (18, "D&A as % of revenue (base year)",
         "=" + pl(PL_COL["depreciation"], co, "$C$6") + "/" + pl(PL_COL["revenue"], co, "$C$6"), PCT, False),
        (19, "Growth capex on top of D&A (% of revenue)", 0.02, PCT, True),
        (20, "Capex as % of revenue  = D&A % + growth capex", "=C18+C19", PCT, False),
        (21, "Working capital invested (% of each extra rupee of revenue)", 0.10, PCT, True),
    ]
    for r, label, val, fmt, is_input in assumptions:
        ws[f"B{r}"] = label
        if is_input:
            input_cell(ws[f"C{r}"], val, fmt)
        else:
            ws[f"C{r}"] = val
            if fmt:
                ws[f"C{r}"].number_format = fmt
    ws["C14"].fill = KEY_FILL

    # Projection table: column D = base year, E..N = 10 forecast years
    years = ["FY26A"] + [f"FY{27 + k}E" for k in range(10)]
    cols = [chr(ord("D") + k) for k in range(11)]  # D..N
    header_row(ws, 23, ["Rs crore"] + years, start_col=3)
    labels = {24: "Year number", 25: "Revenue growth", 26: "Revenue", 27: "EBITDA margin", 28: "EBITDA",
              29: "Depreciation & amortisation", 30: "EBIT", 31: "Tax on EBIT", 32: "NOPAT", 33: "Capex",
              34: "Investment in working capital", 35: "Free cash flow to firm (FCFF)", 36: "Discount factor",
              37: "Present value of FCFF"}
    for r, lab in labels.items():
        ws[f"C{r}"] = lab
    for r in (26, 28, 35, 37):
        ws[f"C{r}"].font = BOLD

    base = cols[0]
    ws[f"{base}24"] = 0
    ws[f"{base}26"] = "=" + pl(PL_COL["revenue"], co, "$C$6")
    ws[f"{base}28"] = "=" + pl(PL_COL["ebitda"], co, "$C$6")
    ws[f"{base}27"] = f"={base}28/{base}26"
    ws[f"{base}29"] = "=" + pl(PL_COL["depreciation"], co, "$C$6")
    ws[f"{base}30"] = f"={base}28-{base}29"
    for k, col in enumerate(cols[1:], start=1):
        prev = cols[k - 1]
        ws[f"{col}24"] = k
        ws[f"{col}25"] = f"=$C$16-($C$16-$C$15)*({col}24-1)/9"
        ws[f"{col}26"] = f"={prev}26*(1+{col}25)"
        ws[f"{col}27"] = f"={prev}27+$C$17"
        ws[f"{col}28"] = f"={col}26*{col}27"
        ws[f"{col}29"] = f"={col}26*$C$18"
        ws[f"{col}30"] = f"={col}28-{col}29"
        ws[f"{col}31"] = f"={col}30*$C$7"
        ws[f"{col}32"] = f"={col}30-{col}31"
        ws[f"{col}33"] = f"={col}26*$C$20"
        ws[f"{col}34"] = f"=({col}26-{prev}26)*$C$21"
        ws[f"{col}35"] = f"={col}32+{col}29-{col}33-{col}34"
        ws[f"{col}36"] = f"=1/(1+$C$14)^{col}24"
        ws[f"{col}37"] = f"={col}35*{col}36"
    for col in cols:
        for r in range(26, 38):
            ws[f"{col}{r}"].number_format = CR
        for r in (25, 27):
            ws[f"{col}{r}"].number_format = PCT
        ws[f"{col}36"].number_format = "0.000"

    last = cols[-1]  # N
    ws["B39"] = "Valuation"
    ws["B39"].font = BOLD
    val_rows = [
        (40, "Sum of PV of FCFF (years 1-10)", "=SUM(E37:N37)", CR),
        (41, "Terminal value at year 10  = FCFF x (1+g) / (WACC - g)", f"={last}35*(1+C15)/(C14-C15)", CR),
        (42, "PV of terminal value", f"=C41*{last}36", CR),
        (43, "Enterprise value", "=C40+C42", CR),
        (44, "Less: net debt (Sep-2026)", "=" + bs(BS_COL["borrowings"], co, '"Sep-2026"') + "-" + bs(BS_COL["investments"], co, '"Sep-2026"'), CR),
        (45, "Equity value", "=C43-C44", CR),
        (46, "Shares outstanding (crore)", "=" + mkt(MK_COL["mcap"], co) + "/" + mkt(MK_COL["price"], co), "0.00"),
        (47, "DCF value per share (Rs)", "=C45/C46", RS),
        (48, "Current share price (Rs)", "=" + mkt(MK_COL["price"], co), RS),
        (49, "Upside / (downside) vs current price", "=C47/C48-1", PCT),
        (50, "Terminal value as % of enterprise value", "=C42/C43", PCT),
        (51, "Implied EV / LTM EBITDA at DCF value", "=C43/" + mkt(MK_COL["ltm_ebitda"], co), MULT),
        (52, "Reverse DCF: perpetual growth the market price implies  = WACC - FCFF(yr 1) / market EV",
         "=C14-E35/(" + mkt(MK_COL["mcap"], co) + "+C44)", PCT),
    ]
    for r, label, f, fmt in val_rows:
        ws[f"B{r}"] = label
        ws[f"C{r}"] = f
        ws[f"C{r}"].number_format = fmt
    for r in (43, 47, 49):
        ws[f"C{r}"].fill = KEY_FILL
        ws[f"B{r}"].font = BOLD

    # Sensitivity: value per share for WACC (rows) x terminal growth (columns)
    ws["B53"] = "Sensitivity: DCF value per share (Rs). Rows = WACC, columns = terminal growth"
    ws["B53"].font = BOLD
    g_steps = [-0.01, -0.005, 0, 0.005, 0.01]
    w_steps = [-0.02, -0.01, 0, 0.01, 0.02]
    ws["C54"] = "WACC \\ g"
    ws["C54"].font = BOLD
    for j, gs in enumerate(g_steps):
        col = chr(ord("D") + j)
        ws[f"{col}54"] = f"=$C$15+({gs})"
        ws[f"{col}54"].number_format = PCT
        ws[f"{col}54"].font = BOLD
    for i, wst in enumerate(w_steps):
        r = 55 + i
        ws[f"C{r}"] = f"=$C$14+({wst})"
        ws[f"C{r}"].number_format = PCT
        ws[f"C{r}"].font = BOLD
        for j in range(len(g_steps)):
            col = chr(ord("D") + j)
            ws[f"{col}{r}"] = (
                f"=(SUMPRODUCT($E$35:$N$35/(1+$C{r})^$E$24:$N$24)"
                f"+$N$35*(1+{col}$54)/($C{r}-{col}$54)/(1+$C{r})^10-$C$44)/$C$46"
            )
            ws[f"{col}{r}"].number_format = RS
    ws["F57"].fill = KEY_FILL  # base case cell (base WACC, base g)

    ws["B61"] = ("How to read: the market price is well above most DCF values because listed diagnostics stocks trade at "
                 "50-70x earnings, which assumes high growth for longer than 10 years or a lower discount rate.")
    widths(ws, {"A": 2, "B": 62, "C": 30, **{c: 11 for c in "DEFGHIJKLMN"}})
    ws.freeze_panes = "A4"


def main():
    wb = Workbook()
    build_readme(wb)
    build_inputs(wb)
    build_ratios(wb)
    build_comps(wb)
    build_precedents(wb)
    build_screening(wb)
    build_dcf(wb)
    wb.calculation.fullCalcOnLoad = True
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    recalculate(OUT)
    print(f"saved {OUT}")


def recalculate(path):
    """openpyxl writes formulas without results. If LibreOffice is installed, open and
    re-save the file so every cell also stores its calculated value (useful for viewers
    that do not recalculate). Excel recalculates on open either way."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice not found: values will be calculated when the file is opened in Excel")
        return
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, str(path)],
                       check=True, capture_output=True)
        shutil.move(Path(tmp) / path.name, path)


if __name__ == "__main__":
    main()
