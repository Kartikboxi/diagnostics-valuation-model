# Indian Diagnostics: Peer Benchmarking, Valuation & Deal Screening

A financial analysis of India's four listed diagnostics chains (Dr Lal PathLabs, Metropolis Healthcare, Thyrocare Technologies, Vijaya Diagnostic Centre) built in **SQL** and **Excel**, using their published consolidated financials for FY22 to FY26 and market prices from November 2026.

The question it answers: **how do these companies compare on growth, margins, returns and cash generation, what are they worth, and what does that tell a buyer about pricing a private diagnostics target?**

## What's inside

| Part | What it does |
|---|---|
| `data/` | Income statement, balance sheet, cash flow, market data and precedent deals as CSV (Rs crore) |
| `sql/` | SQLite schema and 5 analysis queries (window functions, CTEs, joins, CASE) |
| `scripts/run_sql.py` | Loads the CSVs into SQLite and runs every query, saving results to `outputs/sql_results/` |
| `scripts/build_model.py` | Builds the Excel model with live formulas |
| `outputs/Diagnostics_Valuation_Model.xlsx` | The model: Ratios, Comps, Precedents, Screening scorecard and a 10-year DCF |

### The Excel model

- **Ratios**: revenue CAGR, EBITDA and PAT margins, ROE, ROCE, net debt/EBITDA and cash conversion, all pulled from the input tabs with `SUMIFS`
- **Comps**: enterprise value and EV/Revenue, EV/EBITDA and P/E on last-twelve-month numbers, with mean, median and quartiles, plus an implied valuation range for a private target
- **Precedents**: disclosed acquisitions of private diagnostics chains (2025-26) and the multiple paid
- **Screening**: weighted scorecard that normalises six metrics to 0-1 and ranks the peers; weights are editable
- **DCF**: pick any company from a drop-down; 10-year FCFF forecast, WACC build-up, terminal value, WACC x growth sensitivity table, and a reverse DCF showing the growth the market price implies

Inputs are blue on yellow; everything else is a formula, so every number can be traced back to its source row.

## Key findings

**Operating benchmarks (FY26)**

| | Revenue (Rs cr) | Revenue CAGR FY23-26 | EBITDA margin | ROCE | Cash conversion FY23-26 |
|---|---|---|---|---|---|
| Dr Lal PathLabs | 2,763 | 11.1% | 28.3% | 24.7% | 85% |
| Metropolis Healthcare | 1,646 | 12.8% | 24.7% | 16.6% | 87% |
| Thyrocare Technologies | 829 | 16.3% | 31.7% | 33.8% | 98% |
| Vijaya Diagnostic Centre | 814 | 21.0% | 41.4% | 19.6% | 83% |

**Valuation (LTM to Sep-2026, prices 5-6 Nov 2026)**

| | EV (Rs cr) | EV/Revenue | EV/EBITDA | P/E |
|---|---|---|---|---|
| Dr Lal PathLabs | 32,618 | 10.8x | 36.5x | 56.0x |
| Metropolis Healthcare | 11,805 | 6.6x | 26.4x | 52.5x |
| Thyrocare Technologies | 7,717 | 8.3x | 25.8x | 54.7x |
| Vijaya Diagnostic Centre | 15,010 | 16.5x | 38.2x | 72.6x |

1. **Vijaya is the highest-quality grower but the most expensive.** It has the fastest growth and margins over 40%, but it trades at 38x EV/EBITDA and is the only peer carrying net debt (0.58x EBITDA) as it funds new centres.
2. **Thyrocare ranks first on the default scorecard (0.80 vs 0.44 for second place).** It has the best ROCE (34%) and cash conversion (98%), and the lowest EV/EBITDA (25.8x).
3. **Metropolis has the weakest returns (ROCE 16.6%).** It made three acquisitions in 2025 (Core Diagnostics for Rs 247 cr, Scientific Pathology and Dr Ahujas'), which lifted FY26 revenue 24% but have not yet lifted returns.
4. **Private chains sell far cheaper than listed ones.** Dr Ahujas' Pathology went for about 3.0x revenue in April 2025, against an average of 10.5x for the listed peers. For an illustrative private target with Rs 120 cr revenue and a 25% margin, discounted listed multiples imply Rs 511-720 cr, while the precedent implies about Rs 365 cr. That gap is the main point of negotiation in any deal.
5. **Market prices assume very long, very high growth.** With a 12-12.4% WACC and 5% terminal growth, the 10-year DCF values all four companies 62-78% below their share prices. The reverse DCF shows the current prices imply about 10.5-11% growth *forever*, close to the cost of capital itself.

## How to run

```bash
python scripts/run_sql.py        # builds outputs/diagnostics.db and runs the SQL queries
python scripts/build_model.py    # builds the Excel model (needs openpyxl; LibreOffice optional for recalculation)
```

## Assumptions and limitations

- Financials are consolidated figures from screener.in, rounded to Rs crore. For a few balance sheet dates, other liabilities and CWIP are derived as the balancing figure.
- Net debt = borrowings - investments. The source does not split out cash, so investments stand in for cash and net debt is slightly overstated.
- DCF inputs (beta 0.85, equity risk premium 6.5%, 5% terminal growth, capex = D&A + 2% of revenue) are my assumptions and are all editable in the model. The risk-free rate is the 10-year G-sec yield of about 6.9% (Aug 2026).
- Only one precedent deal disclosed target revenue, so the private-market multiple rests on a single data point.

## Sources

- Financials and prices: screener.in consolidated pages for [LALPATHLAB](https://www.screener.in/company/LALPATHLAB/consolidated/), [METROPOLIS](https://www.screener.in/company/METROPOLIS/consolidated/), [THYROCARE](https://www.screener.in/company/THYROCARE/consolidated/), [VIJAYA](https://www.screener.in/company/VIJAYA/consolidated/)
- Precedent deals: links in `data/precedent_transactions.csv`
- 10-year G-sec yield: [Trading Economics](https://tradingeconomics.com/india/government-bond-yield)

*Self-directed learning project. Not investment advice.*
