-- Q2. Who earns the best return on the capital they use? (FY26)
-- ROE  = net profit / average shareholders' equity
-- ROCE = EBIT / average capital employed (equity + borrowings)
-- Averages use the opening (Mar-2025) and closing (Mar-2026) balance sheets.

WITH bs AS (
    SELECT
        company,
        AVG(equity_capital + reserves)              AS avg_equity,
        AVG(equity_capital + reserves + borrowings) AS avg_capital_employed
    FROM balance_sheet
    WHERE as_of IN ('Mar-2025', 'Mar-2026')
    GROUP BY company
),
pl AS (
    SELECT company, net_profit, ebitda - depreciation AS ebit
    FROM income_statement
    WHERE fiscal_year = 'FY26'
)
SELECT
    pl.company,
    ROUND(100.0 * pl.net_profit / bs.avg_equity, 1)     AS roe_pct,
    ROUND(100.0 * pl.ebit / bs.avg_capital_employed, 1) AS roce_pct
FROM pl
JOIN bs ON bs.company = pl.company
ORDER BY roce_pct DESC;
