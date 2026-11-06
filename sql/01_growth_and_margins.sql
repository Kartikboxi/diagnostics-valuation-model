-- Q1. How fast is each company growing, and are margins expanding?
-- Uses LAG() to compare each year with the previous year for the same company.

SELECT
    company,
    fiscal_year,
    revenue,
    ROUND(100.0 * (revenue - LAG(revenue) OVER w) / LAG(revenue) OVER w, 1) AS revenue_yoy_pct,
    ROUND(100.0 * ebitda / revenue, 1)                                       AS ebitda_margin_pct,
    ROUND(100.0 * ebitda / revenue
          - LAG(100.0 * ebitda / revenue) OVER w, 1)                         AS margin_change_pp,
    ROUND(100.0 * net_profit / revenue, 1)                                   AS pat_margin_pct
FROM income_statement
WINDOW w AS (PARTITION BY company ORDER BY fiscal_year)
ORDER BY company, fiscal_year;
