-- Q4. Trading comparables on last-twelve-month (LTM) numbers.
-- Net debt uses the latest balance sheet (Sep-2026):
--   net debt = borrowings - investments  (investments used as a proxy for cash,
--   because the source does not split out cash; this slightly overstates net debt).
-- EV = market cap + net debt.

WITH latest_bs AS (
    SELECT company, borrowings - investments AS net_debt
    FROM balance_sheet
    WHERE as_of = 'Sep-2026'
)
SELECT
    m.company,
    m.market_cap_cr,
    b.net_debt,
    m.market_cap_cr + b.net_debt                                    AS enterprise_value,
    ROUND((m.market_cap_cr + b.net_debt) / m.ltm_revenue, 1)        AS ev_revenue,
    ROUND((m.market_cap_cr + b.net_debt) / m.ltm_ebitda, 1)         AS ev_ebitda,
    ROUND(m.market_cap_cr / m.ltm_net_profit, 1)                    AS pe,
    RANK() OVER (ORDER BY (m.market_cap_cr + b.net_debt) / m.ltm_ebitda) AS cheapest_rank
FROM market_data m
JOIN latest_bs b ON b.company = m.company
ORDER BY ev_ebitda;
