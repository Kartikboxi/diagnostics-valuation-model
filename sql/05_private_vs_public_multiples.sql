-- Q5. How do prices paid for small private chains compare with listed multiples?
-- Only deals that disclosed both price and target revenue give a multiple.

SELECT
    'Precedent deal: ' || target                       AS benchmark,
    announced,
    ROUND(deal_value_cr / target_revenue_cr, 1)        AS ev_revenue
FROM precedent_transactions
WHERE deal_value_cr IS NOT NULL AND target_revenue_cr IS NOT NULL

UNION ALL

SELECT
    'Listed peer average (LTM)',
    MAX(m.price_date),
    ROUND(AVG((m.market_cap_cr + b.borrowings - b.investments) / m.ltm_revenue), 1)
FROM market_data m
JOIN balance_sheet b ON b.company = m.company AND b.as_of = 'Sep-2026';
