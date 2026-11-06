-- Q3. Do reported profits turn into cash?
-- Cash conversion = cash from operations / EBITDA, summed over FY23-FY26
-- so one unusual year does not distort the picture.

SELECT
    i.company,
    SUM(c.cfo)                                AS cfo_4yr,
    SUM(i.ebitda)                             AS ebitda_4yr,
    ROUND(100.0 * SUM(c.cfo) / SUM(i.ebitda), 1) AS cash_conversion_pct,
    CASE
        WHEN 1.0 * SUM(c.cfo) / SUM(i.ebitda) >= 0.80 THEN 'Strong'
        WHEN 1.0 * SUM(c.cfo) / SUM(i.ebitda) >= 0.65 THEN 'Adequate'
        ELSE 'Weak'
    END                                       AS rating
FROM income_statement i
JOIN cash_flow c
  ON c.company = i.company AND c.fiscal_year = i.fiscal_year
GROUP BY i.company
ORDER BY cash_conversion_pct DESC;
