-- Schema for listed Indian diagnostics peer data.
-- All money values are in Rs crore unless the column name says otherwise.

DROP TABLE IF EXISTS income_statement;
DROP TABLE IF EXISTS balance_sheet;
DROP TABLE IF EXISTS cash_flow;
DROP TABLE IF EXISTS market_data;
DROP TABLE IF EXISTS precedent_transactions;

CREATE TABLE income_statement (
    company        TEXT NOT NULL,
    fiscal_year    TEXT NOT NULL,      -- FY26 = year ending 31 Mar 2026
    revenue        REAL,
    ebitda         REAL,               -- operating profit before other income
    other_income   REAL,
    interest       REAL,
    depreciation   REAL,
    pbt            REAL,
    net_profit     REAL,
    eps            REAL,
    PRIMARY KEY (company, fiscal_year)
);

CREATE TABLE balance_sheet (
    company            TEXT NOT NULL,
    as_of              TEXT NOT NULL,  -- Mar-2026, Sep-2026 ...
    equity_capital     REAL,
    reserves           REAL,
    borrowings         REAL,
    other_liabilities  REAL,
    fixed_assets       REAL,
    cwip               REAL,
    investments        REAL,
    other_assets       REAL,
    total_assets       REAL,
    PRIMARY KEY (company, as_of)
);

CREATE TABLE cash_flow (
    company      TEXT NOT NULL,
    fiscal_year  TEXT NOT NULL,
    cfo          REAL,                 -- cash from operating activities
    cfi          REAL,                 -- cash from investing activities
    PRIMARY KEY (company, fiscal_year)
);

CREATE TABLE market_data (
    company           TEXT PRIMARY KEY,
    ticker            TEXT,
    price_inr         REAL,
    market_cap_cr     REAL,
    price_date        TEXT,
    ltm_revenue       REAL,
    ltm_ebitda        REAL,
    ltm_depreciation  REAL,
    ltm_net_profit    REAL
);

CREATE TABLE precedent_transactions (
    announced          TEXT,
    acquirer           TEXT,
    target             TEXT,
    target_city        TEXT,
    deal_value_cr      REAL,
    target_revenue_cr  REAL,
    consideration      TEXT,
    source_url         TEXT
);
