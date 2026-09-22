-- ============================================================================
-- NovaPharma Iberia — Commercial Data Warehouse
-- 01_schema.sql : creates the database and the star/constellation schema
-- Designed by Nix, 18-Sep-2026. MySQL 8.
-- Run order: dimensions first (no dependencies), then facts (depend on dims).
-- ============================================================================

DROP DATABASE IF EXISTS novapharma_dw;
CREATE DATABASE novapharma_dw CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci;
USE novapharma_dw;

-- ----------------------------------------------------------------------------
-- DIMENSIONS
-- ----------------------------------------------------------------------------

CREATE TABLE dim_calendar (
    month_key       INT          NOT NULL,            -- 202401 (key, never summed)
    month_start     DATE         NOT NULL,
    month_name      VARCHAR(8)   NOT NULL,            -- 'Jan-2024'
    year            INT          NOT NULL,
    quarter         CHAR(2)      NOT NULL,            -- 'Q1'
    fiscal_period   VARCHAR(10)  NOT NULL,            -- 'FY2024-P01'
    is_closed       CHAR(1)      NOT NULL DEFAULT 'N',-- 'Y' / 'N'
    PRIMARY KEY (month_key)
) ENGINE=InnoDB;

CREATE TABLE dim_brand (
    brand_name        VARCHAR(30)  NOT NULL,          -- natural key (IQVIA file uses the name)
    therapeutic_area  VARCHAR(20)  NOT NULL,
    PRIMARY KEY (brand_name)
) ENGINE=InnoDB;

CREATE TABLE dim_product (
    sku_code           VARCHAR(10)   NOT NULL,
    sku_name           VARCHAR(60)   NOT NULL,
    brand_name         VARCHAR(30)   NOT NULL,        -- FK: SKU -> Brand -> TA hierarchy
    form_strength      VARCHAR(30),
    launch_date        DATE,
    loe_date           DATE,                          -- loss of exclusivity
    list_price_eur     DECIMAL(10,2),                 -- attribute, not a measure
    cogs_per_unit_eur  DECIMAL(10,2),                 -- attribute, not a measure
    status             VARCHAR(10),
    PRIMARY KEY (sku_code),
    CONSTRAINT fk_product_brand FOREIGN KEY (brand_name) REFERENCES dim_brand (brand_name)
) ENGINE=InnoDB;

CREATE TABLE dim_territory (
    territory_code   VARCHAR(5)   NOT NULL,
    territory_name   VARCHAR(30)  NOT NULL,
    region           VARCHAR(12)  NOT NULL,           -- implicit hierarchy:
    country          VARCHAR(12)  NOT NULL,           --   territory -> region -> country
    sales_rep        VARCHAR(40),
    rep_start_date   DATE,
    PRIMARY KEY (territory_code)
) ENGINE=InnoDB;

CREATE TABLE dim_customer (
    customer_code    VARCHAR(6)   NOT NULL,
    customer_name    VARCHAR(80)  NOT NULL,
    channel          VARCHAR(20)  NOT NULL,           -- Hospital / Retail-Wholesaler / Tender
    territory_code   VARCHAR(5)   NOT NULL,
    city             VARCHAR(30),                     -- attribute, not a hierarchy level
    PRIMARY KEY (customer_code),
    CONSTRAINT fk_customer_territory FOREIGN KEY (territory_code) REFERENCES dim_territory (territory_code)
) ENGINE=InnoDB;

CREATE TABLE dim_version (
    version_name   VARCHAR(20)  NOT NULL,             -- 'Budget 2026', 'FC 2026-Q2'
    version_type   VARCHAR(10)  NOT NULL,             -- 'Budget' / 'Forecast'
    plan_year      INT          NOT NULL,
    cutoff_month   INT          NULL,                 -- last month with actuals; NULL for Budget
    PRIMARY KEY (version_name)
) ENGINE=InnoDB;

-- dim_version is the only table filled by hand: this knowledge lived in people's heads in Excel
INSERT INTO dim_version (version_name, version_type, plan_year, cutoff_month) VALUES
    ('Budget 2024', 'Budget',   2024, NULL),
    ('Budget 2025', 'Budget',   2025, NULL),
    ('Budget 2026', 'Budget',   2026, NULL),
    ('FC 2026-Q1',  'Forecast', 2026, 202603),
    ('FC 2026-Q2',  'Forecast', 2026, 202606);

-- ----------------------------------------------------------------------------
-- FACTS  (one fact table per grain)
-- ----------------------------------------------------------------------------

-- Transactional fact: 1 row = 1 invoice line. Own ID as PK.
CREATE TABLE fact_sales (
    transaction_id    VARCHAR(12)    NOT NULL,
    invoice_date      DATE           NOT NULL,
    month_key         INT            NOT NULL,        -- derived in ETL: year*100+month
    sku_code          VARCHAR(10)    NOT NULL,
    customer_code     VARCHAR(6)     NOT NULL,
    units             INT            NOT NULL,        -- measure (negative = return)
    gross_sales_eur   DECIMAL(12,2)  NOT NULL,        -- measure
    discount_eur      DECIMAL(12,2)  NOT NULL DEFAULT 0,
    rebate_eur        DECIMAL(12,2)  NOT NULL DEFAULT 0,
    PRIMARY KEY (transaction_id),
    CONSTRAINT fk_sales_calendar FOREIGN KEY (month_key)     REFERENCES dim_calendar (month_key),
    CONSTRAINT fk_sales_product  FOREIGN KEY (sku_code)      REFERENCES dim_product  (sku_code),
    CONSTRAINT fk_sales_customer FOREIGN KEY (customer_code) REFERENCES dim_customer (customer_code),
    INDEX ix_sales_month_sku (month_key, sku_code),
    INDEX ix_sales_customer  (customer_code)
) ENGINE=InnoDB;

-- Snapshot fact: Budget + Forecast. 1 row per version x month x SKU x territory.
CREATE TABLE fact_plan (
    version_name     VARCHAR(20)    NOT NULL,
    month_key        INT            NOT NULL,
    sku_code         VARCHAR(10)    NOT NULL,
    territory_code   VARCHAR(5)     NOT NULL,
    units            INT            NOT NULL,         -- measure
    net_sales_eur    DECIMAL(12,2)  NOT NULL,         -- measure
    PRIMARY KEY (version_name, month_key, sku_code, territory_code),
    CONSTRAINT fk_plan_version   FOREIGN KEY (version_name)   REFERENCES dim_version   (version_name),
    CONSTRAINT fk_plan_calendar  FOREIGN KEY (month_key)      REFERENCES dim_calendar  (month_key),
    CONSTRAINT fk_plan_product   FOREIGN KEY (sku_code)       REFERENCES dim_product   (sku_code),
    CONSTRAINT fk_plan_territory FOREIGN KEY (territory_code) REFERENCES dim_territory (territory_code)
) ENGINE=InnoDB;

-- Snapshot fact: IQVIA market data. 1 row per month x brand x territory (brand level, not SKU).
CREATE TABLE fact_market (
    month_key         INT            NOT NULL,
    brand_name        VARCHAR(30)    NOT NULL,
    territory_code    VARCHAR(5)     NOT NULL,
    market_units      INT            NOT NULL,        -- measure
    market_value_eur  DECIMAL(12,2)  NOT NULL,        -- measure
    PRIMARY KEY (month_key, brand_name, territory_code),
    CONSTRAINT fk_market_calendar  FOREIGN KEY (month_key)      REFERENCES dim_calendar  (month_key),
    CONSTRAINT fk_market_brand     FOREIGN KEY (brand_name)     REFERENCES dim_brand     (brand_name),
    CONSTRAINT fk_market_territory FOREIGN KEY (territory_code) REFERENCES dim_territory (territory_code)
) ENGINE=InnoDB;

-- ----------------------------------------------------------------------------
-- Operations tables (added in Step 3.B)
-- ----------------------------------------------------------------------------
USE novapharma_dw;

DROP TABLE IF EXISTS quarantine_sales;
CREATE TABLE quarantine_sales (
    transaction_id    VARCHAR(12)    NOT NULL,
    invoice_date      DATE,
    month_key         INT,
    sku_code          VARCHAR(20),
    customer_code     VARCHAR(20),
    units             INT,
    gross_sales_eur   DECIMAL(12,2),
    discount_eur      DECIMAL(12,2),
    rebate_eur        DECIMAL(12,2),
    reason            VARCHAR(60)    NOT NULL,
    source_file       VARCHAR(80),
    loaded_at         DATETIME       NOT NULL DEFAULT CURRENT_TIMESTAMP,
    INDEX ix_q_reason (reason)
) ENGINE=InnoDB;

CREATE TABLE etl_log (
    run_id            INT AUTO_INCREMENT PRIMARY KEY,
    run_at            DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
    source_file       VARCHAR(80)  NOT NULL,
    rows_read         INT,
    rows_loaded       INT,
    rows_quarantined  INT,
    status            VARCHAR(10)  NOT NULL,      -- OK / FAIL
    message           VARCHAR(255)
) ENGINE=InnoDB;

SHOW TABLES;
-- ----------------------------------------------------------------------------

