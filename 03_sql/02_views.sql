-- ============================================================================
-- NovaPharma Iberia — 02_views.sql
-- Views for Power BI. Run after 01_schema.sql and after loading the data.
-- ============================================================================
USE novapharma_dw;

-- v_sales_enriched: the Sales_Enriched sheet done right.
-- 1 row = 1 invoice line (same grain as fact_sales), enriched with the masters.
CREATE OR REPLACE VIEW v_sales_enriched AS
SELECT
    s.transaction_id,
    s.invoice_date,
    s.month_key,
    c.year,
    c.quarter,
    s.sku_code,
    p.sku_name,
    p.brand_name,
    b.therapeutic_area,
    s.customer_code,
    cu.customer_name,
    cu.channel,
    cu.territory_code,
    t.territory_name,
    t.region,
    t.country,
    t.sales_rep,
    s.units,
    s.gross_sales_eur,
    s.discount_eur,
    s.rebate_eur,
    s.gross_sales_eur - s.discount_eur - s.rebate_eur              AS net_sales_eur,
    s.units * p.cogs_per_unit_eur                                    AS cogs_eur,
    s.gross_sales_eur - s.discount_eur - s.rebate_eur
        - s.units * p.cogs_per_unit_eur                              AS gross_margin_eur
FROM fact_sales    s
JOIN dim_product   p  ON p.sku_code       = s.sku_code
JOIN dim_brand     b  ON b.brand_name     = p.brand_name
JOIN dim_customer  cu ON cu.customer_code = s.customer_code
JOIN dim_territory t  ON t.territory_code = cu.territory_code
JOIN dim_calendar  c  ON c.month_key      = s.month_key;

-- Check: same row count and same net sales as the reconciliation in Step 3
-- (+ the monthly file loaded in Step 3.B: 48,100 + 1,572 = 49,672 rows)
SELECT COUNT(*) AS rows_, ROUND(SUM(net_sales_eur), 2) AS net_sales
FROM v_sales_enriched;

SELECT month_key, COUNT(*) AS rows_, ROUND(SUM(net_sales_eur), 2) AS net_sales
FROM v_sales_enriched
WHERE month_key >= 202608
GROUP BY month_key;

-- v_bva: Actual vs Budget vs Forecast, long format.
-- 1 row = 1 version x month x SKU x territory. 'Actual' is treated as one more version.
CREATE OR REPLACE VIEW v_bva AS
SELECT
    'Actual'                                                  AS version_name,
    s.month_key,
    s.sku_code,
    cu.territory_code,
    SUM(s.units)                                              AS units,
    SUM(s.gross_sales_eur - s.discount_eur - s.rebate_eur)    AS net_sales_eur
FROM fact_sales   s
JOIN dim_customer cu ON cu.customer_code = s.customer_code
GROUP BY s.month_key, s.sku_code, cu.territory_code
UNION ALL
SELECT
    p.version_name,
    p.month_key,
    p.sku_code,
    p.territory_code,
    p.units,
    p.net_sales_eur
FROM fact_plan p;

-- Check: one row per version, totals per version
SELECT version_name, COUNT(*) AS rows_, ROUND(SUM(net_sales_eur), 2) AS net_sales
FROM v_bva
GROUP BY version_name
ORDER BY version_name;
