-- Post-deploy validation. Run as ACCOUNTADMIN, then as COCO_SHEETS_ANALYST
-- (after granting that role to your user). Expected results are in the comments.

-- 1. Security integration exists and is enabled, with PKCE enforced.
SHOW SECURITY INTEGRATIONS LIKE 'COCO_EXCEL_OAUTH';
-- Expect: enabled = true, OAuth client = CUSTOM, client type = PUBLIC.

-- 2. Masking policy is attached to the email column.
SELECT POLICY_NAME, REF_COLUMN_NAME
FROM TABLE(COCO_SHEETS_DEMO.INFORMATION_SCHEMA.POLICY_REFERENCES(
  REF_ENTITY_NAME => 'COCO_SHEETS_DEMO.SALES.ORDERS', REF_ENTITY_DOMAIN => 'TABLE'));
-- Expect: MASK_EMAIL on CUSTOMER_EMAIL, and RAP_REGION as a table-level row access policy (no column).

-- 3. Analyst row visibility (run as COCO_SHEETS_ANALYST). Expect only NA and EMEA.
-- Observed after the NULL-region fix: NA 1,595 rows, EMEA 1,761 rows. Counts change with data.
SELECT REGION, COUNT(*) AS ROWS_VISIBLE
FROM COCO_SHEETS_DEMO.SALES.ORDERS
GROUP BY REGION
ORDER BY REGION;

-- 4. Analyst email masking (run as COCO_SHEETS_ANALYST). Expect ***MASKED***.
SELECT MAX(CUSTOMER_EMAIL) AS SAMPLE_EMAIL FROM COCO_SHEETS_DEMO.SALES.ORDERS;

-- 5. Semantic view metrics return values by region.
SELECT * FROM SEMANTIC_VIEW(
  COCO_SHEETS_DEMO.SALES.SV_SALES_METRICS
  DIMENSIONS orders.region
  METRICS orders.total_revenue, orders.gross_margin_pct, orders.churn_rate
)
WHERE REGION IS NOT NULL
ORDER BY REGION;
-- Expect: one row per region. A NULL-region total row also appears when REGION is not filtered.

-- 6. Cortex call works (SQL equivalent of CORTEX.ASK; CORTEX.ASK is an Excel cell function).
-- Expect: OK. Run as COCO_SHEETS_ANALYST and as ACCOUNTADMIN.
SELECT AI_COMPLETE('llama3.1-8b', 'Reply with OK.') AS CORTEX_CHECK;

-- 7. Data quality: no NULL dimensions after the backfill. Expect 0 for each.
SELECT COUNT_IF(REGION IS NULL) AS NULL_REGION, COUNT_IF(PRODUCT_LINE IS NULL) AS NULL_PRODUCT
FROM COCO_SHEETS_DEMO.SALES.ORDERS;
