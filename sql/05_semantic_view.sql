-- Semantic view: one definition of revenue, margin, and churn for every workbook and sheet.

CREATE OR REPLACE SEMANTIC VIEW COCO_SHEETS_DEMO.SALES.SV_SALES_METRICS
  TABLES (
    orders AS COCO_SHEETS_DEMO.SALES.ORDERS PRIMARY KEY (ORDER_ID)
  )
  FACTS (
    orders.revenue AS REVENUE,
    orders.cost AS COST
  )
  DIMENSIONS (
    orders.order_date AS ORDER_DATE,
    orders.region AS REGION,
    orders.product_line AS PRODUCT_LINE,
    orders.customer_segment AS CUSTOMER_SEGMENT
  )
  METRICS (
    orders.total_revenue AS SUM(orders.revenue),
    orders.total_cost AS SUM(orders.cost),
    orders.gross_margin_pct AS (SUM(orders.revenue) - SUM(orders.cost)) / NULLIF(SUM(orders.revenue), 0),
    orders.churn_rate AS AVG(CASE WHEN orders.churned THEN 1 ELSE 0 END)
  );

GRANT SELECT ON SEMANTIC VIEW COCO_SHEETS_DEMO.SALES.SV_SALES_METRICS TO ROLE COCO_SHEETS_ANALYST;
