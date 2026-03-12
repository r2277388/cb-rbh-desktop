WITH top_isbns AS (
    SELECT TOP ({top_n_isbns}) i.isbn
    FROM [CBQ2].[summary].[TitleMonthlySales] tms
    INNER JOIN ebs.item i ON i.ITEM_ID = tms.ITEM_ID
    WHERE tms.PERIOD BETWEEN '{sales_top_period_start}' AND '{sales_top_period_end}'
      AND [DISTRIBUTION_DIRECT] = 'N'
      AND tms.PRODUCT_TYPE IN ('BK','FT')
      AND tms.[FRONTBACKLISTCODE] = 'B'
      AND tms.SALETYPECODE = 'N'
      AND i.PUBLISHER_CODE = 'Chronicle'
    GROUP BY i.ISBN
    ORDER BY SUM([SalesNetValue]) DESC
)
SELECT
    tms.[PERIOD] AS period,
    i.ISBN AS isbn,
    CASE WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR' ELSE i.PUBLISHING_GROUP END AS pgrp,
    i.[PRODUCT_TYPE] AS pt,
    tms.[SSRROWID] AS ssr_id,
    SUM([SalesQty]) AS qty,
    SUM([SalesNetValue]) AS val
FROM [CBQ2].[summary].[TitleMonthlySales] tms
INNER JOIN ssr.SSRRow ssr_row ON ssr_row.SSRRowID = tms.SSRROWID
INNER JOIN ebs.item i ON i.ITEM_ID = tms.ITEM_ID
INNER JOIN top_isbns ti ON ti.ISBN = i.ISBN
WHERE tms.PERIOD >= '{sales_start_period}'
  AND [DISTRIBUTION_DIRECT] = 'N'
  AND tms.PRODUCT_TYPE IN ('BK','FT')
  AND tms.SALETYPECODE = 'N'
  AND i.PUBLISHER_CODE = 'Chronicle'
GROUP BY
    tms.[PERIOD],
    i.ISBN,
    CASE WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR' ELSE i.PUBLISHING_GROUP END,
    i.[PRODUCT_TYPE],
    tms.[SSRROWID];
