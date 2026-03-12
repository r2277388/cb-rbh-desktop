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
    ho.SSRRowID AS ssr_id,
    ho.ISBN AS isbn,
    CASE WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR' ELSE i.PUBLISHING_GROUP END AS pgrp,
    ho.EnteredDate AS entered_date,
    ho.ReleaseDate AS release_date,
    ho.OrderCancelDate AS order_cancel_date,
    ho.OrderTypeCode AS order_type_code,
    SUM(ho.Quantity) AS qty
FROM hachette.HachetteOrders ho
INNER JOIN ebs.item i ON i.ISBN = ho.ISBN
INNER JOIN ssr.SSRRow ssr_row ON ho.SSRRowID = ssr_row.SSRRowID
INNER JOIN top_isbns ti ON ti.ISBN = i.ISBN
WHERE i.PUBLISHER_CODE = 'Chronicle'
  AND i.PUBLISHING_GROUP NOT IN ('MKT')
  AND ho.EnteredDate > (GETDATE() - {orders_lookback_days})
  AND i.PRICE_AMOUNT > 0
  AND ho.OrderTypeCode NOT IN ('DELETED','CREDIT HOLD')
GROUP BY
    ho.SSRRowID,
    ho.ISBN,
    CASE WHEN LEFT(i.PUBLISHING_GROUP, 3) = 'BAR' THEN 'BAR' ELSE i.PUBLISHING_GROUP END,
    ho.EnteredDate,
    ho.ReleaseDate,
    ho.OrderCancelDate,
    ho.OrderTypeCode
ORDER BY ho.EnteredDate ASC;
