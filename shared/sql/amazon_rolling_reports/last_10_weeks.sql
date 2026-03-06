SELECT TOP 10
    sta.[WEEK],
    COUNT(*) AS [RowCount],
    SUM(sta.[CustomerOrders]) AS [CustomerOrders],
    SUM(sta.[UnitShipped]) AS [UnitShipped],
    SUM(sta.[OnHand]) AS [OnHand]
FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
GROUP BY sta.[WEEK]
ORDER BY sta.[WEEK] DESC;

