SELECT DISTINCT TOP 10
    sta.[WEEK]
    ,count(sta.week) Row_Cnt
    ,sum(sta.OnHand) AS [OnHand]
    ,sum(sta.UnitShipped) AS [UnitShipped]
    ,sum(sta.CustomerOrders) AS [CustomerOrders]
FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
GROUP BY
    sta.[WEEK]
ORDER BY [WEEK] DESC;
