SELECT TOP 100
    tlu.TableName,
    tlu.LastUpdated
FROM metrics.TableLastUpdated tlu
ORDER BY
    tlu.LastUpdated desc

