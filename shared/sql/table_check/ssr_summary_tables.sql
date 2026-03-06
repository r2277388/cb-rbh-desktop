SELECT
    tlu.TableName,
    tlu.LastUpdated
FROM metrics.TableLastUpdated tlu
WHERE tlu.TableName IN ('ssr.SalesSSRRow', 'ebs.Sales', 'ebs.Item')
ORDER BY tlu.LastUpdated DESC;

