def sql_us():
    return """
    -- Params
    DECLARE @start_date date = '2019-01-01';

    -- Vars
    DECLARE @last_date date;
    DECLARE @ty_year   int;
    DECLARE @ly_year   int;
    DECLARE @iso_wk    int;

    DECLARE @wk        date;
    DECLARE @cols      nvarchar(max);
    DECLARE @cols_proj nvarchar(max);  -- ISNULL([col],0) AS [col]
    DECLARE @sum       nvarchar(max);
    DECLARE @sql       nvarchar(max);

    ------------------------------------------------------------
    -- 1) Last date & ISO-week context
    ------------------------------------------------------------
    SELECT @last_date = MAX([Week])
    FROM [CBQ2].[cb].[Sellthrough_Amazon];

    SET @ty_year = YEAR(@last_date);
    SET @ly_year = @ty_year - 1;
    SET @iso_wk  = DATEPART(ISO_WEEK, @last_date);

    ------------------------------------------------------------
    -- 2) Items snapshot with padded join key
    ------------------------------------------------------------
    IF OBJECT_ID('tempdb..#items') IS NOT NULL DROP TABLE #items;

    SELECT
    i.*,
    RIGHT(REPLICATE('0',13) + i.ITEM_TITLE, 13) AS ISBN13
    INTO #items
    FROM ebs.item i
    WHERE i.PRODUCT_TYPE IN ('BK','FT','CP','RP','DI','')
    AND i.PUBLISHER_CODE NOT IN (
        'Benefit','AFO LLC','Glam Media','PQ Blackwell','PRINCETON','AMMO Books',
        'San Francisco Art Institute','FareArts','Sager','In Active','Driscolls',
        'Impossible Foods','Moleskine'
    );

    ------------------------------------------------------------
    -- 3) Aggregations (restored)
    ------------------------------------------------------------
    IF OBJECT_ID('tempdb..#agg') IS NOT NULL DROP TABLE #agg;

    WITH src AS (
    SELECT
        UPPER(RIGHT(REPLICATE('0',13) + CONVERT(varchar(32), sta.ISBN), 13)) AS ISBN13,
        sta.[Week],
        ISNULL(sta.UnitShipped, 0) AS UnitShipped   -- <— add ISNULL here
    FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
    WHERE YEAR(sta.[Week]) >= 2019
    )

    SELECT
    s.ISBN13,
    SUM(CASE WHEN s.[Week] = @last_date THEN s.UnitShipped ELSE 0 END) AS OH,
    SUM(CASE WHEN s.[Week] BETWEEN DATEADD(year,-1,@last_date) AND @last_date THEN s.UnitShipped ELSE 0 END) AS W52,
    SUM(CASE WHEN s.[Week] BETWEEN DATEADD(week,-6,@last_date) AND @last_date THEN s.UnitShipped ELSE 0 END) AS SumLast6W,
    SUM(CASE WHEN YEAR(s.[Week]) = @ty_year AND DATEPART(ISO_WEEK, s.[Week]) <= @iso_wk THEN s.UnitShipped ELSE 0 END) AS TYTD,
    SUM(CASE WHEN YEAR(s.[Week]) = @ly_year AND DATEPART(ISO_WEEK, s.[Week]) <= @iso_wk THEN s.UnitShipped ELSE 0 END) AS LYTD,
    SUM(CASE WHEN s.[Week] >= DATEFROMPARTS(@ly_year,1,1) AND s.[Week] < DATEFROMPARTS(@ty_year,1,1) THEN s.UnitShipped ELSE 0 END) AS LY_FY
    INTO #agg
    FROM src s
    GROUP BY s.ISBN13;

    ------------------------------------------------------------
    -- 4) Build week list & dynamic parts (no STRING_AGG)
    ------------------------------------------------------------
    DECLARE @weeks TABLE (wk date PRIMARY KEY);
    SET @wk = @last_date;
    WHILE (@wk >= @start_date)
    BEGIN
    INSERT INTO @weeks(wk) VALUES(@wk);
    SET @wk = DATEADD(week, -1, @wk);
    END

    -- Column list: [mm-dd-yyyy], [mm-dd-yyyy], ... newest → oldest
    SELECT @cols =
    STUFF((
        SELECT ',' + QUOTENAME(CONVERT(varchar(10), wk, 110))
        FROM @weeks
        ORDER BY wk DESC
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 1, '');

    -- Projection with ISNULL for each pivoted column
    SELECT @cols_proj =
    STUFF((
        SELECT ',ISNULL(' + QUOTENAME(CONVERT(varchar(10), wk, 110)) + ',0) AS ' + QUOTENAME(CONVERT(varchar(10), wk, 110))
        FROM @weeks
        ORDER BY wk DESC
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 1, '');

    -- Total expression across weekly cols
    SELECT @sum =
    STUFF((
        SELECT ' + ISNULL(' + QUOTENAME(CONVERT(varchar(10), wk, 110)) + ',0)'
        FROM @weeks
        ORDER BY wk DESC
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 3, '');  -- trims leading ' + '

    ------------------------------------------------------------
    -- 5) Dynamic pivot with agg columns FIRST, then weeks, then Total
    ------------------------------------------------------------
    SET @sql = N'
    ;WITH base AS (
    SELECT
        UPPER(RIGHT(REPLICATE(''0'',13) + CONVERT(varchar(32), sta.ISBN), 13)) AS ISBN13,
        CONVERT(varchar(10), sta.[Week], 110) AS wk_label,  -- mm-dd-yyyy
        ISNULL(sta.UnitShipped, 0) AS UnitShipped     -- ← add ISNULL here
    FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
    WHERE sta.[Week] >= @start_date
    ),
    p AS (
    SELECT *
    FROM base
    PIVOT (SUM(UnitShipped) FOR wk_label IN (' + @cols + N')) pv
    )
    SELECT
    -- item metadata
    CASE WHEN it.PUBLISHER_CODE = ''Quadrille Publishing Limited'' THEN ''Quadrille'' ELSE it.PUBLISHER_CODE END AS Pub,
    it.PRODUCT_TYPE AS pt,
    it.FORMAT       AS ft,
    CASE WHEN LEFT(it.PUBLISHING_GROUP,3) = ''BAR'' THEN ''BAR'' ELSE it.PUBLISHING_GROUP END AS pgrp,
    it.ITEM_TITLE   AS [ISBN],
    it.SHORT_TITLE  AS Title,
    it.PRICE_AMOUNT AS Price,
    CONVERT(varchar(10), it.AMORTIZATION_DATE, 110) AS PubDate,

    -- AGG COLUMNS FIRST
    ISNULL(ag.OH,0)                                     AS OH,
    ISNULL(ag.W52,0)                                    AS W52,
    CAST(ISNULL(ag.SumLast6W,0) / 6.0 AS decimal(18,2)) AS AvgLast6W,
    ISNULL(ag.TYTD,0)                                   AS TYTD,
    ISNULL(ag.LYTD,0)                                   AS LYTD,
    ISNULL(ag.LY_FY,0)                                  AS LY_FY,

    -- >>> TOTAL BEFORE WEEKLY COLUMNS <<<
    (' + @sum + N') AS [LTD],

    -- WEEKLY COLUMNS (0 instead of NULL), newest → oldest
    ' + @cols_proj + N'

    FROM p
    JOIN #items it
    ON p.ISBN13 = it.ISBN13
    LEFT JOIN #agg ag
    ON p.ISBN13 = ag.ISBN13
    ORDER BY Pub, pt, ft, Title;
    ';

    -- run it
    EXEC sp_executesql @sql, N'@start_date date', @start_date = @start_date;

    """