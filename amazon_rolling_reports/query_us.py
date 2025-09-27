def sql_us():
    return r"""
    DECLARE @start_date date = '2019-01-01';

    DECLARE @last_date date;
    DECLARE @ty_year   int;
    DECLARE @ly_year   int;
    DECLARE @iso_wk    int;

    DECLARE @wk        date;
    DECLARE @cols      nvarchar(max);
    DECLARE @cols_proj nvarchar(max);
    DECLARE @sum       nvarchar(max);
    DECLARE @sql       nvarchar(max);

    SELECT @last_date = MAX([Week]) FROM [CBQ2].[cb].[Sellthrough_Amazon];
    SET @ty_year = YEAR(@last_date);
    SET @ly_year = @ty_year - 1;
    SET @iso_wk  = DATEPART(ISO_WEEK, @last_date);

    -- Helper: derive Monday=1 weekday regardless of @@DATEFIRST
    DECLARE @dow_last int = (DATEPART(WEEKDAY, @last_date) + @@DATEFIRST - 2) % 7 + 1; -- 1..7 with Monday=1
    DECLARE @iso_monday_last date = DATEADD(day, 1-@dow_last, @last_date);

    -- Monday of ISO week 1 (the week containing Jan 4) for TY and LY
    DECLARE @jan4_ty date = DATEFROMPARTS(@ty_year, 1, 4);
    DECLARE @jan4_ly date = DATEFROMPARTS(@ly_year, 1, 4);

    DECLARE @dow_jan4_ty int = (DATEPART(WEEKDAY, @jan4_ty) + @@DATEFIRST - 2) % 7 + 1;
    DECLARE @dow_jan4_ly int = (DATEPART(WEEKDAY, @jan4_ly) + @@DATEFIRST - 2) % 7 + 1;

    DECLARE @iso_start_ty date = DATEADD(day, 1-@dow_jan4_ty, @jan4_ty);
    DECLARE @iso_start_ly date = DATEADD(day, 1-@dow_jan4_ly, @jan4_ly);

    -- LY end aligned to same ISO week number as @last_date (end = Sunday of that ISO week)
    DECLARE @iso_end_ly date = DATEADD(day, 6, DATEADD(week, @iso_wk-1, @iso_start_ly)); -- ✅ use DATEADD(day,6,...)

    IF OBJECT_ID('tempdb..#items') IS NOT NULL DROP TABLE #items;
    SELECT
        i.*,
        UPPER(RIGHT(REPLICATE('0',13) + REPLACE(REPLACE(CONVERT(varchar(32), i.ITEM_TITLE),'-',''),' ',''), 13)) AS ISBN13
    INTO #items
    FROM ebs.item i
    WHERE i.PRODUCT_TYPE IN ('BK','FT','CP','RP','DI','')
      AND i.PUBLISHER_CODE NOT IN (
            'Benefit','AFO LLC','Glam Media','PQ Blackwell','PRINCETON','AMMO Books',
            'San Francisco Art Institute','FareArts','Sager','In Active','Driscolls',
            'Impossible Foods','Moleskine'
      );

    IF OBJECT_ID('tempdb..#agg') IS NOT NULL DROP TABLE #agg;

    WITH src AS (
        SELECT
            UPPER(RIGHT(REPLICATE('0',13) + REPLACE(REPLACE(CONVERT(varchar(32), sta.ISBN),'-',''),' ',''), 13)) AS ISBN13,
            CAST(sta.[Week] AS date) AS [Week],
            ISNULL(sta.UnitShipped, 0) AS UnitShipped,
            ISNULL(sta.OnHand, 0)         AS OnHand
        FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
        WHERE sta.[Week] >= @start_date
    )
    SELECT
        s.ISBN13,
        SUM(CASE WHEN s.[Week] = @last_date THEN s.OnHand ELSE 0 END) AS OH,
        -- rolling 52 weeks (inclusive)
        SUM(CASE WHEN s.[Week] BETWEEN DATEADD(week,-51,@last_date) AND @last_date THEN s.UnitShipped ELSE 0 END) AS W52,
        -- last 6 weeks (inclusive)
        SUM(CASE WHEN s.[Week] BETWEEN DATEADD(week,-5,@last_date)  AND @last_date THEN s.UnitShipped ELSE 0 END) AS SumLast6W,
        -- ISO TYTD: ISO start of TY through @last_date
        SUM(CASE WHEN s.[Week] >= @iso_start_ty AND s.[Week] <= @last_date
                 THEN s.UnitShipped ELSE 0 END) AS TYTD,
        -- ISO LYTD: ISO start of LY through end of same ISO week number as @last_date
        SUM(CASE WHEN s.[Week] >= @iso_start_ly AND s.[Week] <= @iso_end_ly
                 THEN s.UnitShipped ELSE 0 END) AS LYTD,
        -- Last calendar year total (keep if needed)
        SUM(CASE WHEN s.[Week] >= DATEFROMPARTS(@ly_year,1,1)
                      AND s.[Week] <  DATEFROMPARTS(@ty_year,1,1)
                 THEN s.UnitShipped ELSE 0 END) AS LY_FY
    INTO #agg
    FROM src s
    GROUP BY s.ISBN13;

    -- Build week list
    DECLARE @weeks TABLE (wk date PRIMARY KEY);
    SET @wk = @last_date;
    WHILE (@wk >= @start_date)
    BEGIN
        INSERT INTO @weeks(wk) VALUES(@wk);
        SET @wk = DATEADD(week, -1, @wk);
    END;

    -- Dynamic columns
    SELECT @cols =
    STUFF((
        SELECT ',' + QUOTENAME(CONVERT(varchar(10), wk, 110))
        FROM @weeks
        ORDER BY wk DESC
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 1, '');

    SELECT @cols_proj =
    STUFF((
        SELECT ',ISNULL(' + QUOTENAME(CONVERT(varchar(10), wk, 110)) + ',0) AS ' + QUOTENAME(CONVERT(varchar(10), wk, 110))
        FROM @weeks
        ORDER BY wk DESC
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 1, '');

    SELECT @sum =
    STUFF((
        SELECT ' + ISNULL(' + QUOTENAME(CONVERT(varchar(10), wk, 110)) + ',0)'
        FROM @weeks
        ORDER BY wk DESC
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 3, '');

    -- Pivot
    SET @sql = N'
    ;WITH base AS (
        SELECT
            UPPER(RIGHT(REPLICATE(''0'',13) + REPLACE(REPLACE(CONVERT(varchar(32), sta.ISBN),''-'',''''),'' '',''''), 13)) AS ISBN13,
            CONVERT(varchar(10), sta.[Week], 110) AS wk_label,
            ISNULL(sta.UnitShipped, 0) AS UnitShipped
        FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
        WHERE sta.[Week] >= @start_date
    ),
    p AS (
        SELECT *
        FROM base
        PIVOT (SUM(UnitShipped) FOR wk_label IN (' + @cols + N')) pv
    )
    SELECT
        CASE WHEN it.PUBLISHER_CODE = ''Quadrille Publishing Limited'' THEN ''Quadrille'' ELSE it.PUBLISHER_CODE END AS Pub,
        it.PRODUCT_TYPE AS pt,
        it.FORMAT       AS ft,
        CASE WHEN LEFT(it.PUBLISHING_GROUP,3) = ''BAR'' THEN ''BAR'' ELSE it.PUBLISHING_GROUP END AS pgrp,
        it.ITEM_TITLE   AS [ISBN],
        it.SHORT_TITLE  AS Title,
        it.PRICE_AMOUNT AS Price,
        CONVERT(varchar(10), it.AMORTIZATION_DATE, 110) AS PubDate,
        ISNULL(ag.OH,0)                                     AS OH,
        ISNULL(ag.W52,0)                                    AS W52,
        CAST(ISNULL(ag.SumLast6W,0) / 6.0 AS decimal(18,2)) AS AvgLast6W,
        ISNULL(ag.TYTD,0)                                   AS TYTD,
        ISNULL(ag.LYTD,0)                                   AS LYTD,
        ISNULL(ag.LY_FY,0)                                  AS LY_FY,
        (' + @sum + N') AS [LTD],
        ' + @cols_proj + N'
    FROM p
    JOIN #items it  ON p.ISBN13 = it.ISBN13
    LEFT JOIN #agg ag ON p.ISBN13 = ag.ISBN13
    ORDER BY Pub, pt, ft, Title;';

    EXEC sp_executesql @sql, N'@start_date date', @start_date = @start_date;
    """
