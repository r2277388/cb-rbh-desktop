def sql_co():
    'customer orders with weekly pivot and fixed-window aggregates'
    return """
     /* ===== A) Date anchors (SSMS/2016 friendly) ===== */
    DECLARE @last_date         date;
    SELECT  @last_date = MAX([Week]) FROM [CBQ2].[cb].[Sellthrough_Amazon];

    DECLARE @start_pivot       date = DATEFROMPARTS(YEAR(@last_date)-2,1,1);  -- ~3-year pivot start
    DECLARE @d_6w_start        date = DATEADD(week,-6, @last_date);
    DECLARE @d_52w_start       date = DATEADD(year,-1, @last_date);
    DECLARE @d_ty_start        date = DATEFROMPARTS(YEAR(@last_date),1,1);
    DECLARE @d_ly_start        date = DATEFROMPARTS(YEAR(@last_date)-1,1,1);
    DECLARE @d_lytd_end_next   date = DATEADD(day,1, DATEFROMPARTS(YEAR(@last_date)-1,MONTH(@last_date),DAY(@last_date)));
    DECLARE @d_ty_start_next   date = DATEFROMPARTS(YEAR(@last_date),1,1);

    /* ===== B) Optional: publishers to exclude ===== */
    IF OBJECT_ID('tempdb..#excl') IS NOT NULL DROP TABLE #excl;
    CREATE TABLE #excl (publisher_code varchar(100) PRIMARY KEY);
    INSERT INTO #excl(publisher_code)
    VALUES
    ('Benefit'),('AFO LLC'),('Glam Media'),('PQ Blackwell'),('PRINCETON'),
    ('AMMO Books'),('San Francisco Art Institute'),('FareArts'),('Sager'),
    ('In Active'),('Driscolls'),('Impossible Foods'),('Moleskine');
    -- If you don't need exclusions, comment out the INSERT above (leave empty table).

    /* ===== C) One-pass conditional aggregation for fixed-window metrics ===== */
    /* Use ISNULL() to avoid the NULL-eliminated warning */
    IF OBJECT_ID('tempdb..#agg') IS NOT NULL DROP TABLE #agg;
    SELECT
        sta.ISBN,
        OH      = SUM(CASE WHEN sta.[Week] = @last_date THEN ISNULL(sta.UnitShipped,0) ELSE 0 END),
        W52     = SUM(CASE WHEN sta.[Week] >= @d_52w_start AND sta.[Week] <= @last_date THEN ISNULL(sta.UnitShipped,0) ELSE 0 END),
        SumLast6W = SUM(CASE WHEN sta.[Week] >= @d_6w_start  AND sta.[Week] <= @last_date THEN ISNULL(sta.UnitShipped,0) ELSE 0 END),
        TYTD    = SUM(CASE WHEN sta.[Week] >= @d_ty_start  AND sta.[Week] <= @last_date THEN ISNULL(sta.UnitShipped,0) ELSE 0 END),
        LYTD    = SUM(CASE WHEN sta.[Week] >= @d_ly_start  AND sta.[Week] <  @d_lytd_end_next THEN ISNULL(sta.UnitShipped,0) ELSE 0 END),
        LY_FY   = SUM(CASE WHEN sta.[Week] >= @d_ly_start  AND sta.[Week] <  @d_ty_start_next THEN ISNULL(sta.UnitShipped,0) ELSE 0 END)
    INTO #agg
    FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
    GROUP BY sta.ISBN;

    /* ===== D) Materialize the week list once (newest → oldest) ===== */
    IF OBJECT_ID('tempdb..#wk') IS NOT NULL DROP TABLE #wk;
    SELECT DISTINCT [Week]
    INTO #wk
    FROM [CBQ2].[cb].[Sellthrough_Amazon]
    WHERE [Week] BETWEEN @start_pivot AND @last_date;

    CREATE UNIQUE CLUSTERED INDEX IX_wk_Week ON #wk([Week]); -- helps ordering

    /* Build weekly column lists (DESC) */
    DECLARE @colsIn  nvarchar(max);
    DECLARE @colsSel nvarchar(max);

    SELECT @colsIn = STUFF((
    SELECT ',' + QUOTENAME(CONVERT(varchar(10), [Week], 23))
    FROM #wk
    ORDER BY [Week] DESC
    FOR XML PATH(''), TYPE
    ).value('.','nvarchar(max)'),1,1,'');

    SELECT @colsSel = STUFF((
    SELECT ',ISNULL(' + QUOTENAME(CONVERT(varchar(10), [Week], 23)) + ',0) AS '
        + QUOTENAME(CONVERT(varchar(10), [Week], 23))
    FROM #wk
    ORDER BY [Week] DESC
    FOR XML PATH(''), TYPE
    ).value('.','nvarchar(max)'),1,1,'');

    /* ===== E) Dynamic pivot over weekly UnitShipped, then join metrics ===== */
    DECLARE @sql nvarchar(max) = N'
    ;WITH src AS (
    SELECT
        i.PUBLISHER_CODE AS Pub,
        i.PRODUCT_TYPE   AS pt,
        i.FORMAT         AS ft,
        CASE WHEN LEFT(i.PUBLISHING_GROUP,3) = ''BAR'' THEN ''BAR'' ELSE i.PUBLISHING_GROUP END AS pgrp,
        sta.ISBN,
        i.SHORT_TITLE    AS Title,
        i.PRICE_AMOUNT   AS Price,
        i.AMORTIZATION_DATE,
        CAST(sta.[Week] AS date) AS [Week],
        ISNULL(sta.UnitShipped,0) AS UnitShipped
    FROM [CBQ2].[cb].[Sellthrough_Amazon] sta
    JOIN ebs.item i ON sta.ISBN = i.ITEM_TITLE
    WHERE
        sta.[Week] >= @start_pivot AND sta.[Week] <= @last_date
        AND NOT EXISTS (SELECT 1 FROM #excl e WHERE e.publisher_code = i.PUBLISHER_CODE)
    ),
    pvt AS (
    SELECT
        Pub, pt, ft, pgrp, ISBN, Title, Price, AMORTIZATION_DATE,
        CONVERT(varchar(10), [Week], 23) AS wk,
        UnitShipped
    FROM src
    ),
    pivoted AS (
    SELECT
        Pub, pt, ft, pgrp, ISBN, Title, Price, AMORTIZATION_DATE, ' + @colsIn + N'
    FROM pvt
    PIVOT (SUM(UnitShipped) FOR wk IN (' + @colsIn + N')) pv
    )
    SELECT
        pv.Pub, pv.pt, pv.ft, pv.pgrp, pv.ISBN, pv.Title, pv.Price, pv.AMORTIZATION_DATE,
        a.OH,
        a.W52,
        CAST(1.0 * a.SumLast6W / 6 AS decimal(18,2)) AS AvgLast6W,
        a.TYTD,
        a.LYTD,
        a.LY_FY,
        ' + @colsSel + N'
    FROM pivoted pv
    LEFT JOIN #agg a ON a.ISBN = pv.ISBN
    ORDER BY pv.Pub, pv.pt, pv.ft, pv.pgrp, pv.ISBN;';

    EXEC sys.sp_executesql
    @sql,
    N'@start_pivot date, @last_date date',
    @start_pivot=@start_pivot, @last_date=@last_date;


    """