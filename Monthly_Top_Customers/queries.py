from __future__ import annotations


def sql_list(values: list[str]) -> str:
    return ", ".join(f"'{value}'" for value in values)


def title_sales_sql(period: str) -> str:
    return f"""
DECLARE @ty_tp  varchar(6) = '{period}';
DECLARE @ty_p01 varchar(6) = LEFT(@ty_tp, 4) + '01';
DECLARE @ty_p12 varchar(6) = LEFT(@ty_tp, 4) + '12';

DECLARE @ly_tp  varchar(6) = dbo.fnCalcPeriod(@ty_tp, -12);
DECLARE @ly_p01 varchar(6) = dbo.fnCalcPeriod(@ty_p01, -12);
DECLARE @ly_p12 varchar(6) = dbo.fnCalcPeriod(@ty_p12, -12);

DECLARE @lly_p01 varchar(6) = dbo.fnCalcPeriod(@ly_p01, -12);
DECLARE @lly_p12 varchar(6) = dbo.fnCalcPeriod(@ly_p12, -12);

WITH BucketedSales AS (
    SELECT
        b.Bucket,

        ssr.Description AS [type],
        t.PUBLISHER_CODE AS pub,
        t.PRODUCT_TYPE AS pt,
        t.REPORTING_CATEGORY AS cat,
        t.PUBLISHING_GROUP AS pgr,
        t.ITEM_TITLE AS isbn,
        t.SHORT_TITLE AS title,
        t.PRICE_AMOUNT AS price,
        t.AMORTIZATION_DATE AS [pub date],
        t.SEASON AS sea,

        CASE WHEN sd.INVOICE_LINE_TYPE = 'sale'
            THEN sd.REVENUE_AMOUNT ELSE 0 END AS SaleDollars,

        CASE WHEN sd.INVOICE_LINE_TYPE = 'sale'
            THEN sd.UNIT_STANDARD_PRICE * sd.QUANTITY_INVOICED ELSE 0 END AS SaleRetail,

        CASE WHEN sd.INVOICE_LINE_TYPE = 'sale'
            THEN sd.QUANTITY_INVOICED ELSE 0 END AS SaleUnits,

        CASE WHEN sd.INVOICE_LINE_TYPE = 'return'
            THEN sd.REVENUE_AMOUNT ELSE 0 END AS ReturnDollars,

        CASE WHEN sd.INVOICE_LINE_TYPE = 'return'
            THEN sd.QUANTITY_INVOICED ELSE 0 END AS ReturnUnits

    FROM ebs.Sales AS sd
        INNER JOIN ssr.SalesSSRRow AS stie
            ON stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
        INNER JOIN ssr.SSRRow AS ssr
            ON ssr.SSRRowID = stie.SSRRowID
        INNER JOIN ebs.Item AS t
            ON t.ITEM_ID = sd.ITEM_ID

        CROSS APPLY (
            VALUES
                ('TY_Month', @ty_tp,   @ty_tp),
                ('LY_Month', @ly_tp,   @ly_tp),
                ('TY_YTD',   @ty_p01,  @ty_tp),
                ('LY_YTD',   @ly_p01,  @ly_tp),
                ('LY_FY',    @ly_p01,  @ly_p12),
                ('LLY_FY',   @lly_p01, @lly_p12)
        ) AS b(Bucket, StartPeriod, EndPeriod)

    WHERE
        sd.PERIOD BETWEEN @lly_p01 AND @ty_tp
        AND sd.PERIOD BETWEEN b.StartPeriod AND b.EndPeriod
        AND dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N'
        AND ssr.SSRRowID IN ('20', '06', '72', '120', '146', '32')
),

Agg AS (
    SELECT
        [type],
        pub,
        pt,
        cat,
        pgr,
        isbn,
        title,
        price,
        [pub date],
        sea,
        Bucket,

        SUM(SaleDollars) AS SaleDollars,
        SUM(SaleRetail) AS SaleRetail,
        SUM(SaleUnits) AS SaleUnits,
        SUM(ReturnDollars) AS ReturnDollars,
        SUM(ReturnUnits) AS ReturnUnits

    FROM BucketedSales
    GROUP BY
        [type],
        pub,
        pt,
        cat,
        pgr,
        isbn,
        title,
        price,
        [pub date],
        sea,
        Bucket
)

SELECT
    [type],
    pub,
    pt,
    cat,
    pgr,
    isbn,
    title,
    price,
    [pub date],
    sea,

    SUM(CASE WHEN Bucket = 'TY_Month' THEN SaleDollars ELSE 0 END) AS TY_Month_Sale_Dollars,
    SUM(CASE WHEN Bucket = 'TY_Month' THEN SaleRetail  ELSE 0 END) AS TY_Month_Sale_Retail,
    SUM(CASE WHEN Bucket = 'TY_Month' THEN SaleUnits   ELSE 0 END) AS TY_Month_Sale_Units,
    1 - SUM(CASE WHEN Bucket = 'TY_Month' THEN SaleDollars ELSE 0 END)
        / NULLIF(SUM(CASE WHEN Bucket = 'TY_Month' THEN SaleRetail ELSE 0 END), 0) AS TY_Month_Disc,
    SUM(CASE WHEN Bucket = 'TY_Month' THEN ReturnDollars ELSE 0 END) AS TY_Month_Return_Dollars,
    SUM(CASE WHEN Bucket = 'TY_Month' THEN ReturnUnits   ELSE 0 END) AS TY_Month_Return_Units,

    SUM(CASE WHEN Bucket = 'LY_Month' THEN SaleDollars ELSE 0 END) AS LY_Month_Sale_Dollars,
    SUM(CASE WHEN Bucket = 'LY_Month' THEN SaleRetail  ELSE 0 END) AS LY_Month_Sale_Retail,
    SUM(CASE WHEN Bucket = 'LY_Month' THEN SaleUnits   ELSE 0 END) AS LY_Month_Sale_Units,
    1 - SUM(CASE WHEN Bucket = 'LY_Month' THEN SaleDollars ELSE 0 END)
        / NULLIF(SUM(CASE WHEN Bucket = 'LY_Month' THEN SaleRetail ELSE 0 END), 0) AS LY_Month_Disc,
    SUM(CASE WHEN Bucket = 'LY_Month' THEN ReturnDollars ELSE 0 END) AS LY_Month_Return_Dollars,
    SUM(CASE WHEN Bucket = 'LY_Month' THEN ReturnUnits   ELSE 0 END) AS LY_Month_Return_Units,

    SUM(CASE WHEN Bucket = 'TY_YTD' THEN SaleDollars ELSE 0 END) AS TY_YTD_Sale_Dollars,
    SUM(CASE WHEN Bucket = 'TY_YTD' THEN SaleRetail  ELSE 0 END) AS TY_YTD_Sale_Retail,
    SUM(CASE WHEN Bucket = 'TY_YTD' THEN SaleUnits   ELSE 0 END) AS TY_YTD_Sale_Units,
    1 - SUM(CASE WHEN Bucket = 'TY_YTD' THEN SaleDollars ELSE 0 END)
        / NULLIF(SUM(CASE WHEN Bucket = 'TY_YTD' THEN SaleRetail ELSE 0 END), 0) AS TY_YTD_Disc,
    SUM(CASE WHEN Bucket = 'TY_YTD' THEN ReturnDollars ELSE 0 END) AS TY_YTD_Return_Dollars,
    SUM(CASE WHEN Bucket = 'TY_YTD' THEN ReturnUnits   ELSE 0 END) AS TY_YTD_Return_Units,

    SUM(CASE WHEN Bucket = 'LY_YTD' THEN SaleDollars ELSE 0 END) AS LY_YTD_Sale_Dollars,
    SUM(CASE WHEN Bucket = 'LY_YTD' THEN SaleRetail  ELSE 0 END) AS LY_YTD_Sale_Retail,
    SUM(CASE WHEN Bucket = 'LY_YTD' THEN SaleUnits   ELSE 0 END) AS LY_YTD_Sale_Units,
    1 - SUM(CASE WHEN Bucket = 'LY_YTD' THEN SaleDollars ELSE 0 END)
        / NULLIF(SUM(CASE WHEN Bucket = 'LY_YTD' THEN SaleRetail ELSE 0 END), 0) AS LY_YTD_Disc,
    SUM(CASE WHEN Bucket = 'LY_YTD' THEN ReturnDollars ELSE 0 END) AS LY_YTD_Return_Dollars,
    SUM(CASE WHEN Bucket = 'LY_YTD' THEN ReturnUnits   ELSE 0 END) AS LY_YTD_Return_Units,

    SUM(CASE WHEN Bucket = 'LY_FY' THEN SaleDollars ELSE 0 END) AS LY_FY_Sale_Dollars,
    SUM(CASE WHEN Bucket = 'LY_FY' THEN SaleRetail  ELSE 0 END) AS LY_FY_Sale_Retail,
    SUM(CASE WHEN Bucket = 'LY_FY' THEN SaleUnits   ELSE 0 END) AS LY_FY_Sale_Units,
    1 - SUM(CASE WHEN Bucket = 'LY_FY' THEN SaleDollars ELSE 0 END)
        / NULLIF(SUM(CASE WHEN Bucket = 'LY_FY' THEN SaleRetail ELSE 0 END), 0) AS LY_FY_Disc,
    SUM(CASE WHEN Bucket = 'LY_FY' THEN ReturnDollars ELSE 0 END) AS LY_FY_Return_Dollars,
    SUM(CASE WHEN Bucket = 'LY_FY' THEN ReturnUnits   ELSE 0 END) AS LY_FY_Return_Units,

    SUM(CASE WHEN Bucket = 'LLY_FY' THEN SaleDollars ELSE 0 END) AS LLY_FY_Sale_Dollars,
    SUM(CASE WHEN Bucket = 'LLY_FY' THEN SaleRetail  ELSE 0 END) AS LLY_FY_Sale_Retail,
    SUM(CASE WHEN Bucket = 'LLY_FY' THEN SaleUnits   ELSE 0 END) AS LLY_FY_Sale_Units,
    1 - SUM(CASE WHEN Bucket = 'LLY_FY' THEN SaleDollars ELSE 0 END)
        / NULLIF(SUM(CASE WHEN Bucket = 'LLY_FY' THEN SaleRetail ELSE 0 END), 0) AS LLY_FY_Disc,
    SUM(CASE WHEN Bucket = 'LLY_FY' THEN ReturnDollars ELSE 0 END) AS LLY_FY_Return_Dollars,
    SUM(CASE WHEN Bucket = 'LLY_FY' THEN ReturnUnits   ELSE 0 END) AS LLY_FY_Return_Units

FROM Agg
GROUP BY
    [type],
    pub,
    pt,
    cat,
    pgr,
    isbn,
    title,
    price,
    [pub date],
    sea

ORDER BY
    [type] ASC,
    TY_Month_Sale_Dollars DESC;
"""


def rep_code_lookup_sql(rep_codes: list[str]) -> str:
    return f"""
SELECT
    sr.SALESREP_NUMBER AS rep_number,
    sr.NAME AS rep_name
FROM ebs.SalesRep sr
WHERE sr.SALESREP_NUMBER IN ({sql_list(rep_codes)})
ORDER BY sr.SALESREP_NUMBER;
"""


def rep_based_title_sales_sql(period: str, rep_codes: list[str]) -> str:
    period1 = period[:4] + "01"
    rep_filter = sql_list(rep_codes)
    total_sales = (
        f"SUM(CASE WHEN rep.SALESREP_NUMBER IN ({rep_filter}) "
        "THEN sd.SalesNetValue ELSE 0 END)"
    )
    total_units = (
        f"SUM(CASE WHEN rep.SALESREP_NUMBER IN ({rep_filter}) "
        "THEN sd.SalesQty ELSE 0 END)"
    )
    rep_columns = []
    for rep_code in rep_codes:
        rep_sales = (
            f"SUM(CASE WHEN rep.SALESREP_NUMBER = '{rep_code}' "
            "THEN sd.SalesNetValue ELSE 0 END)"
        )
        rep_units = (
            f"SUM(CASE WHEN rep.SALESREP_NUMBER = '{rep_code}' "
            "THEN sd.SalesQty ELSE 0 END)"
        )
        rep_columns.extend(
            [
                f"    ,{rep_sales} AS [{rep_code} $]",
                f"    ,{rep_units} AS [{rep_code} units]",
                (
                    f"    ,CASE WHEN {total_sales} = 0 THEN 0 "
                    f"ELSE {rep_sales} / NULLIF({total_sales}, 0) END AS [{rep_code} %]"
                ),
            ]
        )

    return f"""
DECLARE @period1 varchar(6) = '{period1}';
DECLARE @periodytd varchar(6) = '{period}';

SELECT
    t.PUBLISHER_CODE AS [pub],
    t.PRODUCT_TYPE AS [pt],
    t.REPORTING_CATEGORY AS [cat],
    t.PUBLISHING_GROUP AS [pgr],
    t.ITEM_TITLE AS [isbn],
    t.SHORT_TITLE AS [title],
    t.PRICE_AMOUNT AS [price],
    t.AMORTIZATION_DATE AS [pub date],
    t.SEASON AS [sea],
    {total_sales} AS [$],
    {total_units} AS [units]
{chr(10).join(rep_columns)}
FROM summary.ItemBillToMonthlySales AS sd
    INNER JOIN ebs.Customer AS c
        ON sd.BILL_TO_SITE_USE_ID = c.SITE_USE_ID
    INNER JOIN ebs.SalesRep AS rep
        ON c.PRIMARY_SALESREP_ID = rep.SALESREP_ID
    INNER JOIN ebs.Item AS t
        ON sd.ITEM_ID = t.ITEM_ID
WHERE
    sd.PERIOD BETWEEN @period1 AND @periodytd
    AND rep.SALESREP_NUMBER IN ({rep_filter})
    AND sd.SALETYPECODE = 'n'
GROUP BY
    t.PUBLISHER_CODE,
    t.PRODUCT_TYPE,
    t.REPORTING_CATEGORY,
    t.PUBLISHING_GROUP,
    t.ITEM_TITLE,
    t.SHORT_TITLE,
    t.PRICE_AMOUNT,
    t.AMORTIZATION_DATE,
    t.SEASON
ORDER BY
    {total_sales} DESC;
"""


def national_specialty_title_sales_sql(period: str) -> str:
    period1 = period[:4] + "01"
    accounts = [
        ("Anthro", "sd.SSRROWID = '12'"),
        ("CC", "sd.SSRROWID = '31'"),
        ("Cost Plus", "sd.SSRROWID = '42'"),
        ("C&B", "sd.SSRROWID = '45'"),
        ("DB", "sd.SSRROWID = '50'"),
        ("FedEx", "sd.SSRROWID = '186'"),
        ("Fran", "sd.SSRROWID = '60'"),
        ("Fuego", "sd.SSRROWID = '61'"),
        ("Hobby", "sd.SSRROWID = '157'"),
        ("Pot", "sd.SSRROWID = '21'"),
        ("PBK", "sd.SSRROWID = '116'"),
        ("Spencer", "sd.SSRROWID = '134'"),
        ("Sub Box", "sd.SSRROWID = '204'"),
        ("Container", "sd.SSRROWID = '142'"),
        ("UG", "sd.SSRROWID = '147'"),
        ("Urban", "sd.SSRROWID = '151'"),
        ("West Bway", "sd.SSRROWID = '152'"),
        ("Indie Spec", "sd.SUBCHANNELID IN ('8', '23')"),
        ("Faire", "sd.SSRROWID = '307'"),
    ]
    total_sales = "SUM(CASE WHEN sd.CHANNELID = '5' THEN sd.SalesNetValue ELSE 0 END)"
    total_units = "SUM(CASE WHEN sd.CHANNELID = '5' THEN sd.SalesQty ELSE 0 END)"
    columns = []
    for label, condition in accounts:
        sales = f"SUM(CASE WHEN {condition} THEN sd.SalesNetValue ELSE 0 END)"
        units = f"SUM(CASE WHEN {condition} THEN sd.SalesQty ELSE 0 END)"
        columns.extend(
            [
                f"    ,{sales} AS [{label} $]",
                f"    ,{units} AS [{label} Units]",
                (
                    f"    ,CASE WHEN {total_sales} = 0 THEN 0 "
                    f"ELSE {sales} / NULLIF({total_sales}, 0) END AS [{label} %]"
                ),
            ]
        )
    columns.extend(
        [
            "    ,SUM(CASE WHEN sd.SSRROWID = '140' THEN sd.SalesNetValue ELSE 0 END) AS [Target NR $]",
            "    ,SUM(CASE WHEN sd.SSRROWID = '140' THEN sd.SalesQty ELSE 0 END) AS [Target NR Units]",
        ]
    )
    return f"""
DECLARE @period1 varchar(6) = '{period1}';
DECLARE @periodytd varchar(6) = '{period}';

SELECT
    t.PUBLISHER_CODE AS [pub],
    t.PRODUCT_TYPE AS [pt],
    t.REPORTING_CATEGORY AS [cat],
    t.PUBLISHING_GROUP AS [pgr],
    t.ITEM_TITLE AS [isbn],
    t.SHORT_TITLE AS [title],
    t.PRICE_AMOUNT AS [price],
    t.AMORTIZATION_DATE AS [pub date],
    t.SEASON AS [sea],
    {total_sales} AS [Total Specialty $],
    {total_units} AS [Total Units]
{chr(10).join(columns)}
FROM summary.TitleMonthlySales AS sd
    INNER JOIN ebs.Item AS t
        ON sd.ITEM_ID = t.ITEM_ID
WHERE
    sd.PERIOD BETWEEN @period1 AND @periodytd
    AND sd.SALETYPECODE = 'n'
    AND sd.CHANNELID IN ('3', '5')
    AND t.PUBLISHER_CODE <> 'Moleskine'
    AND sd.PRODUCT_TYPE IN ('BK', 'FT')
    AND t.AVAILABILITY_STATUS NOT IN ('OP', 'OSI')
GROUP BY
    t.PUBLISHER_CODE,
    t.PRODUCT_TYPE,
    t.REPORTING_CATEGORY,
    t.PUBLISHING_GROUP,
    t.ITEM_TITLE,
    t.SHORT_TITLE,
    t.PRICE_AMOUNT,
    t.AMORTIZATION_DATE,
    t.SEASON
ORDER BY
    {total_sales} DESC;
"""


def x_gap_title_sales_sql(period: str, account_label: str, ssr_row_id: str, include_subject: bool = False) -> str:
    periodty1 = period[:4] + "01"
    period_ly_fy = f"{int(period[:4]) - 1}12"
    subject_column = "    CAST('0' AS varchar(20)) AS [B&N Subject],\n" if include_subject else ""
    account_pct_alias = "B&N %" if account_label == "Barnes & Noble" else "AMAZ %"
    return f"""
DECLARE @periodty1 varchar(6) = '{periodty1}';
DECLARE @periodLYFY varchar(6) = '{period_ly_fy}';
DECLARE @periodTY varchar(6) = '{period}';
DECLARE @periodly1 varchar(6) = [dbo].fnCalcPeriod(@periodty1, -12);

SELECT
    t.PUBLISHER_CODE AS [pub],
    t.PRODUCT_TYPE AS [pt],
    t.REPORTING_CATEGORY AS [cat],
    t.PUBLISHING_GROUP AS [pgr],
{subject_column}    t.ITEM_TITLE AS [isbn],
    t.SHORT_TITLE AS [title],
    t.PRICE_AMOUNT AS [price],
    t.AMORTIZATION_DATE AS [pub date],
    t.SEASON AS [sea],
    SUM(CASE WHEN sd.PERIOD BETWEEN @periodty1 AND @periodTY AND sd.SSRROWID = '{ssr_row_id}' THEN sd.SalesNetValue ELSE 0 END) AS [{account_label} YTD $],
    SUM(CASE WHEN sd.PERIOD BETWEEN @periodty1 AND @periodTY AND sd.SSRROWID = '{ssr_row_id}' THEN sd.SalesQty ELSE 0 END) AS [{account_label} YTD Units],
    CASE
        WHEN SUM(CASE WHEN sd.PERIOD BETWEEN @periodty1 AND @periodTY THEN sd.SalesNetValue ELSE 0 END) = 0 THEN 0
        ELSE SUM(CASE WHEN sd.PERIOD BETWEEN @periodty1 AND @periodTY AND sd.SSRROWID = '{ssr_row_id}' THEN sd.SalesNetValue ELSE 0 END)
            / NULLIF(SUM(CASE WHEN sd.PERIOD BETWEEN @periodty1 AND @periodTY THEN sd.SalesNetValue ELSE 0 END), 0)
    END AS [{account_pct_alias}],
    SUM(CASE WHEN sd.PERIOD BETWEEN @periodty1 AND @periodTY THEN sd.SalesNetValue ELSE 0 END) AS [CB YTD $],
    SUM(CASE WHEN sd.PERIOD BETWEEN @periodty1 AND @periodTY THEN sd.SalesQty ELSE 0 END) AS [CB YTD Units],
    SUM(CASE WHEN sd.PERIOD BETWEEN @periodly1 AND @periodLYFY AND sd.SSRROWID = '{ssr_row_id}' THEN sd.SalesNetValue ELSE 0 END) AS [{account_label} FY $],
    SUM(CASE WHEN sd.PERIOD BETWEEN @periodly1 AND @periodLYFY AND sd.SSRROWID = '{ssr_row_id}' THEN sd.SalesQty ELSE 0 END) AS [{account_label} FY Units],
    CASE
        WHEN SUM(CASE WHEN sd.PERIOD BETWEEN @periodly1 AND @periodLYFY THEN sd.SalesNetValue ELSE 0 END) = 0 THEN 0
        ELSE SUM(CASE WHEN sd.PERIOD BETWEEN @periodly1 AND @periodLYFY AND sd.SSRROWID = '{ssr_row_id}' THEN sd.SalesNetValue ELSE 0 END)
            / NULLIF(SUM(CASE WHEN sd.PERIOD BETWEEN @periodly1 AND @periodLYFY THEN sd.SalesNetValue ELSE 0 END), 0)
    END AS [{account_pct_alias} FY],
    SUM(CASE WHEN sd.PERIOD BETWEEN @periodly1 AND @periodLYFY THEN sd.SalesNetValue ELSE 0 END) AS [CB FY $],
    SUM(CASE WHEN sd.PERIOD BETWEEN @periodly1 AND @periodLYFY THEN sd.SalesQty ELSE 0 END) AS [CB FY Units]
FROM summary.TitleMonthlySales AS sd
    INNER JOIN ebs.Item AS t
        ON sd.ITEM_ID = t.ITEM_ID
WHERE
    sd.PERIOD BETWEEN @periodly1 AND @periodTY
    AND sd.SALETYPECODE = 'n'
    AND t.ITEM_TYPE <> 'pack'
    AND sd.DISTRIBUTION_DIRECT = 'n'
    AND t.PRODUCT_TYPE IN ('bk', 'ft')
    AND t.PUBLISHER_CODE <> 'Moleskine'
GROUP BY
    t.PUBLISHER_CODE,
    t.PRODUCT_TYPE,
    t.REPORTING_CATEGORY,
    t.PUBLISHING_GROUP,
    t.ITEM_TITLE,
    t.SHORT_TITLE,
    t.PRICE_AMOUNT,
    t.AMORTIZATION_DATE,
    t.SEASON
ORDER BY
    SUM(CASE WHEN sd.PERIOD BETWEEN @periodty1 AND @periodTY AND sd.SSRROWID = '{ssr_row_id}' THEN sd.SalesNetValue ELSE 0 END) DESC;
"""
