def saldet():  
    return '''
    SELECT
        i.ISBN
        ,i.PUBLISHER_CODE pub,
        CASE
            WHEN i.PRODUCT_TYPE IN('RP','CP') THEN 'CPRP'
            WHEN i.PUBLISHING_GROUP IN('LIF','FWN') THEN 'FLS'
            WHEN i.PUBLISHING_GROUP IN('ART','CHL','ENT','MOD') THEN i.PUBLISHING_GROUP
            WHEN i.PUBLISHING_GROUP IN('BAR-ART','BAR-ENT','BAR-LIF') THEN 'BAR'
            WHEN i.PUBLISHING_GROUP = 'DOB' THEN 'DO'
            WHEN i.PUBLISHER_CODE = 'Tourbillon' THEN 'TW'
            WHEN i.PUBLISHER_CODE = 'Sierra Club' THEN 'SC'
            WHEN i.PUBLISHER_CODE = 'Levine Querido' THEN 'LQ'
            WHEN i.PUBLISHER_CODE IN('Glam Media','Benefit','PQ Blackwell','San Francisco Art Institute','AFO LLC','FareArts','Sager') THEN 'CD'
            WHEN i.PUBLISHER_CODE = 'Creative Company' THEN 'CC'
            WHEN i.PUBLISHER_CODE = 'GALISON' AND stie.DISTRIBUTION_DIRECT = 'Y' THEN 'GAD'
            WHEN i.PUBLISHING_GROUP = 'GAL' THEN 'GA'
            WHEN i.PUBLISHING_GROUP = 'GAL-CL' THEN 'CL'
            WHEN i.PUBLISHING_GROUP = 'MUD' THEN 'MP'
            WHEN i.PUBLISHING_GROUP IN('LAU-BIS') THEN 'LKBS'
            WHEN i.PUBLISHER_CODE = 'Laurence King' AND i.PRODUCT_TYPE = 'FT' THEN 'LKGI'
            WHEN i.PUBLISHER_CODE = 'Laurence King' AND i.PRODUCT_TYPE <> 'FT' THEN 'LKBK'
            WHEN i.PUBLISHER_CODE = 'PRINCETON' AND i.PUBLISHING_GROUP IN('PAP-MS','PAP-MP') THEN 'PAMS'
            WHEN i.PUBLISHER_CODE = 'Princeton' AND i.PRODUCT_TYPE = 'FT' THEN 'PAGI'
            WHEN i.PUBLISHER_CODE = 'Princeton' AND i.PRODUCT_TYPE <> 'FT' THEN 'PABK'
            WHEN i.PUBLISHING_GROUP = 'MOL-PLANNER' THEN 'MSPL'
            WHEN i.PUBLISHING_GROUP IN('MOL-NON PAPER','MOL-PAPER') THEN 'MSCO'
            WHEN i.PUBLISHER_CODE = 'Hardie Grant Publishing' THEN 'HG'
            ELSE i.PUBLISHING_GROUP        
        END AS tab
        ,ssr.Description ssr_row
        ,subchan.Description subchannel
        ,chan.Description channel
        ,datediff(month,i.amortization_date,getdate()) months_old
        ,SUM(CASE WHEN sd.TRX_DATE < dateadd(year,1,i.amortization_date) THEN sd.QUANTITY_INVOICED ELSE 0 END) AS sales_y1
        ,count(distinct billto.PARTYSITENUMBER) acct_cnt
        ,count(distinct i.isbn) isbn_cnt
    FROM
        ebs.Sales sd
        INNER JOIN ebs.Item i ON i.ITEM_ID = sd.ITEM_ID
        INNER JOIN ssr.SalesSSRRow stie ON stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
        inner join ssr.SSRRow as ssr on stie.SSRRowID=ssr.SSRRowID	
        inner join ssr.SubChannel as subchan on ssr.SubChannelID=subchan.SubChannelID	
        inner join ssr.Channel as chan on subchan.ChannelID=chan.ChannelID
        inner join ebs.Customer shipto on shipto.SITE_USE_ID = sd.SHIP_TO_SITE_USE_ID
        left join ebs.customer billto on billto.SITE_USE_ID = shipto.BILL_TO_SITE_USE_ID

    WHERE 
        sd.PERIOD >= '201501'
        and i.AMORTIZATION_DATE between '2015-01-01' and '2023-08-31'
        AND i.PRODUCT_TYPE IN('BK','FT')
        AND sd.INVOICE_LINE_TYPE = 'SALE'
        AND cbq2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N'
        AND i.PUBLISHER_CODE NOT IN('ZZZ','MKT')
        AND i.isbn is not null
    GROUP BY
        i.isbn
        ,i.PUBLISHER_CODE,
        CASE
            WHEN i.PRODUCT_TYPE IN('RP','CP') THEN 'CPRP'
            WHEN i.PUBLISHING_GROUP IN('LIF','FWN') THEN 'FLS'
            WHEN i.PUBLISHING_GROUP IN('ART','CHL','ENT','MOD') THEN i.PUBLISHING_GROUP
            WHEN i.PUBLISHING_GROUP IN('BAR-ART','BAR-ENT','BAR-LIF') THEN 'BAR'
            WHEN i.PUBLISHING_GROUP = 'DOB' THEN 'DO'
            WHEN i.PUBLISHER_CODE = 'Tourbillon' THEN 'TW'
            WHEN i.PUBLISHER_CODE = 'Sierra Club' THEN 'SC'
            WHEN i.PUBLISHER_CODE = 'Levine Querido' THEN 'LQ'
            WHEN i.PUBLISHER_CODE IN('Glam Media','Benefit','PQ Blackwell','San Francisco Art Institute','AFO LLC','FareArts','Sager') THEN 'CD'
            WHEN i.PUBLISHER_CODE = 'Creative Company' THEN 'CC'
            WHEN i.PUBLISHER_CODE = 'GALISON' AND stie.DISTRIBUTION_DIRECT = 'Y' THEN 'GAD'
            WHEN i.PUBLISHING_GROUP = 'GAL' THEN 'GA'
            WHEN i.PUBLISHING_GROUP = 'GAL-CL' THEN 'CL'
            WHEN i.PUBLISHING_GROUP = 'MUD' THEN 'MP'
            WHEN i.PUBLISHING_GROUP IN('LAU-BIS') THEN 'LKBS'
            WHEN i.PUBLISHER_CODE = 'Laurence King' AND i.PRODUCT_TYPE = 'FT' THEN 'LKGI'
            WHEN i.PUBLISHER_CODE = 'Laurence King' AND i.PRODUCT_TYPE <> 'FT' THEN 'LKBK'
            WHEN i.PUBLISHER_CODE = 'PRINCETON' AND i.PUBLISHING_GROUP IN('PAP-MS','PAP-MP') THEN 'PAMS'
            WHEN i.PUBLISHER_CODE = 'Princeton' AND i.PRODUCT_TYPE = 'FT' THEN 'PAGI' 
            WHEN i.PUBLISHER_CODE = 'Princeton' AND i.PRODUCT_TYPE <> 'FT' THEN 'PABK' 
            WHEN i.PUBLISHING_GROUP = 'MOL-PLANNER' THEN 'MSPL' 
            WHEN i.PUBLISHING_GROUP IN('MOL-NON PAPER','MOL-PAPER') THEN 'MSCO' 
            WHEN i.PUBLISHER_CODE = 'Hardie Grant Publishing' THEN 'HG'
            ELSE i.PUBLISHING_GROUP
        END,
        ssr.Description,
        subchan.Description,
        chan.Description
        ,datediff(month,i.amortization_date,getdate())
    HAVING SUM(CASE WHEN sd.TRX_DATE < dateadd(year,1,i.amortization_date) THEN sd.QUANTITY_INVOICED ELSE 0 END) > 100
    '''