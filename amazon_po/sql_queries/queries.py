def item_sql() -> str:
    """
    Obtaining data from sql to get up-to-date meta data for the ISBN's being processed.
    
    """
    
    return """
        SELECT
            i.PUBLISHER_CODE Publisher
            ,i.PRODUCT_TYPE pt
            ,i.REPORTING_CATEGORY cat
            ,case
                when i.PUBLISHING_GROUP in('BAR-ENT','BAR-ART','BAR-FWN','BAR-LIF','BAR-CHL') then 'BAR'
                else i.PUBLISHING_GROUP
            end pgrp
            ,i.ITEM_TITLE ISBN
            ,i.SHORT_TITLE title
            ,i.PRICE_AMOUNT price
            ,convert(char,coalesce(convert(varchar,i.AMORTIZATION_DATE,101),shdt.shipdate),101) pub
            ,case
                when i.AMORTIZATION_DATE is not null then year(i.AMORTIZATION_DATE)
                when substring(i.season,1,4) <> 'No S' then substring(i.season,1,4)
                else year(getdate())
            end [year]

        FROM
            ebs.Item i
            left join (
                Select tt.ean13 [ISBN],tt.active_datevalue [SHIPDATE]
                from tmm.cb_Import_Title_Tasks tt
                Where tt.date_desc = 'On Sale Date' AND tt.active_datevalue is not null AND tt.printingnumber = 1
                ) shdt on shdt.ISBN = i.ISBN
        WHERE
            i.PUBLISHER_CODE NOT IN('PQ Blackwell','San Francisco Art Institute','Driscolls','In Active')
            AND i.PUBLISHING_GROUP <> '???'
            AND i.SHORT_TITLE is not NULL
            AND i.PRICE_AMOUNT is not NULL
	        AND i.PRODUCT_TYPE IN('BK','FT')
        """