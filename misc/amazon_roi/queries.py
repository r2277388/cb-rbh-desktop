def db_item() -> str:
    return """
    SELECT
        i.ISBN
        ,SHORT_TITLE Title
        ,i.PUBLISHER_CODE pub
        ,case
            when left(i.PUBLISHING_GROUP,3)='BAR' then 'BAR'
            when i.PUBLISHER_CODE = 'Chronicle' then i.PUBLISHING_GROUP
            else i.PUBLISHER_CODE
        end pgrp
        ,i.PRICE_AMOUNT price
		,i.PRODUCT_TYPE PT
		,CASE 
			WHEN i.FORMAT IN('EG', 'CL','DA' ,'WL') then 'CL'
			else i.FORMAT
		END [FT]
        ,SUBSTRING(i.Season, 6, LEN(i.Season)) AS SeasonOnly
        ,i.AMORTIZATION_DATE pub_date
        ,i.ROYALTY_FLAG RF
    FROM
        ebs.item i
    WHERE
        i.PRODUCT_TYPE IN('BK','FT')
        AND i.SEASON <> 'No Season Found'
        AND i.AMORTIZATION_DATE is not null
        AND i.PRICE_AMOUNT is not null
        AND i.PUBLISHER_CODE = 'Chronicle'
        AND i.PUBLISHER_CODE NOT IN('San Francisco Art Institute','Do Books','Moleskine'
            ,'Sager','Glam Media','FareArts','Princeton','AMMO Books','Driscolls','PQ Blackwell','AFO LLC')
    """