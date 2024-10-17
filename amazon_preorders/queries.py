def item_sql():
    return """
    SELECT
		i.ISBN
		,i.SHORT_TITLE title
		,i.PUBLISHER_CODE publisher
		-- ,case                             
		--     when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'                      
		--     when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'                     
		--     when i.PUBLISHER_CODE IN('Glam Media','Benefit','PQ Blackwell','San Francisco Art Institute','AFO LLC','FareArts','Sager') then 'CD'                 
		--     when i.PUBLISHER_CODE = 'Creative Company' then 'CC'   
		--     when i.PUBLISHER_CODE = 'Do Books' then 'DO'
		--     when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
		--     when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'                                           
		--     when i.PUBLISHING_GROUP = 'GAL' then 'GA'                                                      
		--     when i.PUBLISHING_GROUP = 'GAL-CL' then 'CL'                        
		--     when i.PUBLISHING_GROUP = 'MUD' then 'MP'
		--     when i.PUBLISHING_GROUP = 'GAL-BM' then 'BM'             
		--     when i.PUBLISHING_GROUP in('LAU-BIS') then 'LKBS'                          
		--     when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE = 'FT' then 'LKGI'                      
		--     when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE <> 'FT' then 'LKBK'         
		--     when i.PUBLISHER_CODE = 'Hardie Grant Publishing' then 'HG'  
		--     when i.PUBLISHING_GROUP in('BAR-ART','BAR-ENT','BAR-LIF') then 'BAR'                  
		--     else i.PUBLISHING_GROUP                 
		--  end pgrp
    FROM                
        ebs.Item i
    WHERE
        i.PRODUCT_TYPE in('BK','FT','DI')
        --AND i.AVAILABILITY_STATUS not in('OP','WIT','OPR','NOP','OSI','PC','DIS','CS','POS')
        AND i.AVAILABILITY_STATUS is not null

    """