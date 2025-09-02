# -*- coding: utf-8 -*-
"""
Created on Sat Aug  3 18:36:17 2024

@author: RBH
"""

def query1(prior_day,tp,tply):
    '''
    Returns
    -------
    Prior Day and MTD sales by publishing groups broken up by core and non-core

    '''

    return f'''
        SELECT
            case
                when i.PUBLISHER_CODE = 'Chronicle' then 'CB'
                else 'DP'
            END CBDP
            ,case
                when i.PRODUCT_TYPE  in('RP','CP') then 'CPRP'
                when i.PUBLISHING_GROUP in('LIF','FWN') then 'FLS'
                when i.PUBLISHING_GROUP in('ART','CHL','ENT','MOD') then i.PUBLISHING_GROUP
                when i.PUBLISHING_GROUP in('BAR-ART','BAR-ENT','BAR-LIF') then 'BAR'
                when i.PUBLISHING_GROUP = 'DOB' then 'DO'
                when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'
                when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'
                when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
                when i.PUBLISHER_CODE IN('Glam Media','Benefit','PQ Blackwell','San Francisco Art Institute','AFO LLC','FareArts','Sager') then 'CD'
        		when i.PUBLISHER_CODE = 'Creative Company' then 'CC'            
                WHEN i.PUBLISHER_CODE = 'GALISON' AND stie.DISTRIBUTION_DIRECT = 'Y' then 'GAD'  
                when i.PUBLISHING_GROUP = 'GAL' then 'GA'   
                when i.PUBLISHING_GROUP = 'GAL-CL' then 'CL'       
                when i.PUBLISHING_GROUP = 'MUD' then 'MP'
                when i.PUBLISHING_GROUP in('LAU-BIS') then 'LKBS'          
                when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE = 'FT' then 'LKGI'           
                when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE <> 'FT' then 'LKBK' 
                when i.PUBLISHER_CODE = 'PRINCETON' and i.PUBLISHING_GROUP in('PAP-MS','PAP-MP') then 'PAMS' 
                when i.PUBLISHER_CODE = 'Princeton' and i.PRODUCT_TYPE = 'FT' then 'PAGI'          
                when i.PUBLISHER_CODE = 'Princeton' and i.PRODUCT_TYPE <> 'FT' then 'PABK' 
                when i.PUBLISHING_GROUP = 'MOL-PLANNER' then 'MSPL'      
                when i.PUBLISHING_GROUP in('MOL-NON PAPER','MOL-PAPER') then 'MSCO' 
                when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
		        when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'
                else i.PUBLISHING_GROUP        
            end tab
            ,sum(case when i.PRODUCT_TYPE in('BK','FT') AND TRX_DATE = '{prior_day}' then sd.REVENUE_AMOUNT else 0 end) [PriorDay_Core]
            ,sum(case when i.PRODUCT_TYPE in('BK','FT') AND sd.period = '{tp}' then sd.REVENUE_AMOUNT else 0 end) tp_core
            ,sum(case when i.PRODUCT_TYPE in('BK','FT') AND sd.period = '{tply}' then sd.REVENUE_AMOUNT else 0 end) tply_core
            ,sum(case when i.PRODUCT_TYPE not in('BK','FT') AND TRX_DATE = '{prior_day}' then sd.REVENUE_AMOUNT else 0 end) [PriorDay_NonCore]
            ,sum(case when i.PRODUCT_TYPE not in('BK','FT') AND sd.period = '{tp}' then sd.REVENUE_AMOUNT else 0 end) tp_NonCore
            ,sum(case when i.PRODUCT_TYPE not in('BK','FT') AND sd.period = '{tply}' then sd.REVENUE_AMOUNT else 0 end) tply_NonCore
        FROM
            ebs.Sales sd
            inner join ssr.SalesSSRRow stie on stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
            INNER JOIN ebs.Item i on i.ITEM_ID = sd.ITEM_ID
        WHERE 
            Sd.PERIOD in('{tply}','{tp}')
            AND sd.INVOICE_LINE_TYPE='SALE'
            AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID)='N'
            AND i.PUBLISHER_CODE not in('ZZZ','MKT')
        GROUP BY
            case
                when i.PUBLISHER_CODE = 'Chronicle' then 'CB'
                else 'DP'
            END
            ,case
                when i.PRODUCT_TYPE  in('RP','CP') then 'CPRP'
                when i.PUBLISHING_GROUP in('LIF','FWN') then 'FLS'
                when i.PUBLISHING_GROUP in('ART','CHL','ENT','MOD') then i.PUBLISHING_GROUP
                when i.PUBLISHING_GROUP in('BAR-ART','BAR-ENT','BAR-LIF') then 'BAR'
                when i.PUBLISHING_GROUP = 'DOB' then 'DO'
                when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'
                when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'
                when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
                when i.PUBLISHER_CODE IN('Glam Media','Benefit','PQ Blackwell','San Francisco Art Institute','AFO LLC','FareArts','Sager') then 'CD'
                when i.PUBLISHER_CODE = 'Creative Company' then 'CC'
                WHEN i.PUBLISHER_CODE = 'GALISON' AND stie.DISTRIBUTION_DIRECT = 'Y' then 'GAD'
                when i.PUBLISHING_GROUP = 'GAL' then 'GA'
                when i.PUBLISHING_GROUP = 'GAL-CL' then 'CL' 
                when i.PUBLISHING_GROUP = 'MUD' then 'MP'
                when i.PUBLISHING_GROUP in('LAU-BIS') then 'LKBS'
                when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE = 'FT' then 'LKGI'
                when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE <> 'FT' then 'LKBK'
                when i.PUBLISHER_CODE = 'PRINCETON' and i.PUBLISHING_GROUP in('PAP-MS','PAP-MP') then 'PAMS'
                when i.PUBLISHER_CODE = 'Princeton' and i.PRODUCT_TYPE = 'FT' then 'PAGI' 
                when i.PUBLISHER_CODE = 'Princeton' and i.PRODUCT_TYPE <> 'FT' then 'PABK' 
                when i.PUBLISHING_GROUP = 'MOL-PLANNER' then 'MSPL' 
                when i.PUBLISHING_GROUP in('MOL-NON PAPER','MOL-PAPER') then 'MSCO' 
                when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
		        when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'
                 else i.PUBLISHING_GROUP
            end'''


def query2(tp,prior_day):
    '''
    Top 10 SSR Accounts for prior day   
    '''
    
    return f'''
    SELECT TOP 10
        ssr_row.Description ssr_row
        ,sum(case when t.PUBLISHER_CODE = 'Chronicle' then sd.REVENUE_AMOUNT else 0 end) [Chronicle]
        ,sum(case when not(t.PUBLISHER_CODE = 'Chronicle') then sd.REVENUE_AMOUNT else 0 end) [Distribution]
        ,sum(sd.revenue_amount) [Total]
    FROM
        ebs.Sales sd
        inner join ssr.SalesSSRRow stie on stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
        inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID= stie.SSRRowID
        INNER JOIN ebs.Item t on t.ITEM_ID=sd.ITEM_ID
        INNER JOIN ssr.SubChannel subchan on subchan.SubChannelID = ssr_row.SubChannelID
        inner join ssr.Channel chan on chan.ChannelID = subchan.ChannelID
    
    WHERE 
        Sd.PERIOD = '{tp}'
        AND TRX_DATE = '{prior_day}'
        AND sd.INVOICE_LINE_TYPE='SALE'
        AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID)='N'
        and t.PRODUCT_TYPE in('BK','FT')
    GROUP BY
        ssr_row.Description
    ORDER BY
        sum(sd.REVENUE_AMOUNT) desc
    '''

def query3(tp,prior_day):
    '''
    Returns
    -------
    sql query
        Top prior day top title sales by CB and DP

    '''

    return f'''
    SELECT
        case
            when t.PUBLISHER_CODE = 'Chronicle' then 'cb'
            else 'dp'
        end div
        ,t.ISBN
        ,t.SHORT_TITLE [Title]
        ,t.PRICE_AMOUNT [Price]
        ,case
            when t.PUBLISHING_GROUP in('LIF','FWN') then 'FLS'
            else t.PUBLISHING_GROUP
        end pgc
        ,sum(sd.QUANTITY_INVOICED) [Units]
        ,sum(sd.REVENUE_AMOUNT) [Dollars]
        ,ROW_NUMBER() over(PARTITION BY case when t.PUBLISHER_CODE = 'Chronicle' then 'cb' else 'dp' end ORDER BY sum(sd.REVENUE_AMOUNT) desc) 'order'
    
    FROM
        ebs.Sales sd
        INNER JOIN ebs.Item t on t.ITEM_ID=sd.ITEM_ID
        INNER JOIN ebs.Customer c on c.SITE_USE_ID=sd.SHIP_TO_SITE_USE_ID
        INNER JOIN ebs.SalesRep sr on sr.SALESREP_ID=sd.PRIMARY_SALESREP_ID
        LEFT OUTER JOIN ebs.Customer BillTo ON BillTo.SITE_USE_ID = c.BILL_TO_SITE_USE_ID
    WHERE
        Sd.PERIOD = '{tp}'
        AND sd.TRX_DATE= '{prior_day}'
        AND sd.INVOICE_LINE_TYPE='SALE'
        AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID)='N'
        AND t.PRICE_AMOUNT <> 0
        AND sr.SALESREP_NUMBER not in ('1055','1061','2055')
        AND not(t.FORMAT in ('EB', 'EN', 'PP', 'PS', 'SA', 'AP'))
        AND not(t.PRODUCT_TYPE in ('DI','PZ','?'))
    
    GROUP BY 
        case
            when t.PUBLISHER_CODE = 'Chronicle' then 'cb'
            else 'dp'
        end
        ,t.ISBN
        ,t.SHORT_TITLE
        ,t.PRICE_AMOUNT
        ,case
            when t.PUBLISHING_GROUP in('LIF','FWN') then 'FLS'
            else t.PUBLISHING_GROUP
        end
        '''

def query4(tp):
    '''
    Top 10 MTD titles by CB and DP
    '''

    return f'''
    SELECT
        case
            when t.PUBLISHER_CODE = 'Chronicle' then 'cb'
            else 'dp'
        end div
        ,t.ISBN
        ,t.SHORT_TITLE [Title]
        ,t.PRICE_AMOUNT [Price]
        ,case
            when t.PUBLISHING_GROUP in('LIF','FWN') then 'FLS'
            else t.PUBLISHING_GROUP
        end pgc
        ,sum(sd.QUANTITY_INVOICED) [Units]
        ,sum(sd.REVENUE_AMOUNT) [Dollars]
        ,ROW_NUMBER() over(PARTITION BY case when t.PUBLISHER_CODE = 'Chronicle' then 'cb' else 'dp' end ORDER BY sum(sd.REVENUE_AMOUNT) desc) 'order'
    
    FROM
        ebs.Sales sd
        INNER JOIN ebs.Item t on t.ITEM_ID=sd.ITEM_ID
        INNER JOIN ebs.Customer c on c.SITE_USE_ID=sd.SHIP_TO_SITE_USE_ID
        INNER JOIN ebs.SalesRep sr on sr.SALESREP_ID=sd.PRIMARY_SALESREP_ID
        LEFT OUTER JOIN ebs.Customer BillTo ON BillTo.SITE_USE_ID = c.BILL_TO_SITE_USE_ID
    WHERE
        Sd.PERIOD = '{tp}'
        AND sd.INVOICE_LINE_TYPE='SALE'
        AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID)='N'
        AND t.PRICE_AMOUNT <> 0
        AND sr.SALESREP_NUMBER not in ('1055','1061','2055')
        AND not(t.FORMAT in ('EB', 'EN', 'PP', 'PS', 'SA', 'AP'))
        AND not(t.PRODUCT_TYPE in ('DI','PZ','?'))
    
    GROUP BY 
        case
            when t.PUBLISHER_CODE = 'Chronicle' then 'cb'
            else 'dp'
        end
        ,t.ISBN
        ,t.SHORT_TITLE
        ,t.PRICE_AMOUNT
        ,case
            when t.PUBLISHING_GROUP in('LIF','FWN') then 'FLS'
            else t.PUBLISHING_GROUP
        end
        '''

def query5(tp): 
    '''
    Napkin Top Forecast query
    '''
    
    return f'''
    SELECT
        case
            when i.PUBLISHING_GROUP in('BAR-ART','BAR-ENT','BAR-FWN','BAR-LIF') then 'BAR'
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
            else i.PUBLISHING_GROUP
        end pgrp
        ,sum(case when SSR_row.SSRRowID in('32','146') then sd.REVENUE_AMOUNT else 0 end) [cons]
        ,sum(case when SSR_row.SSRRowID in('6') then sd.REVENUE_AMOUNT else 0 end) [amaz]
    
    FROM
        ebs.Sales sd
        inner join ssr.SalesSSRRow stie on stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
            inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID= stie.SSRRowID
            INNER JOIN ebs.Item i on i.ITEM_ID = sd.ITEM_ID
    
    WHERE
        Sd.PERIOD = '{tp}'
        AND sd.INVOICE_LINE_TYPE='SALE'
        AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID)='N'
        AND i.PUBLISHING_GROUP not in('ZZZ','MKT')
        and i.PUBLISHER_CODE = 'Chronicle'
        AND SSR_row.SSRRowID in('32','146','6')
    GROUP BY
        case
            when i.PUBLISHING_GROUP in('BAR-ART','BAR-ENT','BAR-FWN','BAR-LIF') then 'BAR'
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
            else i.PUBLISHING_GROUP
        end
    ORDER BY
        pgrp
        '''

def query6(prior_day):
    '''
    Returns
    -------
    Order Query - Top 10 CB Prior Day SSR_ROW Orders

    '''

    return f'''
    SELECT top 10
        chan.Description channel
        ,ssr_row.Description SSR_Row
        ,sum(ho.Quantity) qty
    FROM
        hachette.HachetteOrders ho
            inner join ebs.item i on i.ITEM_TITLE = ho.isbn
            inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID = ho.SSRRowID
            inner join ssr.SubChannel subchan on subchan.SubChannelID = ssr_row.SubChannelID
            inner join ssr.Channel chan on chan.ChannelID = subchan.ChannelID
    WHERE
        i.PRICE_AMOUNT > 0
        and ho.EnteredDate = '{prior_day}' 
        and i.PUBLISHER_CODE = 'Chronicle'
        GROUP BY
        chan.Description
        ,ssr_row.Description
    ORDER BY
        sum(ho.Quantity) desc
        '''

def query7(prior_day):
    '''
    Returns
    -------
    Order Query - Top 10 DP Prior Day SSR_ROW Orders

    '''

    return f'''
    SELECT top 10
        chan.Description channel
        ,ssr_row.Description SSR_Row
        ,sum(ho.Quantity) qty
    FROM
        hachette.HachetteOrders ho
            inner join ebs.item i on i.ITEM_TITLE = ho.isbn
            inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID = ho.SSRRowID
            inner join ssr.SubChannel subchan on subchan.SubChannelID = ssr_row.SubChannelID
            inner join ssr.Channel chan on chan.ChannelID = subchan.ChannelID
    WHERE
        i.PRICE_AMOUNT > 0
        and ho.EnteredDate = '{prior_day}' 
        and i.PUBLISHER_CODE != 'Chronicle'
        GROUP BY
        chan.Description
        ,ssr_row.Description
    ORDER BY
        sum(ho.Quantity) desc
        '''

def query8(prior_day):
    '''
    Returns
    -------
    Order Query - Top 10 CB Prior Day Title Orders

    '''
    return f'''
    
    SELECT top 10
        i.PUBLISHING_GROUP pgrp
        ,i.ISBN
        ,i.SHORT_TITLE
        ,sum(ho.Quantity) qty
    FROM
        hachette.HachetteOrders ho
        inner join ebs.item i on i.ITEM_TITLE = ho.isbn
        inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID = ho.SSRRowID
        inner join ssr.SubChannel subchan on subchan.SubChannelID = ssr_row.SubChannelID
        inner join ssr.Channel chan on chan.ChannelID = subchan.ChannelID
    WHERE
        i.PRICE_AMOUNT > 0
        AND i.PUBLISHER_CODE = 'Chronicle'
        and ho.EnteredDate = '{prior_day}' 
    
    GROUP BY
        i.PUBLISHING_GROUP
        ,i.ISBN
        ,i.SHORT_TITLE
        ORDER BY
        sum(ho.Quantity) desc
    '''

def query9(prior_day):
    '''
    Returns
    -------
    Order Query - Top 10 CB Prior Day Title Orders

    '''

    return f'''
    SELECT top 10
        i.PUBLISHER_CODE pub
        ,i.ISBN
        ,i.SHORT_TITLE
        ,sum(ho.Quantity) qty
    FROM
        hachette.HachetteOrders ho
        inner join ebs.item i on i.ITEM_TITLE = ho.isbn
        inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID = ho.SSRRowID
        inner join ssr.SubChannel subchan on subchan.SubChannelID = ssr_row.SubChannelID
        inner join ssr.Channel chan on chan.ChannelID = subchan.ChannelID
    WHERE
        i.PRICE_AMOUNT > 0
        AND i.PUBLISHER_CODE != 'Chronicle'
        and ho.EnteredDate = '{prior_day}' 
    GROUP BY
        i.PUBLISHER_CODE
        ,i.ISBN
        ,i.SHORT_TITLE
        ORDER BY
        sum(ho.Quantity) desc
        '''

def query10(typ1):
    '''
    Returns
    -------
    SQL Query
        YTD CB / DP Sales

    '''

    return f'''                 
    SELECT                     
        case
            when i.PUBLISHER_CODE = 'Chronicle' then 'CB'
            else 'DP'
        end div
        ,case
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
            when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
            when i.PUBLISHER_CODE = 'Laurence King' then 'LK'
            when i.PUBLISHER_CODE = 'Creative Company' then 'CC'
            when i.PUBLISHER_CODE = 'Princeton' then 'PA'
            when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
	        when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'
            when i.PUBLISHER_CODE = 'Galison' then 'GA'
            when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'
            when left(i.PUBLISHING_GROUP,3) = 'BAR' then 'BAR'
            when i.PUBLISHER_CODE = 'Do Books' then 'DO'
            when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'
            when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'
            else i.PUBLISHING_GROUP
        end pgrp
        ,case
            when i.PRODUCT_TYPE in('BK','FT') then 'core'
            when i.PRODUCT_TYPE in('RP','CP') then 'cprp'
            else i.PRODUCT_TYPE
        end pt
        ,sum(ims.SalesNetValue) val               
    FROM                       
        [summary].TitleMonthlySales ims                 
                inner join ebs.Item i on ims.ITEM_ID = i.ITEM_ID              
    WHERE                      
        ims.PERIOD >= '{typ1}'              
        AND ims.SALETYPECODE = 'N'               
        AND i.PRODUCT_TYPE in ('BK','FT','CP','RP','DI','PZ')  
        and i.PUBLISHING_GROUP not in('MKT','PQB','GLM')
        and i.PUBLISHER_CODE not in('Moleskine','Benefit')
        and ims.DISTRIBUTION_DIRECT = 'N'
    GROUP BY                          
        case
            when i.PUBLISHER_CODE = 'Chronicle' then 'CB'
            else 'DP'
        end
        ,case
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
            when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
            when i.PUBLISHER_CODE = 'Laurence King' then 'LK'
            when i.PUBLISHER_CODE = 'Creative Company' then 'CC'
            when i.PUBLISHER_CODE = 'Princeton' then 'PA'
            when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
	        when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'
            when i.PUBLISHER_CODE = 'Galison' then 'GA'
            when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'
            when left(i.PUBLISHING_GROUP,3) = 'BAR' then 'BAR'
            when i.PUBLISHER_CODE = 'Do Books' then 'DO'
            when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'
            when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'
            else i.PUBLISHING_GROUP
        end
        ,case
            when i.PRODUCT_TYPE in('BK','FT') then 'core'
            when i.PRODUCT_TYPE in('RP','CP') then 'cprp'
            else i.PRODUCT_TYPE
        end
    '''

def query10_b(typ1):
    '''
    Returns
    -------
    YTD CB / DP Sales but using the ebs.sales table

    '''
    
    return f'''
    SELECT
        case
            when i.PUBLISHER_CODE = 'Chronicle' then 'CB'
            else 'DP'
        end div
        ,case
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
            when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
            when i.PUBLISHER_CODE = 'Laurence King' then 'LK'
            when i.PUBLISHER_CODE = 'Creative Company' then 'CC'
            when i.PUBLISHER_CODE = 'Princeton' then 'PA'
            when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
	    when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'
            when i.PUBLISHER_CODE = 'Galison' then 'GA'
            when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'
            when left(i.PUBLISHING_GROUP,3) = 'BAR' then 'BAR'
            when i.PUBLISHER_CODE = 'Do Books' then 'DO'
            when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'
            when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'
            else i.PUBLISHING_GROUP
        end pgrp
        ,case
            when i.PRODUCT_TYPE in('BK','FT') then 'core'
            when i.PRODUCT_TYPE in('RP','CP') then 'cprp'
            else i.PRODUCT_TYPE
        end pt
        ,sum(sd.REVENUE_AMOUNT) val 
    
    FROM
        ebs.Sales sd
        inner join ssr.SalesSSRRow stie on stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
            inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID= stie.SSRRowID
            INNER JOIN ebs.Item i on i.ITEM_ID = sd.ITEM_ID
    
        WHERE
            Sd.PERIOD >= '{typ1}'
            AND sd.INVOICE_LINE_TYPE='SALE'
            AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID)='N'
            AND i.PRODUCT_TYPE in ('BK','FT','CP','RP','DI','PZ')  
            and i.PUBLISHING_GROUP not in('MKT','PQB','GLM')
            and i.PUBLISHER_CODE not in('Moleskine','Benefit')
            and stie.DISTRIBUTION_DIRECT = 'N'
        GROUP BY
        case
            when i.PUBLISHER_CODE = 'Chronicle' then 'CB'
            else 'DP'
        end
        ,case
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
            when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
            when i.PUBLISHER_CODE = 'Laurence King' then 'LK'
            when i.PUBLISHER_CODE = 'Creative Company' then 'CC'
            when i.PUBLISHER_CODE = 'Princeton' then 'PA'
            when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
	    when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'
            when i.PUBLISHER_CODE = 'Galison' then 'GA'
            when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'
            when left(i.PUBLISHING_GROUP,3) = 'BAR' then 'BAR'
            when i.PUBLISHER_CODE = 'Do Books' then 'DO'
            when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'
            when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'
            else i.PUBLISHING_GROUP
        end
        ,case
            when i.PRODUCT_TYPE in('BK','FT') then 'core'
            when i.PRODUCT_TYPE in('RP','CP') then 'cprp'
            else i.PRODUCT_TYPE
        end
        '''

def query11(typ1):
    '''
    SSR YTD Sales
    '''
    
    return f'''                 
    SELECT
        ssr_row.Description ssr
        ,case
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
            when left(i.PUBLISHING_GROUP,3) = 'BAR' then 'FLS'
            when i.PUBLISHER_CODE = 'Chronicle' then i.PUBLISHING_GROUP
            when i.PUBLISHER_CODE = 'Laurence King' then 'LK'
            when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
	        when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'
            when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
            when i.PUBLISHER_CODE = 'Princeton' then 'PA'
            when i.PUBLISHER_CODE = 'Laurence King' then 'LK'
            when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'
            when i.PUBLISHER_CODE = 'Creative Company' then 'CC'
            when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'
            when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'
            when i.PUBLISHING_GROUP = 'GAL' then 'GA'
            when i.PUBLISHING_GROUP = 'GAL' then 'GA'
            when i.PUBLISHING_GROUP = 'MUD' then 'MP'
            when i.PUBLISHING_GROUP = 'GAL-CL' then 'CL'
            when i.PUBLISHING_GROUP = 'GAL-BM' then 'BM'
            when i.PUBLISHER_CODE = 'Do Books' then 'DO'
            else i.PUBLISHING_GROUP
        end pbgrp
        ,sum(sd.REVENUE_AMOUNT) val
        ,sum(sd.QUANTITY_INVOICED) units
    FROM ebs.Sales sd
        inner join ssr.SalesSSRRow stie on stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID
        inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID= stie.SSRRowID
        INNER JOIN ebs.Item i on i.ITEM_ID=sd.ITEM_ID  
        inner join ssr.SubChannel subchan on subchan.SubChannelID = ssr_row.SubChannelID
        inner join ssr.Channel chan on chan.ChannelID = subchan.ChannelID
        inner join ebs.SalesRep sr on sr.SALESREP_ID = sd.PRIMARY_SALESREP_ID
    WHERE
        sd.PERIOD >= '{typ1}'
        and cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID)='N'   
        AND sd.INVOICE_LINE_TYPE = 'SALE'
        and i.PRODUCT_TYPE in('FT','BK')
        and i.PUBLISHER_CODE not in('Moleskine','Benefit','Glam Media','PQ Blackwell')
        and stie.DISTRIBUTION_DIRECT = 'N'
    GROUP BY
        ssr_row.Description
        ,case
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'
            when left(i.PUBLISHING_GROUP,3) = 'BAR' then 'FLS'
            when i.PUBLISHER_CODE = 'Chronicle' then i.PUBLISHING_GROUP
            when i.PUBLISHER_CODE = 'Laurence King' then 'LK'
            when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
	        when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'
            when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
            when i.PUBLISHER_CODE = 'Princeton' then 'PA'
            when i.PUBLISHER_CODE = 'Laurence King' then 'LK'
            when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'
            when i.PUBLISHER_CODE = 'Creative Company' then 'CC'
            when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'
            when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'
            when i.PUBLISHING_GROUP = 'GAL' then 'GA'
            when i.PUBLISHING_GROUP = 'GAL' then 'GA'
            when i.PUBLISHING_GROUP = 'MUD' then 'MP'
            when i.PUBLISHING_GROUP = 'GAL-CL' then 'CL'
            when i.PUBLISHING_GROUP = 'GAL-BM' then 'BM'
            when i.PUBLISHER_CODE = 'Do Books' then 'DO'
            else i.PUBLISHING_GROUP
        end
        '''
        

def query_mtd(tp):
    '''
    Gives us daily sales and units shipped for the current month
    '''
    return f'''
    SELECT 
        sd.TRX_DATE,
        SUM(sd.QUANTITY_INVOICED) AS [Units],
        SUM(sd.REVENUE_AMOUNT) AS [Dollars]
    FROM
        ebs.Sales sd
        INNER JOIN ebs.Item t ON t.ITEM_ID = sd.ITEM_ID
        INNER JOIN ebs.Customer c ON c.SITE_USE_ID = sd.SHIP_TO_SITE_USE_ID
        INNER JOIN ebs.SalesRep sr ON sr.SALESREP_ID = sd.PRIMARY_SALESREP_ID
        LEFT OUTER JOIN ebs.Customer BillTo ON BillTo.SITE_USE_ID = c.BILL_TO_SITE_USE_ID
    WHERE 
        sd.PERIOD = '{tp}'
        AND sd.INVOICE_LINE_TYPE = 'SALE'
        AND cbq2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N'
        AND t.PUBLISHER_CODE = 'Chronicle'
        AND t.PRODUCT_TYPE IN ('BK', 'FT')
    GROUP BY 
        sd.TRX_DATE
    ORDER BY 
        sd.TRX_DATE DESC;
    '''
    
def query_ytd(typ1):
    '''
    Gives us daily sales and units shipped for the current month
    '''
    return f'''
    SELECT 
        sd.period [Period],
        SUM(sd.QUANTITY_INVOICED) AS [Units],
        SUM(sd.REVENUE_AMOUNT) AS [Dollars]
    FROM
        ebs.Sales sd
        INNER JOIN ebs.Item t ON t.ITEM_ID = sd.ITEM_ID
        INNER JOIN ebs.Customer c ON c.SITE_USE_ID = sd.SHIP_TO_SITE_USE_ID
        INNER JOIN ebs.SalesRep sr ON sr.SALESREP_ID = sd.PRIMARY_SALESREP_ID
        LEFT OUTER JOIN ebs.Customer BillTo ON BillTo.SITE_USE_ID = c.BILL_TO_SITE_USE_ID
    WHERE 
        sd.PERIOD >= '{typ1}'
        AND sd.INVOICE_LINE_TYPE = 'SALE'
        AND cbq2.dbo.fnSaleTypeCode(sd.AR_TRX_TYPE_ID) = 'N'
        AND t.PUBLISHER_CODE = 'Chronicle'
        AND t.PRODUCT_TYPE IN ('BK', 'FT')
    GROUP BY 
        sd.period
    ORDER BY 
        sd.period DESC;
    '''
    
def query_check_cbq_metrics(rows):
    '''
    Looks at the top N last updated tables in CBQ2
    '''
    return f'''
    SELECT top {rows} *
    FROM metrics.TableLastUpdated			
    Order by LastUpdated DESC
    '''

def query_viz_daily(ty,ly):
    '''
    the quiz for the altair viz
    '''
    return f'''
    SELECT
        i.PUBLISHER_CODE Publisher
        ,left(sd.period,4) [year]
        ,right(sd.period,2) [month]
        ,case
            when right(sd.period,2) in('01','02','03') then 'Q1'
            when right(sd.period,2) in('04','05','06') then 'Q2'
            when right(sd.period,2) in('07','08','09') then 'Q3'
            when right(sd.period,2) in('10','11','12') then 'Q4'
        end [quarter]
        ,case
            when left(period,4) in('{ty}','{ly}') then chan.Description
            else '-'
        end channel
        ,case
            when ssr_row.SSRRowID = 6 then 'Amaz'
            else 'ROM'
        end [Group]
        ,case                             
            when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'                      
            when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'                     
            when i.PUBLISHER_CODE IN('Glam Media','Benefit','PQ Blackwell','San Francisco Art Institute','AFO LLC','FareArts','Sager') then 'CD'                 
            when i.PUBLISHER_CODE = 'Creative Company' then 'CC'   
            when i.PUBLISHER_CODE = 'Do Books' then 'DO'
            when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
            when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'                                           
            when i.PUBLISHING_GROUP = 'GAL' then 'GA'                                                      
            when i.PUBLISHING_GROUP = 'GAL-CL' then 'CL'                        
            when i.PUBLISHING_GROUP = 'MUD' then 'MP'
            when i.PUBLISHING_GROUP = 'GAL-BM' then 'BM'             
            when i.PUBLISHING_GROUP in('LAU-BIS') then 'LKBS'                          
            when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE = 'FT' then 'LKGI'                      
            when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE <> 'FT' then 'LKBK'         
            when i.PUBLISHER_CODE = 'PRINCETON' and i.PUBLISHING_GROUP in('PAP-MS','PAP-MP') then 'PAMS'       
            when i.PUBLISHER_CODE = 'Princeton' and i.PRODUCT_TYPE = 'FT' then 'PAGI'                    
            when i.PUBLISHER_CODE = 'Princeton' and i.PRODUCT_TYPE <> 'FT' then 'PABK'                      
            when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
            when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'  
            when i.PUBLISHING_GROUP in('BAR-ART','BAR-ENT','BAR-LIF') then 'IMP'
            when i.PUBLISHING_GROUP in('CPB','CCB') then 'IMP'
            when i.PUBLISHING_GROUP in('RID','PTC','GAM') then 'RPG'
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'                         
            else i.PUBLISHING_GROUP                 
        end PubGroup
        ,sum(sd.REVENUE_AMOUNT) rev
        ,sum(sd.QUANTITY_INVOICED) qty
    FROM
        ebs.Sales sd
        inner join ssr.SalesSSRRow stie on stie.CUSTOMER_TRX_LINE_ID = sd.CUSTOMER_TRX_LINE_ID                  
        inner join ebs.Item i on sd.ITEM_ID = i.ITEM_ID                      
        inner join ssr.SSRRow ssr_row on ssr_row.SSRRowID= stie.SSRRowID
        inner join ssr.SubChannel sub on sub.SubChannelID = ssr_row.SubChannelID
        inner join ssr.Channel chan on chan.ChannelID = sub.ChannelID

    WHERE
        Sd.PERIOD >= '201901'
        AND sd.INVOICE_LINE_TYPE='SALE'
        AND cbq2.dbo.fnSaleTypeCode(SD.AR_TRX_TYPE_ID)='N'
        and i.PRODUCT_TYPE in('BK','FT')
        AND chan.ChannelID <> 4
        AND i.PUBLISHING_GROUP NOT IN('MKT','ZZZ')
        AND i.publisher_code not in('Benefit', 'Glam Media'
            ,'PQ Blackwell', 'Moleskine', 'AFO LLC','Sager','San Francisco Art Institute','FareArts')
    GROUP BY
        i.PUBLISHER_CODE
        ,left(sd.period,4)
        ,right(sd.period,2)
        ,case
            when right(sd.period,2) in('01','02','03') then 'Q1'
            when right(sd.period,2) in('04','05','06') then 'Q2'
            when right(sd.period,2) in('07','08','09') then 'Q3'
            when right(sd.period,2) in('10','11','12') then 'Q4'
        end
        ,case
            when left(period,4) in('{ty}','{ly}') then chan.Description
            else '-'
        end
        ,case
            when ssr_row.SSRRowID = 6 then 'Amaz'
            else 'ROM'
        end
        ,case                             
            when i.PUBLISHER_CODE = 'Tourbillon' then 'TW'                      
            when i.PUBLISHER_CODE = 'Sierra Club' then 'SC'                     
            when i.PUBLISHER_CODE IN('Glam Media','Benefit','PQ Blackwell','San Francisco Art Institute','AFO LLC','FareArts','Sager') then 'CD'                 
            when i.PUBLISHER_CODE = 'Creative Company' then 'CC'   
            when i.PUBLISHER_CODE = 'Do Books' then 'DO'
            when i.PUBLISHER_CODE = 'Levine Querido' then 'LQ'
            when i.PUBLISHER_CODE = 'AMMO Books' then 'AM'                                           
            when i.PUBLISHING_GROUP = 'GAL' then 'GA'                                                      
            when i.PUBLISHING_GROUP = 'GAL-CL' then 'CL'                        
            when i.PUBLISHING_GROUP = 'MUD' then 'MP'
            when i.PUBLISHING_GROUP = 'GAL-BM' then 'BM'             
            when i.PUBLISHING_GROUP in('LAU-BIS') then 'LKBS'                          
            when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE = 'FT' then 'LKGI'                      
            when i.PUBLISHER_CODE = 'Laurence King' and i.PRODUCT_TYPE <> 'FT' then 'LKBK'         
            when i.PUBLISHER_CODE = 'PRINCETON' and i.PUBLISHING_GROUP in('PAP-MS','PAP-MP') then 'PAMS'       
            when i.PUBLISHER_CODE = 'Princeton' and i.PRODUCT_TYPE = 'FT' then 'PAGI'                    
            when i.PUBLISHER_CODE = 'Princeton' and i.PRODUCT_TYPE <> 'FT' then 'PABK'                      
            when i.PUBLISHER_CODE = 'Hardie Grant' then 'HG'
	        when i.PUBLISHER_CODE = 'Quadrille Publishing Limited' then 'QD'  
            when i.PUBLISHING_GROUP in('BAR-ART','BAR-ENT','BAR-LIF') then 'IMP'
            when i.PUBLISHING_GROUP in('CPB','CCB') then 'IMP'
            when i.PUBLISHING_GROUP in('RID','PTC','GAM') then 'RPG'
            when i.PUBLISHING_GROUP in('FWN','LIF') then 'FLS'                         
            else i.PUBLISHING_GROUP                 
        end
    '''