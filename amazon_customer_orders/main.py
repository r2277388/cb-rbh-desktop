import pandas as pd
from pathlib import Path
from tab_stats_cb import stats_cb
from tab_glance_cb import glance_views
from tab_TEMPLATE import top_titles

from loader.loader_weeklysales import uploader_weeklysales
from loader.loader_traffic import upload_traffic
from loader.loader_item import upload_item
from asin_isbn_converter import asin_isbn_conversion

from tqdm import tqdm
import time

def main():
    start_time = time.time() # Start time of the script
    
    # Loading shared dataframes
    df_weeklysales = uploader_weeklysales()
    df_converter = asin_isbn_conversion()
    df_item = upload_item()
    df_glance = upload_traffic()
    
    # Reports using reloaded data
    df_stats_cb = stats_cb()
    df_glance_cb = glance_views()

    publishers = [
        ("Chronicle", "FL", 20, 'top20_fl_cb'),
        ("!Chronicle", "FL", 20, 'top20_fl_dp'),
        ("Chronicle", "BL", 30, 'top30_cb_bl'),
        ("Sierra Club", None, 5, 'Sierra Club'),
        ("Galison", None, 5, 'Galison'),
        ("Laurence King", None, 5, 'Laurence_King'),
        ("Hardie Grant Publishing", None, 5, 'Hardie Grant Publishing'),
        ("Quadrille", None, 5, 'Quadrille'),
        ("Levine Querido", None, 5, 'Levine Querido'),
        ("Tourbillon", None, 5, 'Tourbillon'),
        ("Do Books", None, 5, 'Do Books'),
        ("Creative Company", None, 5, 'Creative Company'),
        ("Paperblanks", None, 5, 'Paperblanks')
        ]

    dfs = {}  # Dictionary to store all dataframes
    
        # Process each publisher and display progress with tqdm
    for publisher, flbl, num_rows, sheet_name in publishers:
        try:
            df = top_titles(df_weeklysales, df_converter, df_item, df_glance, publisher=publisher, flbl=flbl, num_rows=num_rows)
            dfs[sheet_name] = df
        except Exception as e:
            print(f"Error processing {sheet_name}: {e}")

    path = Path(fr'G:\SALES\Amazon\RBH\weekly_customer_order\atelier\amazon_weekly_customer_order_py.xlsx')

    try:
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            # Write each dataframe to a separate sheet
            df_stats_cb.to_excel(writer, sheet_name='stats_cb', index=False)
            df_glance_cb.to_excel(writer, sheet_name='glance_cb', index=False)
            for sheet_name, df in dfs.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        print(f"Error writing to Excel: {e}")
        
    end_time = time.time()  # End timing
    print(f"Total execution time: {end_time - start_time:.2f} seconds")
    
if __name__ == '__main__':
    main()