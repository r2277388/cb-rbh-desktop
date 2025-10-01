import load_catalog
import load_inventory
import load_sales
import load_traffic

import pandas as pd
import time

def combine_weekly_files():
    sales_df = load_sales.df_sales()
    inventory_df = load_inventory.df_inventory()
    traffic_df = load_traffic.df_traffic()
    catalog_df = load_catalog.df_catalog()

    ## Below i'm trying to get a unique list of ASINS across all files 
    ## and a product title for each ASIN
    ## then merge each file onto that master list of ASINS

    # Step 1: Build ASIN to Product Title mapping
    asin_title_df = pd.concat([
        sales_df[['ASIN', 'Product Title']],
        inventory_df[['ASIN', 'Product Title']],
        traffic_df[['ASIN', 'Product Title']],
        catalog_df[['ASIN', 'Product Title']]
    ]).drop_duplicates(subset=['ASIN'])

    # Step 2: Build master ASIN list
    asin_master = pd.Series(
        pd.concat([
            sales_df['ASIN'],
            inventory_df['ASIN'],
            traffic_df['ASIN'],
            catalog_df['ASIN']
        ]).unique(),
        name='ASIN'
    ).to_frame()
    
    # Step 3: Add Product Title to asin_master (no duplicate warning needed)
    asin_master = asin_master.merge(asin_title_df, on='ASIN', how='left')

    # Merge DataFrames on 'SKU'
    # Merge all DataFrames onto master ASIN list
    combined_df = asin_master.merge(sales_df, on='ASIN', how='left', suffixes=('', '_sales'))
    combined_df = combined_df.merge(inventory_df, on='ASIN', how='left', suffixes=('', '_inventory'))
    combined_df = combined_df.merge(traffic_df, on='ASIN', how='left')
    combined_df = combined_df.merge(catalog_df, on='ASIN', how='left', suffixes=('', '_catalog'))

    # Drop all Product Title columns except the one from asin_master
    product_title_cols = [col for col in combined_df.columns if col.startswith('Product Title') and col != 'Product Title']
    combined_df = combined_df.drop(columns=product_title_cols)
    
    return combined_df

def main():
    start_time = time.time()  # Start timer
    df = combine_weekly_files()

    # Save combined_df as a pickle file
    df.to_pickle("combined_weekly_data.pkl")
    print("✅ Saved combined weekly data as pickle: combined_weekly_data.pkl")
    print()
    print(f"This dataframe has the following shape {df.shape}")
    print()
    print(df.info())
    print()
    print(df.head())
    print(df.columns.tolist())
    print()
    
    end_time = time.time()  # End timer
    elapsed = end_time - start_time
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    print(f"Total runtime: {minutes} minutes, {seconds} seconds.")
    
if __name__ == "__main__":
    main()