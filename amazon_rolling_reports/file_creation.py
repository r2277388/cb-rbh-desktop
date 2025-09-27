import pandas as pd
pd.set_option('future.no_silent_downcasting', True)

def create_rolling_report(pickle_file,pickle_po):
    df_co = pd.read_pickle(pickle_file)
    df_po = pd.read_pickle(pickle_po)
    df_combined = pd.merge(df_co, df_po, how='left', left_on='ISBN', right_on='ISBN')
    df_combined['PO_Qty'] = df_combined['PO_Qty'].fillna(0).astype(int)

    if 'PO_Qty' in df_combined.columns and 'AvgLast6W' in df_combined.columns:
        df_combined['AvgLast6W'] = pd.to_numeric(df_combined['AvgLast6W'], errors='coerce')
        divisor = df_combined['AvgLast6W'].replace(0, pd.NA)
        df_combined['OH_Avg'] = (df_combined['PO_Qty'] / divisor).round(2)
        df_combined['OH_Avg'] = df_combined['OH_Avg'].fillna(0).infer_objects(copy=False)
    else:
        print("PO_Qty or AvgLast6W column missing!")
        return df_combined

    # Reorder columns: place PO_Qty after PubDate and before OH, and OH_Avg after OH
    cols = list(df_combined.columns)
    if 'PO_Qty' in cols and 'PubDate' in cols and 'OH' in cols:
        cols.remove('PO_Qty')
        pubdate_index = cols.index('PubDate')
        cols.insert(pubdate_index + 1, 'PO_Qty')
    if 'OH_Avg' in cols and 'OH' in cols:
        cols.remove('OH_Avg')
        oh_index = cols.index('OH')
        cols.insert(oh_index + 1, 'OH_Avg')
    df_combined = df_combined[cols]
    return df_combined

def main():
    
    pickle_file1 = "rr_customer_orders.pkl"
    
    pickle_po = "latest_amazon_po.pkl"
    
    df_combined = create_rolling_report(pickle_file1,pickle_po)
    print(df_combined.shape)
    print(df_combined.columns[:20])
    print(df_combined.head())

if __name__ == "__main__":
    main()