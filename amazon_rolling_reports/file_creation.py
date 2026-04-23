import pandas as pd
pd.set_option('future.no_silent_downcasting', True)

from paths import amazon_po_pickle_file, customer_orders_pickle_file

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

    if 'TYTD' in df_combined.columns and 'LYTD' in df_combined.columns:
        tytd = pd.to_numeric(df_combined['TYTD'], errors='coerce').fillna(0)
        lytd = pd.to_numeric(df_combined['LYTD'], errors='coerce').fillna(0)
        df_combined['YTD Var'] = tytd - lytd

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
    if 'YTD Var' in cols and 'LYTD' in cols:
        cols.remove('YTD Var')
        lytd_index = cols.index('LYTD')
        cols.insert(lytd_index + 1, 'YTD Var')
    df_combined = df_combined[cols]
    if 'AvgLast6W' in df_combined.columns:
        df_combined = df_combined.rename(columns={'AvgLast6W': '6Wk Avg'})
    return df_combined

def main():
    pickle_file1 = customer_orders_pickle_file
    pickle_po = amazon_po_pickle_file
    
    df_combined = create_rolling_report(pickle_file1,pickle_po)
    print(df_combined.shape)
    print(df_combined.columns[:20])
    print(df_combined.head())

if __name__ == "__main__":
    main()
