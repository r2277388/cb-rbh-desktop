import pandas as pd
from pathlib import Path
from tkinter import Tk, filedialog

from combined_file import get_merged_files
import excel_report

def publisher_summary(df):
    df = df.groupby('Publisher').agg({'Requested quantity': 'sum',
                                        'Accepted Quantity': 'sum', 
                                        'Total accepted cost': 'sum',
                                        'Lost Sales': 'sum',
                                        'Requested Cost': 'sum',
                                       })
    df.sort_values(by='Total accepted cost', ascending=False, inplace=True)
    
    return df

def ordered_summary_cb(df):
    df_summary = df[df['Publisher'] == 'Chronicle'].copy()
    df_summary.sort_values(by='Requested quantity', ascending=False, inplace=True)
    return df_summary.head(20)

def ordered_summary_dp(df):
    df_summary = df[df['Publisher'] != 'Chronicle'].copy()
    df_summary.sort_values(by='Requested quantity', ascending=False, inplace=True)
    return df_summary.head(20)

def accepted_summary_cb(df):
    df_summary = df[df['Publisher'] == 'Chronicle'].copy()
    df_summary.sort_values(by='Total accepted cost', ascending=False, inplace=True)
    return df_summary.head(20)

def accepted_summary_dp(df):
    df_summary = df[df['Publisher'] != 'Chronicle'].copy()
    df_summary.sort_values(by='Total accepted cost', ascending=False, inplace=True)
    return df_summary.head(20)

def lost_sales_summary(df):
    df_summary = df[df['Publisher'] != 'Chronicle'].copy()
    df_summary.sort_values(by='Lost Sales', ascending=False, inplace=True)
    return df_summary.head(20)

def main():
    
    excel_report.main()
    
    df = get_merged_files()
    
    filename = Path(rf'G:\SALES\Amazon\PURCHASE ORDERS\atelier\po_analysis\amazon_order_py_dump.xlsx')
    with pd.ExcelWriter(filename) as writer:
        publisher_summary(df).to_excel(writer, sheet_name='pub_summary')
        ordered_summary_cb(df).to_excel(writer, sheet_name='ordered_summary_cb',index=False)
        ordered_summary_dp(df).to_excel(writer, sheet_name='ordered_summary_dp',index=False)
        accepted_summary_cb(df).to_excel(writer, sheet_name='accepted_summary_cb',index=False)
        accepted_summary_dp(df).to_excel(writer, sheet_name='accepted_summary_dp',index=False)
        lost_sales_summary(df).to_excel(writer, sheet_name='lost_sales_summary',index=False)

if __name__ == "__main__":
    main()