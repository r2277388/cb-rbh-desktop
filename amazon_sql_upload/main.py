import pandas as pd

from asin_isbn_conversion import asin_isbn_conversion

def main():
    df = asin_isbn_conversion()
    
    # Select and reorder columns: ASIN, ISBN, then all float columns
    float_cols = df.select_dtypes(include='float64').columns.tolist()
    df = df[['ASIN', 'ISBN'] + float_cols]
    
    rename = {
        'ISBN': 'External ID',
        'Ordered Units': 'Customer Orders',
        'Shipped Units': 'Units Shipped',
        'Sellable On Hand Units': 'Units at Amazon',
        'Open Purchase Order Quantity': 'Open PO qty'
    }
    
    df = df.rename(columns=rename)

        # Save DataFrame to Excel
    # Save DataFrame to Excel starting at row 4 (Excel row 5)
    with pd.ExcelWriter("amazon_sql_upload.xlsx") as writer:
        df.to_excel(writer, index=False, startrow=3)
    print("✅ Saved Excel file: amazon_sql_upload.xlsx")
    
    print(df.head())
    print(df.info())

if __name__ == "__main__":
    main()