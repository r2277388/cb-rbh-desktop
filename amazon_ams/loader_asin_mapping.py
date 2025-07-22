import pandas as pd

default_file_path = fr"G:\SALES\Amazon\RBH\DOWNLOADED_FILES\Chronicle-AsinMapping.xlsx"

def load_asin_mapping(file_path=default_file_path):
    df = pd.read_excel(
        file_path,
        usecols=['Asin', 'Isbn13'],
        sheet_name='Sheet1',
        header=0,
        engine='openpyxl'
    )
    df.columns = df.columns.str.lower()
    df.rename(columns={'asin': 'ASIN', 'isbn13': 'ISBN'}, inplace=True)
    df['ASIN'] = df['ASIN'].astype(str).str.zfill(10)
    return df

if __name__ == '__main__':
    asin_mapping = load_asin_mapping()
    print(asin_mapping.info())
    print(asin_mapping.head())
    # Save to a pickle file if needed