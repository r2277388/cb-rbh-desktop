from paths import ypticod
import pandas as pd

def load_ypticod():
    df = pd.read_excel(ypticod, usecols=['ISBN', 'ISBN10','Availability Status','Publisher Name'])
    # Clean whitespace from ISBN10
    df['ISBN10'] = df['ISBN10'].astype(str).str.strip().str.zfill(10)
    # Ensure ISBN is string and zero-padded to 13 characters
    df['ISBN'] = df['ISBN'].astype(str).str.zfill(13)
    
    # Removing certain Publishers from this list
    publisher_delete_list = ['Princeton Architectural Press', 'AFO LLC'
                             , 'Benefit', 'Driscolls', 'FareArts', 'Moleskine'
                             , 'No Publisher Name', 'PQ Blackwell', 'Sager'
                             , 'San Francisco Art Institute','Glam Media']

    df = df[~df['Publisher Name'].isin(publisher_delete_list)]
    
    # Remove certain Availability Status's causing duplicates
    # avail_status_delete_list = ['OSI', 'PC']
    # df = df[~df['Availability Status'].isin(avail_status_delete_list)]

    # Rename ISBN10 to ASIN for clarity
    rename_dict = {'ISBN10': 'ASIN'}
    df = df.rename(columns=rename_dict)
    
    # Drop rows with missing ISBN or ASIN
    df = df.dropna(subset=['ISBN', 'ASIN']) 
    
    # Remove rows where ASIN is empty or 'nan' (case insensitive)
    df = df[df['ASIN'].notna() & (df['ASIN'] != '') & (df['ASIN'].str.lower() != 'nan')]
    
    # Remove all rows where ASIN is duplicated (keep only unique ASINs)
    # There are lots of duplicated ISBN10 in the ypticod file e.g. 6164302911
    df = df[~df['ASIN'].duplicated(keep=False)]
    df = df[~df['ISBN'].duplicated(keep=False)]
    
    
    # Only keep ASIN and ISBN columns
    df = df[['ASIN', 'ISBN']]
    
    
    print(f"YPTICOD rows: {len(df)}")
    print(df.head(3))
    
    # Check for duplicated ISBNs
    dup_isbn = df[df['ISBN'].duplicated(keep=False)]
    if not dup_isbn.empty:
        print(f"⚠️ Found {len(dup_isbn)} duplicated ISBNs:")
        print(dup_isbn[['ISBN', 'ASIN']])

    # Check for duplicated ASINs
    dup_asin = df[df['ASIN'].duplicated(keep=False)]
    if not dup_asin.empty:
        print(f"⚠️ Found {len(dup_asin)} duplicated ASINs:")
        print(dup_asin[['ISBN', 'ASIN']])
    return df

def main():
    return load_ypticod()

if __name__ == "__main__":
    main()