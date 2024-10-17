import pandas as pd
from pathlib import Path
import os
import glob

# Define the folder path where the traffic files are located
folder_path = Path(r'G:\SALES\Amazon\RBH\DOWNLOADED_FILES')

# Define the pattern to match the traffic files
file_glob_traffic = r'*Traffic*csv'

# Find all files matching the pattern
files = glob.glob(str(folder_path / file_glob_traffic))

# Select the most recent file based on creation time
file_traffic = max(files, key=os.path.getctime)

# Define the columns you want to load from the traffic file
cols_traffic = ['ASIN','Glance Views','Glance Views - Prior Period','Glance Views - Same Period Last Year']  # Replace with actual column names

def upload_traffic():
    df = pd.read_csv(file_traffic,
                             skiprows=1,  # Adjust if needed
                             na_values='—',
                             usecols=cols_traffic)
    
    df['Glance Views'] = (
        df['Glance Views']
        .replace(',','',regex=True)
        .fillna(0)
        .astype(int)
        )
    df['Glance Views - Prior Period'] = (
        df['Glance Views - Prior Period']
        .replace(r'[%,]','',regex=True)
        .fillna(0)
        .astype(float) / 100
        )
    df['Glance Views - Same Period Last Year'] = (
        df['Glance Views - Same Period Last Year']
        .replace(r'[%,]','',regex=True)
        .fillna(0)
        .astype(float) / 100
    )
    
    return df
    # Example: Cleaning and converting columns (customize based on your traffic 


def main():
    df = upload_traffic()
    print(df.info())
    print(df.head())
    
if __name__ == '__main__':
    main()