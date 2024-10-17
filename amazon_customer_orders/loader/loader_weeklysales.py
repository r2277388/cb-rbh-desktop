import pandas as pd
from pathlib import Path
import os
import glob

def load_sales_data(folder_path: Path, file_glob_pattern: str, columns: list) -> pd.DataFrame:
    """
    Finds the most recent file matching the glob pattern in the specified folder
    and loads the data into a DataFrame.
    
    :param folder_path: The path to the folder containing the files.
    :param file_glob_pattern: The glob pattern to match files.
    :param columns: The list of columns to load from the file.
    :return: A pandas DataFrame with the loaded sales data.
    """
    # Find the most recent file matching the pattern
    files = glob.glob(str(folder_path) + file_glob_pattern)
    most_recent_file = max(files, key=os.path.getctime)
    
    # Load the sales data into a DataFrame
    df = pd.read_csv(most_recent_file, skiprows=1, na_values='—', usecols=columns)
    
    return df

def clean_and_convert_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans and converts the necessary columns in the sales DataFrame.
    
    :param df: The DataFrame containing the sales data.
    :return: The cleaned and converted DataFrame.
    """
    df['Ordered Revenue'] = df['Ordered Revenue'].replace(r'[$,]', '', regex=True).fillna(0).astype(float)
    df['Ordered Revenue - Prior Period'] = df['Ordered Revenue - Prior Period'].replace(r'[%,]', '', regex=True).fillna(0).astype(float)/100
    df['Ordered Revenue - Same Period Last Year'] = df['Ordered Revenue - Same Period Last Year'].replace(r'[%,]', '', regex=True).fillna(0).astype(float)/100
    df['Ordered Units'] = df['Ordered Units'].replace(',', '', regex=True).fillna(0).astype(int)
    df['Ordered Units - Prior Period'] = df['Ordered Units - Prior Period'].replace(r'[%,]', '', regex=True).fillna(0).astype(float)/100
    df['Ordered Units - Same Period Last Year'] = df['Ordered Units - Same Period Last Year'].replace(r'[%,]', '', regex=True).fillna(0).astype(float)/100
    
    return df

def uploader_weeklysales() -> pd.DataFrame:
    """
    Loads, cleans, and returns the weekly sales DataFrame.
    
    :return: The cleaned and loaded sales DataFrame.
    """
    # Define the folder path and file pattern
    folder_path = Path(r'G:\SALES\Amazon\RBH\DOWNLOADED_FILES')
    file_glob_sales_weekly = r'\*Sales*Weekly*csv'

    # Define the columns to load
    cols_sales_weekly = [
        'ASIN', 'Ordered Revenue', 'Ordered Revenue - Prior Period','Ordered Revenue - Same Period Last Year',
        'Ordered Units', 'Ordered Units - Prior Period','Ordered Units - Same Period Last Year'
    ]

    # Load the most recent sales data into a DataFrame
    df_sales_weekly = load_sales_data(folder_path, file_glob_sales_weekly, cols_sales_weekly)

    # Clean and convert the necessary columns
    df_sales_weekly = clean_and_convert_columns(df_sales_weekly)
    
    return df_sales_weekly

def main():
    df = uploader_weeklysales()  # Get the cleaned DataFrame
    print(df.info())
    print(df.head())

if __name__ == '__main__':
    main()
