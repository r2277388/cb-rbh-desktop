from pathlib import Path
import pandas as pd
from datetime import date

def save_to_excel(df: pd.DataFrame, folder: Path) -> None:
    """Save the DataFrame to Excel files with the date in the format 'yyyy_mm_dd'."""
    # Get the current date in the desired format
    current_date = date.today().strftime('%Y_%m_%d')
    
    # Create the filenames with the date in the specified format
    dated_file_name = f'preorders_{current_date}.xlsx'
    current_file_name = 'current_amaz_preorders.xlsx'

    # Create the full paths
    dated_path = folder / dated_file_name
    current_path = folder / current_file_name

    # Write the DataFrame to the Excel files
    with pd.ExcelWriter(dated_path, engine='xlsxwriter') as dated_writer:
        df.to_excel(dated_writer, sheet_name='nyp', index=False)
        df.loc[df['publisher'] == 'Chronicle'].to_excel(dated_writer, sheet_name='nyp_cb', index=False)
        df.loc[df['publisher'] != 'Chronicle'].to_excel(dated_writer, sheet_name='nyp_dp', index=False)

    with pd.ExcelWriter(current_path, engine='xlsxwriter') as current_writer:
        df.to_excel(current_writer, sheet_name='nyp', index=False)
        df.loc[df['publisher'] == 'Chronicle'].to_excel(current_writer, sheet_name='nyp_cb', index=False)
        df.loc[df['publisher'] != 'Chronicle'].to_excel(current_writer, sheet_name='nyp_dp', index=False)
