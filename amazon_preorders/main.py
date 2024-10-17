from pathlib import Path
from data_processing import merge_catalog_inventory
from save_utils import save_to_excel

def main():
    # Merge catalog and inventory to create the DataFrame
    df = merge_catalog_inventory()  # Ensure this function does not require arguments or modify as needed

    # Define the folder path where the Excel files will be saved
    folder_path = Path(r'G:\SALES\Amazon\PREORDERS\2024')
    
    # Save the DataFrame to Excel
    save_to_excel(df, folder_path)
    
    # Print DataFrame information to confirm successful processing
    print(df.info())
    
if __name__ == '__main__':
    main()