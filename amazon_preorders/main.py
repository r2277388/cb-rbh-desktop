from pathlib import Path

from data_processing import merge_catalog_inventory
from paths import AMAZON_PREORDERS_OUTPUT_FOLDER
from save_utils import save_to_excel


def main():
    # Merge catalog and inventory to create the DataFrame
    df = (
        merge_catalog_inventory()
    )  # Ensure this function does not require arguments or modify as needed

    # Define the folder path where the Excel files will be saved
    folder_path = AMAZON_PREORDERS_OUTPUT_FOLDER

    # Save the DataFrame to Excel
    dated_path, current_path = save_to_excel(df, folder_path)

    print()
    print("Saved Excel files:")
    print(f"  {dated_path}")
    print(f"  {current_path}")
    print()

    # Print DataFrame information to confirm successful processing
    df.info()
    print()


if __name__ == "__main__":
    main()
