import time

import pandas as pd
from loader_asin_mapping import load_asin_mapping
from loader_item import upload_item
from loader_monthly_reports import load_monthly_data
from UPDATE_ams_config import month_list, tab_dict

pd.set_option("future.no_silent_downcasting", True)


def main():
    pd.reset_option("display.max_columns")
    start_time = time.time()

    asin_mapping = load_asin_mapping()  # uses the default path
    item_df = upload_item()  # Load item data, if needed

    combined_df = pd.DataFrame()
    errors = []

    for month in month_list:
        if month not in tab_dict:
            warning = f"⚠️ Skipping {month}: not found in tab_dict"
            print(warning)
            errors.append(warning)
            continue

        try:
            print(f"🔄 Processing {month}...")
            df_month = load_monthly_data(tab_dict[month], asin_mapping, month)
            combined_df = pd.concat([combined_df, df_month], ignore_index=True)
        except Exception as e:
            error_msg = f"❌ Failed for {month}: {str(e)}"
            print(error_msg)
            errors.append(error_msg)

    # Merge with item_df on ISBN (after all months are processed)
    if not combined_df.empty and not item_df.empty:
        combined_df = pd.merge(combined_df, item_df, on="ISBN", how="left")

    # Save results
    if not combined_df.empty:
        # Excel output
        output_file_excel = "combined_amazon_ads_by_asin.xlsx"
        combined_df.to_excel(output_file_excel, index=False)
        print(f"✅ Combined data saved to Excel: {output_file_excel}")

        # Pickle output
        output_file = "combined_amazon_ads_by_asin.pkl"
        combined_df.to_pickle(output_file)
        print(f"✅ Combined data saved to pickle: {output_file}")
    else:
        print("❗No data was successfully combined.")

    # Save error log
    if errors:
        with open("processing_errors.log", "w") as f:
            for line in errors:
                f.write(line + "\n")
        print("⚠️ Some issues occurred. See processing_errors.log.")

    end_time = time.time()
    print(f"⏱️ Finished in {end_time - start_time:.2f} seconds.")


if __name__ == "__main__":
    main()
