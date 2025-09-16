import pandas as pd
import os

# Set your folder path containing the invoices
folder_path = fr"C:\Users\rbh\code\power_query_tests"

# List all CSV files in the folder
csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]

all_dataframes = []

for file in csv_files:
    file_path = os.path.join(folder_path, file)
    try:
        # Read and clean file
        df = pd.read_csv(file_path)
        df.columns = df.columns.str.strip()

        # Rename 'Unnamed: 2' to 'grouping'
        df = df.rename(columns={'Unnamed: 2': 'grouping'})

        # Drop summary columns if present
        drop_cols = ['Total Units', 'Gross Sales', 'Net Sales', 'Owing']
        df = df.drop(columns=[col for col in drop_cols if col in df.columns])

        # Identify date columns dynamically
        static_cols = ['isbn', 'title', 'grouping', 'Unit Price']
        date_columns = [col for col in df.columns if col not in static_cols]

        # Melt into long format
        df_melted = df.melt(
            id_vars=static_cols,
            value_vars=date_columns,
            var_name='Date',
            value_name='Value'
        )

        # Clean data types
        df_melted['Date'] = pd.to_datetime(df_melted['Date'], errors='coerce')
        df_melted['Value'] = pd.to_numeric(df_melted['Value'], errors='coerce')

        # Pivot groupings into wide format
        df_clean = df_melted.pivot_table(
            index=['isbn', 'title', 'Unit Price', 'Date'],
            columns='grouping',
            values='Value',
            aggfunc='sum'
        ).reset_index()

        df_clean.columns.name = None
        all_dataframes.append(df_clean)

    except Exception as e:
        print(f"Skipping file {file} due to error: {e}")

# Combine all valid dataframes
if all_dataframes:
    final_df = pd.concat(all_dataframes, ignore_index=True)
else:
    final_df = pd.DataFrame()