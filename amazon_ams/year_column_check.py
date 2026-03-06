import numpy as np
import pandas as pd
from UPDATE_ams_config import tab_dict, month_list

pd.reset_option("display.max_columns")


file_asin_mapping = r"G:\SALES\Amazon\RBH\DOWNLOADED_FILES\Chronicle-AsinMapping.xlsx"

df_asin_mapping = pd.read_excel(
    file_asin_mapping,
    usecols=["Asin", "Isbn13"],  # specify the columns to read
    sheet_name="Sheet1",
    header=0,  # use the next row as header
    engine="openpyxl",
)

# Lowercase all column names
df_asin_mapping.columns = df_asin_mapping.columns.str.lower()
# Rename to standard names
df_asin_mapping.rename(columns={"asin": "ASIN", "isbn13": "ISBN"}, inplace=True)

# Baseline year used to validate new files against known-good structure.
BASELINE_YEAR = "2025"

# Change this if you want to inspect a specific month.
month = month_list[-1]
if month not in tab_dict:
    raise KeyError(
        f"Selected month '{month}' is not in tab_dict. "
        f"Available months: {month_list}"
    )

df = pd.read_excel(
    tab_dict[month]["file"],
    sheet_name=tab_dict[month]["tab"],
    skiprows=tab_dict[month]["skiprows"],  # skip the first row
    header=0,  # use the next row as header
    engine="openpyxl",
)

df.columns = df.columns.str.strip().str.lower()
# Remove unwanted columns
df = df.drop(
    columns=[
        "isbn",
        "title",
        "pub",
        "pub group",
        "osd",
        "ctr",
        "cvr",
        "acos",
        "roas",
        "product type description",
    ],
    errors="ignore",
)
df.rename(columns={"asin": "ASIN"}, inplace=True)
df = df.merge(df_asin_mapping, on="ASIN", how="left")

df["ISBN"] = df["ISBN"].fillna(df["ISBN"].str.replace("-", "", regex=False))

df["CTR"] = df["clicks"].div(df["impressions"]).replace([np.inf, -np.inf], np.nan)
df["CRV"] = df["units sold"].div(df["clicks"]).replace([np.inf, -np.inf], np.nan)
df["ACOS"] = (
    df["spend"].div(df["14 day total sales"]).replace([np.inf, -np.inf], np.nan)
)
df["ROAS"] = (
    df["14 day total sales"].div(df["spend"]).replace([np.inf, -np.inf], np.nan)
)
df["period"] = month

#######################

column_sets = {}

for month in month_list:
    try:
        df = pd.read_excel(
            tab_dict[month]["file"],
            sheet_name=tab_dict[month]["tab"],
            skiprows=tab_dict[month]["skiprows"],
            header=0,
            engine="openpyxl",
        )
        df.columns = df.columns.str.strip().str.lower()
        column_sets[month] = set(df.columns)
    except Exception as e:
        print(f"ERROR for {month}: {e}")

# Print columns for each month
for month, cols in column_sets.items():
    print(f"{month}: {sorted(cols)}")

# Check if all months have the same columns
all_columns = list(column_sets.values())
if all(all_columns[0] == cols for cols in all_columns):
    print("OK: All months have the same columns.")
else:
    print("WARNING: Columns differ between months.")
    # Optionally, show which months are different
    from collections import Counter

    col_counter = Counter(tuple(sorted(cols)) for cols in all_columns)
    print("Unique column sets and their counts:", col_counter)

# Compare newer months to baseline year
baseline_months = [m for m in month_list if m.startswith(f"{BASELINE_YEAR}-")]
new_months = [m for m in month_list if m[:4] > BASELINE_YEAR]

print()
print(f"Baseline year: {BASELINE_YEAR}")
if not baseline_months:
    print(f"ERROR: No baseline months found for {BASELINE_YEAR}.")
else:
    baseline_sets = {m: column_sets[m] for m in baseline_months if m in column_sets}
    if not baseline_sets:
        print(f"ERROR: Baseline months for {BASELINE_YEAR} could not be loaded.")
    else:
        first_baseline_month = sorted(baseline_sets.keys())[0]
        baseline_columns = baseline_sets[first_baseline_month]
        baseline_consistent = all(cols == baseline_columns for cols in baseline_sets.values())

        print(f"Baseline months loaded: {len(baseline_sets)}")
        if baseline_consistent:
            print("Baseline columns are consistent across baseline months.")
        else:
            print("WARNING: Baseline year has column differences across months.")

        print()
        if not new_months:
            print("No newer months found beyond baseline year.")
        else:
            print("Comparison of newer months vs baseline columns:")
            for m in sorted(new_months):
                if m not in column_sets:
                    print(f"{m}: ERROR (could not load columns)")
                    continue
                new_cols = column_sets[m]
                missing = sorted(baseline_columns - new_cols)
                extra = sorted(new_cols - baseline_columns)
                if not missing and not extra:
                    print(f"{m}: MATCH")
                else:
                    print(f"{m}: DIFF")
                    print(f"  Missing vs baseline: {missing}")
                    print(f"  Extra vs baseline: {extra}")
