import glob
import os

import pandas as pd

from paths import DOWNLOADS_FOLDER


folder_path = DOWNLOADS_FOLDER
file_glob_traffic = r"*Traffic*csv"

column_mapping = {
    "ASIN": "ASIN",
    "Featured Offer Page Views": "Glance Views",
    "Featured Offer Page Views - Prior Period (%)": "Glance Views - Prior Period",
    "Featured Offer Page Views - Same Period Last Year (%)": "Glance Views - Same Period Last Year",
}

cols_traffic = [
    "ASIN",
    "Featured Offer Page Views",
    "Featured Offer Page Views - Prior Period (%)",
    "Featured Offer Page Views - Same Period Last Year (%)",
]


def upload_traffic():
    files = glob.glob(str(folder_path / file_glob_traffic))
    if not files:
        raise FileNotFoundError(f"No files found in {folder_path} matching {file_glob_traffic}")

    file_traffic = max(files, key=os.path.getctime)
    df = pd.read_csv(
        file_traffic,
        skiprows=1,
        na_values="â€”",
        usecols=cols_traffic,
    )

    df = df.rename(columns=column_mapping)
    df["Glance Views"] = df["Glance Views"].replace(",", "", regex=True).fillna(0).astype(int)
    df["Glance Views - Prior Period"] = (
        df["Glance Views - Prior Period"].replace(r"[%,]", "", regex=True).fillna(0).astype(float) / 100
    )
    df["Glance Views - Same Period Last Year"] = (
        df["Glance Views - Same Period Last Year"]
        .replace(r"[%,]", "", regex=True)
        .fillna(0)
        .astype(float)
        / 100
    )
    return df


def main():
    df = upload_traffic()
    print(df.info())
    print(df.head())


if __name__ == "__main__":
    main()
