import pandas as pd

from paths import ATELIER_AMAZON_TRAFFIC_FOLDER


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
    files = list(ATELIER_AMAZON_TRAFFIC_FOLDER.glob("*Traffic*csv"))
    if not files:
        raise FileNotFoundError(
            f"No files found in {ATELIER_AMAZON_TRAFFIC_FOLDER} matching *Traffic*csv"
        )

    file_traffic = max(files, key=lambda path: path.stat().st_mtime)
    df = pd.read_csv(
        file_traffic,
        skiprows=1,
        na_values="Ã¢â‚¬â€",
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
