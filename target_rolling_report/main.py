from __future__ import annotations

import argparse
import re
import shutil
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

try:
    from paths import process_paths
    from shared.db import fetch_data_from_db, get_connection
except ImportError:  # Allows direct execution from this folder.
    import sys

    sys.path.append(str(Path(__file__).resolve().parents[1]))
    from paths import process_paths
    from shared.db import fetch_data_from_db, get_connection


LOCAL_DIR = Path(__file__).resolve().parent
SKELETON_FILE = LOCAL_DIR / "Target_Skeleton_Look.xlsx"
LOCAL_SAMPLE_SALES = LOCAL_DIR / "Target NOC Sales.csv"
LOCAL_SAMPLE_INVENTORY = LOCAL_DIR / "Target NOC Inventory.csv"

SALES_CACHE = process_paths.TARGET_NOC_CACHE_DIR / "target_noc_weekly_sales.parquet"
YEARLY_CACHE = process_paths.TARGET_NOC_CACHE_DIR / "target_noc_yearly_sales.parquet"
METADATA_CACHE = process_paths.TARGET_NOC_CACHE_DIR / "target_noc_metadata.parquet"
INVENTORY_CACHE = process_paths.TARGET_NOC_CACHE_DIR / "target_noc_inventory.parquet"

METADATA_COLUMNS = ["DPCI", "Pub", "PT", "FT", "PGRP", "ISBN", "Title", "Price", "PubDate"]
REPORT_MIN_POSITIVE_SALES_DATE = pd.Timestamp("2023-01-01")

ELIGIBLE_TITLE_SQL_TEMPLATE = """
SELECT
    i.item_title AS ISBN,
    i.PUBLISHER_CODE AS Pub,
    i.PRODUCT_TYPE AS PT,
    i.FORMAT AS FT,
    i.PUBLISHING_GROUP AS PGRP,
    i.SHORT_TITLE AS Title,
    i.PRICE_AMOUNT AS Price,
    CAST(i.AMORTIZATION_DATE AS date) AS PubDate
FROM ebs.Item i
WHERE
    i.ITEM_TITLE IN ({isbn_list})
    AND i.PUBLISHER_CODE NOT IN (
        'Benefit', 'AFO LLC', 'Glam Media', 'PQ Blackwell', 'PRINCETON', 'AMMO Books',
        'San Francisco Art Institute', 'FareArts', 'Sager', 'In Active', 'Driscolls',
        'Impossible Foods', 'Moleskine'
    )
    AND i.SHORT_TITLE IS NOT NULL
    AND i.ITEM_TITLE IS NOT NULL
    AND i.PRODUCT_TYPE IN ('BK', 'FT', 'CP', 'RP')
"""


def use_cache_dir(cache_dir: Path) -> None:
    global SALES_CACHE, YEARLY_CACHE, METADATA_CACHE, INVENTORY_CACHE
    SALES_CACHE = cache_dir / "target_noc_weekly_sales.parquet"
    YEARLY_CACHE = cache_dir / "target_noc_yearly_sales.parquet"
    METADATA_CACHE = cache_dir / "target_noc_metadata.parquet"
    INVENTORY_CACHE = cache_dir / "target_noc_inventory.parquet"


@dataclass(frozen=True)
class SourceFiles:
    sales: Path
    inventory: Path


def normalize_isbn(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    digits = re.sub(r"\D", "", str(value).strip())
    if not digits:
        return ""
    if len(digits) < 13:
        return digits.zfill(13)
    if len(digits) > 13 and digits.startswith("0"):
        return digits[-13:]
    return digits[:13]


def parse_number(value: object) -> float:
    if value is None or pd.isna(value):
        return 0.0
    text = str(value).replace(",", "").replace("$", "").strip()
    if text.endswith("%"):
        text = text[:-1]
        try:
            return float(text) / 100
        except ValueError:
            return 0.0
    if text in {"", "-"}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_date(value: object) -> pd.Timestamp | None:
    if value is None or pd.isna(value):
        return None
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return pd.Timestamp(parsed).normalize()


def quote_sql(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def saturday_range_ending(latest_week: pd.Timestamp, count: int) -> list[pd.Timestamp]:
    latest_week = pd.Timestamp(latest_week).normalize()
    return [latest_week - pd.Timedelta(days=7 * (count - 1 - idx)) for idx in range(count)]


def latest_file(folder: Path, patterns: tuple[str, ...]) -> Path | None:
    if not folder.exists():
        return None
    files: list[Path] = []
    for pattern in patterns:
        files.extend(
            path
            for path in folder.glob(pattern)
            if path.is_file() and not path.name.startswith("~$")
        )
    if not files:
        return None
    return max(files, key=lambda path: path.stat().st_mtime)


def resolve_source_files(use_samples: bool = False) -> SourceFiles:
    if use_samples:
        return SourceFiles(LOCAL_SAMPLE_SALES, LOCAL_SAMPLE_INVENTORY)

    sales = latest_file(process_paths.TARGET_NOC_SALES_FOLDER, ("*.csv", "*.xlsx", "*.xls"))
    inventory = latest_file(process_paths.TARGET_NOC_INVENTORY_FOLDER, ("*.csv", "*.xlsx", "*.xls"))

    return SourceFiles(
        sales=sales or LOCAL_SAMPLE_SALES,
        inventory=inventory or LOCAL_SAMPLE_INVENTORY,
    )


def read_two_row_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path, header=None, dtype=object)
    return pd.read_csv(path, header=None, dtype=object)


def bootstrap_from_skeleton(force: bool = False) -> None:
    if not force and SALES_CACHE.exists() and METADATA_CACHE.exists() and YEARLY_CACHE.exists():
        return
    if not SKELETON_FILE.exists():
        raise FileNotFoundError(f"Skeleton workbook not found: {SKELETON_FILE}")

    SALES_CACHE.parent.mkdir(parents=True, exist_ok=True)
    raw = pd.read_excel(SKELETON_FILE, sheet_name=0, header=None, dtype=object)

    header_row = raw.iloc[4]
    weeknum_row = raw.iloc[3]
    data = raw.iloc[5:].copy()
    data = data[data.iloc[:, 0].notna()]

    metadata = pd.DataFrame(
        {
            "DPCI": data.iloc[:, 0].astype(str).str.strip(),
            "Pub": data.iloc[:, 1],
            "PT": data.iloc[:, 2],
            "FT": data.iloc[:, 3],
            "PGRP": data.iloc[:, 4],
            "ISBN": data.iloc[:, 5].map(normalize_isbn),
            "Title": data.iloc[:, 6],
            "Price": data.iloc[:, 7].map(parse_number),
            "PubDate": pd.to_datetime(data.iloc[:, 8], errors="coerce"),
        }
    )
    metadata = metadata.drop_duplicates(subset=["DPCI", "ISBN"], keep="first")

    weekly_rows = []
    yearly_rows = []
    for col_idx in range(26, raw.shape[1]):
        label = header_row.iloc[col_idx]
        week = parse_date(label)
        if week is not None:
            weeknum = int(parse_number(weeknum_row.iloc[col_idx])) or week.isocalendar().week
            values = data.iloc[:, col_idx].map(parse_number)
            for row_idx, units in values.items():
                if units == 0:
                    continue
                weekly_rows.append(
                    {
                        "DPCI": str(data.loc[row_idx].iloc[0]).strip(),
                        "ISBN": normalize_isbn(data.loc[row_idx].iloc[5]),
                        "Week": week,
                        "WeekNum": weeknum,
                        "Units": units,
                    }
                )
            continue

        year_match = re.fullmatch(r"\d{4}", str(label).strip()) if label is not None else None
        if year_match:
            year = int(str(label).strip())
            values = data.iloc[:, col_idx].map(parse_number)
            for row_idx, units in values.items():
                if units == 0:
                    continue
                yearly_rows.append(
                    {
                        "DPCI": str(data.loc[row_idx].iloc[0]).strip(),
                        "ISBN": normalize_isbn(data.loc[row_idx].iloc[5]),
                        "Year": year,
                        "Units": units,
                    }
                )

    weekly = pd.DataFrame(weekly_rows, columns=["DPCI", "ISBN", "Week", "WeekNum", "Units"])
    yearly = pd.DataFrame(yearly_rows, columns=["DPCI", "ISBN", "Year", "Units"])

    metadata.to_parquet(METADATA_CACHE, index=False)
    weekly.to_parquet(SALES_CACHE, index=False)
    yearly.to_parquet(YEARLY_CACHE, index=False)
    print(f"Bootstrapped Target NOC cache from {SKELETON_FILE}")
    print(f"  Metadata rows: {len(metadata):,}")
    print(f"  Weekly sales rows: {len(weekly):,}")
    print(f"  Yearly sales rows: {len(yearly):,}")


def load_inventory_snapshot(path: Path) -> tuple[pd.Timestamp, pd.DataFrame]:
    raw = read_two_row_file(path)
    snapshot_date = None
    for value in raw.iloc[0].tolist():
        snapshot_date = parse_date(value)
        if snapshot_date is not None:
            break
    if snapshot_date is None:
        raise ValueError(f"No inventory snapshot date found in row 1 of {path}")

    headers = raw.iloc[1].fillna("").astype(str).str.strip().tolist()
    df = raw.iloc[2:].copy()
    df.columns = headers
    if "DPCI" not in df.columns:
        raise ValueError(f"Inventory file is missing DPCI: {path}")

    store_col = next((col for col in df.columns if col.lower() == "store count (active locs)"), None)
    on_hand_col = next((col for col in df.columns if col.lower() == "eoh u"), None)
    instock_col = next((col for col in df.columns if col.lower() == "instock %"), None)
    missing = [
        name
        for name, col in {
            "Store Count (Active Locs)": store_col,
            "EOH U": on_hand_col,
            "Instock %": instock_col,
        }.items()
        if col is None
    ]
    if missing:
        raise ValueError(f"Inventory file is missing columns: {', '.join(missing)}")

    snapshot = pd.DataFrame(
        {
            "DPCI": df["DPCI"].astype(str).str.strip(),
            "Week": snapshot_date,
            "Store Ct": df[store_col].map(parse_number),
            "On Hand": df[on_hand_col].map(parse_number),
            "INSTK %": df[instock_col].map(parse_number),
        }
    )
    snapshot = snapshot[snapshot["DPCI"].ne("")]
    return snapshot_date, snapshot


def load_sales_file(path: Path, latest_week: pd.Timestamp) -> tuple[pd.Timestamp, pd.DataFrame, pd.DataFrame]:
    raw = read_two_row_file(path)
    headers = raw.iloc[1].fillna("").astype(str).str.strip().tolist()
    data = raw.iloc[2:].copy()
    data.columns = headers

    if "Barcode (UPC)" not in data.columns or "DPCI" not in data.columns:
        raise ValueError(f"Sales file must include Barcode (UPC) and DPCI columns: {path}")

    week_columns = [idx for idx, header in enumerate(headers) if header.lower() == "sales u"]
    if not week_columns:
        raise ValueError(f"Sales file has no Sales U week columns: {path}")

    week_dates = saturday_range_ending(latest_week, len(week_columns))
    rows = []
    metadata_rows = []
    for _, row in data.iterrows():
        dpci = str(row.get("DPCI", "")).strip()
        isbn = normalize_isbn(row.get("Barcode (UPC)", ""))
        if not dpci or not isbn:
            continue
        metadata_rows.append({"DPCI": dpci, "ISBN": isbn})
        for col_idx, week in zip(week_columns, week_dates):
            units = parse_number(row.iloc[col_idx])
            rows.append({"DPCI": dpci, "ISBN": isbn, "Week": week, "Units": units})

    sales = pd.DataFrame(rows, columns=["DPCI", "ISBN", "Week", "Units"])
    if sales.empty:
        raise ValueError(f"No sales rows were loaded from {path}")
    sales = sales.groupby(["DPCI", "ISBN", "Week"], as_index=False)["Units"].sum()
    metadata = pd.DataFrame(metadata_rows).drop_duplicates()
    return max(week_dates), sales, metadata


def assign_week_numbers(new_sales: pd.DataFrame, existing: pd.DataFrame) -> pd.DataFrame:
    known = existing[["Week", "WeekNum"]].drop_duplicates() if not existing.empty else pd.DataFrame()
    sales = new_sales.merge(known, on="Week", how="left")
    missing_weeks = sorted(sales.loc[sales["WeekNum"].isna(), "Week"].drop_duplicates())
    if not missing_weeks:
        sales["WeekNum"] = sales["WeekNum"].astype(int)
        return sales

    known_by_year = known.copy()
    if not known_by_year.empty:
        known_by_year["Year"] = pd.to_datetime(known_by_year["Week"]).dt.year

    for week in missing_weeks:
        prior = known_by_year[
            (known_by_year["Year"] == week.year) & (pd.to_datetime(known_by_year["Week"]) < week)
        ]
        if not prior.empty:
            weeknum = int(prior.sort_values("Week").iloc[-1]["WeekNum"]) + 1
        else:
            weeknum = week.isocalendar().week
        sales.loc[sales["Week"].eq(week), "WeekNum"] = weeknum
        known_by_year = pd.concat(
            [known_by_year, pd.DataFrame([{"Week": week, "WeekNum": weeknum, "Year": week.year}])],
            ignore_index=True,
        )
    sales["WeekNum"] = sales["WeekNum"].astype(int)
    return sales


def apply_sales_dpci_mapping_to_cache(sales_mapping: pd.DataFrame) -> None:
    if sales_mapping.empty:
        return
    mapping = sales_mapping[["ISBN", "DPCI"]].dropna().drop_duplicates(subset=["ISBN"], keep="last")
    mapping["ISBN"] = mapping["ISBN"].map(normalize_isbn)
    mapping["DPCI"] = mapping["DPCI"].astype(str).str.strip()
    mapping_dict = dict(zip(mapping["ISBN"], mapping["DPCI"]))
    if not mapping_dict:
        return

    if SALES_CACHE.exists():
        weekly = pd.read_parquet(SALES_CACHE)
        weekly["ISBN"] = weekly["ISBN"].map(normalize_isbn)
        weekly["DPCI"] = weekly.apply(
            lambda row: mapping_dict.get(row["ISBN"], row["DPCI"]), axis=1
        )
        weekly["Week"] = pd.to_datetime(weekly["Week"]).dt.normalize()
        weekly = (
            weekly.groupby(["DPCI", "ISBN", "Week"], as_index=False)
            .agg({"WeekNum": "max", "Units": "sum"})
            .sort_values(["DPCI", "ISBN", "Week"])
        )
        weekly.to_parquet(SALES_CACHE, index=False)

    if METADATA_CACHE.exists():
        metadata = pd.read_parquet(METADATA_CACHE)
        metadata["ISBN"] = metadata["ISBN"].map(normalize_isbn)
        metadata["DPCI"] = metadata.apply(
            lambda row: mapping_dict.get(row["ISBN"], row["DPCI"]), axis=1
        )
        metadata = metadata.drop_duplicates(subset=["DPCI", "ISBN"], keep="first")
        metadata.to_parquet(METADATA_CACHE, index=False)


def fetch_eligible_title_metadata(isbns: pd.Series | list[str]) -> pd.DataFrame:
    clean_isbns = sorted({normalize_isbn(isbn) for isbn in isbns if normalize_isbn(isbn)})
    if not clean_isbns:
        return pd.DataFrame(columns=METADATA_COLUMNS)

    engine = get_connection()
    frames = []
    chunk_size = 900
    for start in range(0, len(clean_isbns), chunk_size):
        chunk = clean_isbns[start : start + chunk_size]
        query = ELIGIBLE_TITLE_SQL_TEMPLATE.format(
            isbn_list=", ".join(quote_sql(isbn) for isbn in chunk)
        )
        frames.append(fetch_data_from_db(engine, query))

    metadata = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if metadata.empty:
        return pd.DataFrame(columns=METADATA_COLUMNS)

    metadata["ISBN"] = metadata["ISBN"].map(normalize_isbn)
    metadata["Price"] = metadata["Price"].map(parse_number)
    metadata["PubDate"] = pd.to_datetime(metadata["PubDate"], errors="coerce")
    metadata = metadata.drop_duplicates(subset=["ISBN"], keep="first")
    return metadata[["Pub", "PT", "FT", "PGRP", "ISBN", "Title", "Price", "PubDate"]]


def sync_eligible_metadata_from_sql(
    isbns: pd.Series | list[str] | None = None,
    required: bool = True,
) -> pd.DataFrame:
    bootstrap_from_skeleton()
    if not SALES_CACHE.exists():
        return pd.DataFrame(columns=METADATA_COLUMNS)

    weekly = pd.read_parquet(SALES_CACHE)
    if weekly.empty:
        return pd.DataFrame(columns=METADATA_COLUMNS)

    target_isbns = weekly["ISBN"] if isbns is None else isbns
    eligible = fetch_eligible_title_metadata(target_isbns)
    if eligible.empty:
        message = "The eligible Target title SQL returned no matching ISBNs."
        if required:
            raise RuntimeError(message)
        print(f"Warning: {message}")
        return eligible

    dpci_mapping = (
        weekly[["ISBN", "DPCI"]]
        .dropna()
        .assign(ISBN=lambda df: df["ISBN"].map(normalize_isbn), DPCI=lambda df: df["DPCI"].astype(str).str.strip())
        .drop_duplicates(subset=["ISBN"], keep="last")
    )
    metadata = eligible.merge(dpci_mapping, on="ISBN", how="inner")
    metadata = metadata[METADATA_COLUMNS].drop_duplicates(subset=["DPCI", "ISBN"], keep="first")
    metadata.to_parquet(METADATA_CACHE, index=False)

    eligible_isbns = set(metadata["ISBN"])
    weekly["ISBN"] = weekly["ISBN"].map(normalize_isbn)
    filtered_weekly = weekly[weekly["ISBN"].isin(eligible_isbns)].copy()
    removed = len(weekly) - len(filtered_weekly)
    if removed:
        filtered_weekly.to_parquet(SALES_CACHE, index=False)
        print(f"Removed {removed:,} cached weekly rows for Target ISBNs not in the eligible SQL title list.")

    if YEARLY_CACHE.exists():
        yearly = pd.read_parquet(YEARLY_CACHE)
        if not yearly.empty:
            yearly["ISBN"] = yearly["ISBN"].map(normalize_isbn)
            filtered_yearly = yearly[yearly["ISBN"].isin(eligible_isbns)].copy()
            yearly_removed = len(yearly) - len(filtered_yearly)
            if yearly_removed:
                filtered_yearly.to_parquet(YEARLY_CACHE, index=False)
                print(f"Removed {yearly_removed:,} cached yearly rows for Target ISBNs not in the eligible SQL title list.")

    print(f"Synced {len(metadata):,} eligible Target title metadata rows from SQL.")
    return metadata


def refresh_cache(use_samples: bool = False, assume_yes: bool = False) -> pd.Timestamp:
    bootstrap_from_skeleton()
    sources = resolve_source_files(use_samples=use_samples)
    if not sources.sales.exists():
        raise FileNotFoundError(f"Sales source file not found: {sources.sales}")
    if not sources.inventory.exists():
        raise FileNotFoundError(f"Inventory source file not found: {sources.inventory}")

    inventory_week, inventory = load_inventory_snapshot(sources.inventory)
    weekly = pd.read_parquet(SALES_CACHE)
    weekly["Week"] = pd.to_datetime(weekly["Week"]).dt.normalize()
    latest_cached = weekly["Week"].max() if not weekly.empty else None

    print("\nTarget NOC source files:")
    print(f"  Sales:     {sources.sales}")
    print(f"  Inventory: {sources.inventory}")
    if latest_cached is not None:
        print(f"  Last cached week: {latest_cached:%B %d, %Y}")
    print(f"  Incoming week ending: {inventory_week:%B %d, %Y}")

    if not assume_yes:
        answer = input("Continue with this week ending date? [Y/n]: ").strip().lower()
        if answer in {"n", "no"}:
            override = input("Enter the Saturday week ending date (YYYY-MM-DD): ").strip()
            parsed = parse_date(override)
            if parsed is None:
                raise ValueError(f"Invalid week ending date: {override}")
            inventory_week = parsed
            inventory["Week"] = inventory_week

    latest_week, sales, new_metadata = load_sales_file(sources.sales, inventory_week)
    if latest_week != inventory_week:
        inventory_week = latest_week
        inventory["Week"] = latest_week

    sales = assign_week_numbers(sales, weekly)
    apply_sales_dpci_mapping_to_cache(sales[["ISBN", "DPCI"]])
    weekly = pd.read_parquet(SALES_CACHE)
    weekly["Week"] = pd.to_datetime(weekly["Week"]).dt.normalize()
    combined_sales = pd.concat([weekly, sales], ignore_index=True)
    combined_sales["Week"] = pd.to_datetime(combined_sales["Week"]).dt.normalize()
    combined_sales = combined_sales.sort_values(["DPCI", "ISBN", "Week"])
    combined_sales = combined_sales.drop_duplicates(subset=["DPCI", "ISBN", "Week"], keep="last")
    combined_sales.to_parquet(SALES_CACHE, index=False)

    metadata = pd.read_parquet(METADATA_CACHE)
    for col in METADATA_COLUMNS:
        if col not in metadata.columns:
            metadata[col] = None
    new_metadata = new_metadata.merge(metadata[["DPCI", "ISBN"]], on=["DPCI", "ISBN"], how="left", indicator=True)
    new_metadata = new_metadata[new_metadata["_merge"].eq("left_only")][["DPCI", "ISBN"]]
    if not new_metadata.empty:
        for col in METADATA_COLUMNS:
            if col not in new_metadata.columns:
                if col == "Price":
                    new_metadata[col] = 0
                elif col == "PubDate":
                    new_metadata[col] = pd.NaT
                else:
                    new_metadata[col] = ""
        metadata = pd.concat([metadata, new_metadata[METADATA_COLUMNS]], ignore_index=True)
        metadata = metadata.drop_duplicates(subset=["DPCI", "ISBN"], keep="first")
        metadata.to_parquet(METADATA_CACHE, index=False)

    print("Refreshing Target title metadata from SQL...")
    sync_eligible_metadata_from_sql(combined_sales["ISBN"], required=True)

    if INVENTORY_CACHE.exists():
        existing_inventory = pd.read_parquet(INVENTORY_CACHE)
    else:
        existing_inventory = pd.DataFrame()
    inventory_all = inventory.copy() if existing_inventory.empty else pd.concat([existing_inventory, inventory], ignore_index=True)
    inventory_all["Week"] = pd.to_datetime(inventory_all["Week"]).dt.normalize()
    inventory_all = inventory_all.drop_duplicates(subset=["DPCI", "Week"], keep="last")
    inventory_all.to_parquet(INVENTORY_CACHE, index=False)

    print(f"\nCache refreshed through {inventory_week:%B %d, %Y}.")
    print(f"  Weekly sales rows: {len(combined_sales):,}")
    print(f"  Inventory rows: {len(inventory_all):,}")
    return inventory_week


def repair_cache_from_sales(use_samples: bool = False) -> None:
    bootstrap_from_skeleton()
    sources = resolve_source_files(use_samples=use_samples)
    inventory_week, _ = load_inventory_snapshot(sources.inventory)
    _, sales, _ = load_sales_file(sources.sales, inventory_week)
    apply_sales_dpci_mapping_to_cache(sales[["ISBN", "DPCI"]])
    print(f"Corrected Target NOC cache DPCI/ISBN mapping from {sources.sales}")


def refresh_title_metadata() -> None:
    sync_eligible_metadata_from_sql(required=True)


def safe_divide(numerator: pd.Series, denominator: pd.Series | float) -> pd.Series:
    result = numerator / denominator
    return result.replace([float("inf"), -float("inf")], 0).fillna(0)


def build_summary(metadata: pd.DataFrame, weekly: pd.DataFrame, yearly: pd.DataFrame, inventory: pd.DataFrame) -> tuple[pd.DataFrame, pd.Timestamp, int]:
    weekly["Week"] = pd.to_datetime(weekly["Week"]).dt.normalize()
    latest_week = weekly["Week"].max()
    current_year = latest_week.year
    prior_year = current_year - 1

    latest_inventory_week = inventory["Week"].max() if not inventory.empty else latest_week
    latest_inventory = inventory[inventory["Week"].eq(latest_inventory_week)].drop_duplicates("DPCI", keep="last")

    keys = metadata[METADATA_COLUMNS].copy()
    keys["DPCI"] = keys["DPCI"].astype(str).str.strip()
    keys["ISBN"] = keys["ISBN"].map(normalize_isbn)

    grouped = weekly.groupby(["DPCI", "ISBN", "Week"], as_index=False).agg({"Units": "sum", "WeekNum": "max"})
    pivot = grouped.pivot_table(index=["DPCI", "ISBN"], columns="Week", values="Units", aggfunc="sum", fill_value=0)

    def sum_between(start: pd.Timestamp, end: pd.Timestamp) -> pd.Series:
        cols = [col for col in pivot.columns if start <= col <= end]
        return pivot[cols].sum(axis=1) if cols else pd.Series(0, index=pivot.index)

    current_weeks = sorted([col for col in pivot.columns if col.year == current_year and col <= latest_week])
    prior_weeks = sorted([col for col in pivot.columns if col.year == prior_year])[: len(current_weeks)]
    last_52 = sorted([col for col in pivot.columns if col <= latest_week])[-52:]
    last_26 = sorted([col for col in pivot.columns if col <= latest_week])[-26:]
    fyly_cols = [col for col in pivot.columns if col.year == prior_year]

    summary = pd.DataFrame(index=pivot.index)
    summary["52 WK"] = pivot[last_52].sum(axis=1) if last_52 else 0
    summary["YTD"] = pivot[current_weeks].sum(axis=1) if current_weeks else 0
    summary["LYTD"] = pivot[prior_weeks].sum(axis=1) if prior_weeks else 0
    summary["FYLY"] = pivot[fyly_cols].sum(axis=1) if fyly_cols else 0
    summary["26 AvgWks"] = pivot[last_26].mean(axis=1) if last_26 else 0

    ltd = pivot.sum(axis=1)
    if not yearly.empty:
        yearly_grouped = yearly.groupby(["DPCI", "ISBN"], as_index=True)["Units"].sum()
        ltd = ltd.add(yearly_grouped, fill_value=0)
    summary["LTD"] = ltd
    summary = summary.reset_index()

    report = keys.merge(summary, on=["DPCI", "ISBN"], how="left").fillna(
        {"52 WK": 0, "YTD": 0, "LYTD": 0, "FYLY": 0, "26 AvgWks": 0, "LTD": 0}
    )
    report = report.merge(latest_inventory.drop(columns=["Week"], errors="ignore"), on="DPCI", how="left")
    for col in ["Store Ct", "On Hand", "INSTK %"]:
        if col not in report.columns:
            report[col] = 0
        report[col] = report[col].fillna(0)

    report["Price"] = report["Price"].map(parse_number)
    report["WOS"] = safe_divide(report["Store Ct"], report["26 AvgWks"] * 26)
    report["UPSPW 52WK"] = safe_divide(report["52 WK"], report["Store Ct"] * 52)
    report["$PSPW 52WK"] = safe_divide(report["52 WK"] * report["Price"], report["Store Ct"] * 52)
    weeks_so_far = max(len(current_weeks), 1)
    report["UPSPW YTD"] = safe_divide(report["YTD"], report["Store Ct"] * weeks_so_far)
    report["YTD Ret$"] = report["YTD"] * report["Price"]
    report["LYTD Ret$"] = report["LYTD"] * report["Price"]
    report["FYLY Ret$"] = report["FYLY"] * report["Price"]
    report["$PSPW YTD"] = safe_divide(report["YTD Ret$"], report["Store Ct"] * weeks_so_far)
    latest_week_values = pivot[latest_week] if latest_week in pivot.columns else pd.Series(0, index=pivot.index)
    latest_week_sort = latest_week_values.reset_index().rename(columns={latest_week: "_LatestWeekSort"})
    report = report.merge(latest_week_sort, on=["DPCI", "ISBN"], how="left")
    report["_LatestWeekSort"] = report["_LatestWeekSort"].fillna(0)

    weeknum_lookup = grouped[["Week", "WeekNum"]].drop_duplicates().sort_values("Week")
    latest_weeknum = int(weeknum_lookup[weeknum_lookup["Week"].eq(latest_week)]["WeekNum"].max())
    report = report.sort_values(["_LatestWeekSort", "YTD", "52 WK", "Title"], ascending=[False, False, False, True])
    report = report.drop(columns=["_LatestWeekSort"])
    return report, latest_week, latest_weeknum


def output_path_for(latest_week: pd.Timestamp, weeknum: int, output_folder: Path | None = None) -> Path:
    folder = output_folder or process_paths.TARGET_NOC_OUTPUT_FOLDER
    filename = f"Week {weeknum:02d} - {latest_week.year} Rolling Target NOC ({latest_week:%m%d%y}).xlsx"
    return folder / filename


def filter_report_to_recent_positive_sales(report: pd.DataFrame, weekly: pd.DataFrame) -> pd.DataFrame:
    weekly = weekly.copy()
    weekly["DPCI"] = weekly["DPCI"].astype(str).str.strip()
    weekly["ISBN"] = weekly["ISBN"].map(normalize_isbn)
    weekly["Week"] = pd.to_datetime(weekly["Week"]).dt.normalize()
    weekly["Units"] = pd.to_numeric(weekly["Units"], errors="coerce").fillna(0)

    recent_keys = weekly[
        weekly["Week"].ge(REPORT_MIN_POSITIVE_SALES_DATE) & weekly["Units"].gt(0)
    ][["DPCI", "ISBN"]].drop_duplicates()

    filtered = report.merge(recent_keys, on=["DPCI", "ISBN"], how="inner")
    removed = len(report) - len(filtered)
    if removed:
        print(
            f"Removed {removed:,} Target title rows with no positive sales since "
            f"{REPORT_MIN_POSITIVE_SALES_DATE:%Y-%m-%d}."
        )
    return filtered


def build_report(
    output_folder: Path | None = None,
    publisher_filter: str | None = None,
    sync_metadata: bool = True,
) -> Path | None:
    bootstrap_from_skeleton()
    metadata = pd.read_parquet(METADATA_CACHE)
    weekly = pd.read_parquet(SALES_CACHE)
    yearly = pd.read_parquet(YEARLY_CACHE)
    inventory = pd.read_parquet(INVENTORY_CACHE) if INVENTORY_CACHE.exists() else pd.DataFrame()
    if weekly.empty:
        raise ValueError("Target NOC weekly sales cache is empty.")

    if sync_metadata:
        print("Refreshing Target title metadata from SQL...")
        metadata = sync_eligible_metadata_from_sql(weekly["ISBN"], required=True)
        weekly = pd.read_parquet(SALES_CACHE)
        yearly = pd.read_parquet(YEARLY_CACHE)

    report, latest_week, weeknum = build_summary(metadata, weekly, yearly, inventory)
    report = filter_report_to_recent_positive_sales(report, weekly)
    if publisher_filter is not None:
        publisher_key = publisher_filter.strip().casefold()
        report = report[
            report["Pub"].astype("string").str.strip().str.casefold().eq(publisher_key)
        ].copy()
        if report.empty:
            print(f"No Target NOC data for {publisher_filter}; skipping partner workbook.")
            return None
    output_path = output_path_for(latest_week, weeknum, output_folder)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    weekly["Week"] = pd.to_datetime(weekly["Week"]).dt.normalize()
    grouped = weekly.groupby(["DPCI", "ISBN", "Week"], as_index=False).agg({"Units": "sum", "WeekNum": "max"})
    weekly_cols = sorted(
        [week for week in grouped["Week"].drop_duplicates() if week.year >= 2023 and week <= latest_week],
        reverse=True,
    )
    weekly_pivot = grouped.pivot_table(index=["DPCI", "ISBN"], columns="Week", values="Units", aggfunc="sum", fill_value=0)
    weeknum_lookup = grouped.drop_duplicates("Week").set_index("Week")["WeekNum"].to_dict()

    historical_years = sorted(
        set(yearly["Year"].dropna().astype(int).tolist())
        | {week.year for week in grouped["Week"].drop_duplicates() if week.year < 2023},
        reverse=True,
    )
    yearly_totals = yearly.groupby(["DPCI", "ISBN", "Year"], as_index=False)["Units"].sum() if not yearly.empty else pd.DataFrame()
    old_weekly = grouped[grouped["Week"].dt.year < 2023].copy()
    if not old_weekly.empty:
        old_weekly["Year"] = old_weekly["Week"].dt.year
        old_weekly = old_weekly.groupby(["DPCI", "ISBN", "Year"], as_index=False)["Units"].sum()
        yearly_totals = pd.concat([yearly_totals, old_weekly], ignore_index=True)
    yearly_pivot = (
        yearly_totals.pivot_table(index=["DPCI", "ISBN"], columns="Year", values="Units", aggfunc="sum", fill_value=0)
        if not yearly_totals.empty
        else pd.DataFrame()
    )

    base_columns = [
        "DPCI",
        "Pub",
        "PT",
        "FT",
        "PGRP",
        "ISBN",
        "Title",
        "Price",
        "PubDate",
        "Store Ct",
        "On Hand",
        "WOS",
        "INSTK %",
        "UPSPW 52WK",
        "$PSPW 52WK",
        "UPSPW YTD",
        "$PSPW YTD",
        "52 WK",
        "YTD",
        "LYTD",
        "FYLY",
        "LTD",
        "YTD Ret$",
        "LYTD Ret$",
        "FYLY Ret$",
        "26 AvgWks",
    ]
    for col in base_columns:
        if col not in report.columns:
            report[col] = None
    output = report[base_columns].copy().reset_index(drop=True)

    extra_columns = {}
    output_keys = list(output[["DPCI", "ISBN"]].itertuples(index=False, name=None))
    for week in weekly_cols:
        values = weekly_pivot[week] if week in weekly_pivot.columns else pd.Series(dtype=float)
        extra_columns[week.strftime("%Y-%m-%d")] = [values.get(key, 0) for key in output_keys]
    for year in historical_years:
        values = yearly_pivot[year] if not yearly_pivot.empty and year in yearly_pivot.columns else pd.Series(dtype=float)
        extra_columns[str(year)] = [values.get(key, 0) for key in output_keys]
    if extra_columns:
        output = pd.concat([output, pd.DataFrame(extra_columns)], axis=1)

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Target NOC")
        writer.sheets["Target NOC"] = worksheet

        title_format = workbook.add_format(
            {"font_size": 16, "bg_color": "#C4BD97", "border": 1, "align": "center", "valign": "vcenter"}
        )
        group_format = workbook.add_format(
            {"bold": True, "bg_color": "#B8CCE4", "align": "center", "valign": "vcenter", "border": 1}
        )
        weeknum_label_format = workbook.add_format(
            {"bold": True, "bg_color": "#DDD9C4", "align": "center", "valign": "vcenter", "border": 1}
        )
        header_format = workbook.add_format(
            {"bold": True, "bg_color": "#CFE8F3", "border": 1, "align": "center", "valign": "vcenter"}
        )
        green_week_format = workbook.add_format(
            {"bold": True, "bg_color": "#D8E4BC", "border": 1, "align": "center", "valign": "vcenter"}
        )
        pink_week_format = workbook.add_format(
            {"bold": True, "bg_color": "#F2DCDB", "border": 1, "align": "center", "valign": "vcenter"}
        )
        green_week_date_format = workbook.add_format(
            {"bold": True, "bg_color": "#D8E4BC", "border": 1, "align": "center", "valign": "vcenter", "num_format": "m/d/yyyy"}
        )
        pink_week_date_format = workbook.add_format(
            {"bold": True, "bg_color": "#F2DCDB", "border": 1, "align": "center", "valign": "vcenter", "num_format": "m/d/yyyy"}
        )
        number_format = workbook.add_format({"num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'})
        decimal_format = workbook.add_format({"num_format": "#,##0.00"})
        money_format = workbook.add_format({"num_format": "$#,##0.00"})
        date_format = workbook.add_format({"num_format": "mm/dd/yyyy"})
        percent_format = workbook.add_format({"num_format": "0.0%"})
        summary_label_format = workbook.add_format(
            {"bold": True, "bg_color": "#CCC0DA", "border": 1, "align": "left", "valign": "vcenter"}
        )
        summary_number_format = workbook.add_format(
            {"bold": True, "bg_color": "#E4DFEC", "border": 1, "align": "right", "num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'}
        )
        summary_decimal_format = workbook.add_format(
            {"bold": True, "bg_color": "#E4DFEC", "border": 1, "align": "right", "num_format": "#,##0.00"}
        )
        summary_percent_format = workbook.add_format(
            {"bold": True, "bg_color": "#E4DFEC", "border": 1, "align": "right", "num_format": "0.0%"}
        )
        blank_summary_format = workbook.add_format({"bg_color": "#E4DFEC", "border": 1})

        worksheet.write(1, 6, "Rolling Target NOC POS", title_format)
        worksheet.write(2, 6, f"Week Ending: {latest_week:%B %d, %Y}", title_format)
        worksheet.merge_range(3, 9, 3, 12, "Inventory", group_format)
        worksheet.merge_range(3, 13, 3, 14, "52WK", group_format)
        worksheet.merge_range(3, 15, 3, 16, "YTD", group_format)
        worksheet.write(3, len(base_columns) - 1, "WeekNum", weeknum_label_format)

        for col_idx, column in enumerate(output.columns):
            if col_idx < len(base_columns):
                header = column
            else:
                parsed = parse_date(column)
                if parsed is not None:
                    month_format = green_week_format if parsed.month % 2 == 1 else pink_week_format
                    worksheet.write(3, col_idx, int(weeknum_lookup.get(parsed, parsed.isocalendar().week)), month_format)
                    header = parsed
                else:
                    worksheet.write(3, col_idx, "YEAR", group_format)
                    header = column
            if isinstance(header, pd.Timestamp):
                month_format = green_week_date_format if header.month % 2 == 1 else pink_week_date_format
                worksheet.write_datetime(4, col_idx, header.to_pydatetime(), month_format)
            else:
                worksheet.write(4, col_idx, header, header_format)

        for row_idx, row in enumerate(output.itertuples(index=False), start=5):
            for col_idx, value in enumerate(row):
                if pd.isna(value):
                    worksheet.write_blank(row_idx, col_idx, None)
                elif isinstance(value, pd.Timestamp):
                    worksheet.write_datetime(row_idx, col_idx, value.to_pydatetime())
                elif isinstance(value, datetime):
                    worksheet.write_datetime(row_idx, col_idx, value)
                else:
                    worksheet.write(row_idx, col_idx, value)

        last_data_row = len(output) + 4
        label_col_idx = 8
        worksheet.write(0, label_col_idx, "Total", summary_label_format)
        worksheet.write(1, label_col_idx, "Subtotal", summary_label_format)
        for row_idx in (0, 1):
            for col_idx in range(label_col_idx + 1, len(output.columns)):
                worksheet.write_blank(row_idx, col_idx, None, blank_summary_format)

        col_index = {column: idx for idx, column in enumerate(output.columns)}
        sum_columns = {
            "Store Ct",
            "On Hand",
            "52 WK",
            "YTD",
            "LYTD",
            "FYLY",
            "LTD",
            "YTD Ret$",
            "LYTD Ret$",
            "FYLY Ret$",
            "26 AvgWks",
        }
        sum_columns.update(
            column
            for column in output.columns[len(base_columns):]
            if parse_date(column) is not None or re.fullmatch(r"\d{4}", str(column))
        )

        def formula_cell(column: str, row: int) -> str:
            return xl_rowcol_to_cell(row, col_index[column])

        def write_summary_formula(row: int, column: str, formula: str, fmt) -> None:
            worksheet.write_formula(row, col_index[column], formula, fmt)

        for column in sum_columns:
            if column not in col_index:
                continue
            start_cell = xl_rowcol_to_cell(5, col_index[column])
            end_cell = xl_rowcol_to_cell(last_data_row, col_index[column])
            total_formula = f"=SUM({start_cell}:{end_cell})"
            subtotal_formula = f"=SUBTOTAL(9,{start_cell}:{end_cell})"
            write_summary_formula(0, column, total_formula, summary_number_format)
            write_summary_formula(1, column, subtotal_formula, summary_number_format)

        if "INSTK %" in col_index:
            start_cell = xl_rowcol_to_cell(5, col_index["INSTK %"])
            end_cell = xl_rowcol_to_cell(last_data_row, col_index["INSTK %"])
            write_summary_formula(0, "INSTK %", f"=AVERAGE({start_cell}:{end_cell})", summary_percent_format)
            write_summary_formula(1, "INSTK %", f"=SUBTOTAL(1,{start_cell}:{end_cell})", summary_percent_format)

        for row in (0, 1):
            store_ct = formula_cell("Store Ct", row)
            avg_26 = formula_cell("26 AvgWks", row)
            wk52 = formula_cell("52 WK", row)
            ytd = formula_cell("YTD", row)
            ytd_dollars = formula_cell("YTD Ret$", row)
            if "WOS" in col_index:
                write_summary_formula(row, "WOS", f'=IFERROR({store_ct}/({avg_26}*26),0)', summary_decimal_format)
            if "UPSPW 52WK" in col_index:
                write_summary_formula(row, "UPSPW 52WK", f'=IFERROR({wk52}/{store_ct}/52,0)', summary_decimal_format)
            if "$PSPW 52WK" in col_index:
                price_range = f"{xl_rowcol_to_cell(5, col_index['Price'])}:{xl_rowcol_to_cell(last_data_row, col_index['Price'])}"
                wk52_range = f"{xl_rowcol_to_cell(5, col_index['52 WK'])}:{xl_rowcol_to_cell(last_data_row, col_index['52 WK'])}"
                write_summary_formula(row, "$PSPW 52WK", f'=IFERROR(SUMPRODUCT({price_range},{wk52_range})/{store_ct}/52,0)', summary_decimal_format)
            if "UPSPW YTD" in col_index:
                write_summary_formula(row, "UPSPW YTD", f'=IFERROR({ytd}/{store_ct}/{max(len([week for week in weekly_cols if week.year == latest_week.year]), 1)},0)', summary_decimal_format)
            if "$PSPW YTD" in col_index:
                write_summary_formula(
                    row,
                    "$PSPW YTD",
                    f'=IFERROR({ytd_dollars}/{store_ct}/{max(len([week for week in weekly_cols if week.year == latest_week.year]), 1)},0)',
                    summary_decimal_format,
                )

        worksheet.autofilter(4, 0, max(4, len(output) + 4), len(output.columns) - 1)
        worksheet.freeze_panes(5, 9)
        worksheet.set_column(0, 0, 14)
        worksheet.set_column(1, 4, 10)
        worksheet.set_column(5, 5, 15)
        worksheet.set_column(6, 6, 36)
        worksheet.set_column(7, 7, 10, money_format)
        worksheet.set_column(8, 8, 12, date_format)
        worksheet.set_column(9, 10, 11, number_format)
        worksheet.set_column(11, 11, 11, decimal_format)
        worksheet.set_column(12, 12, 10, percent_format)
        worksheet.set_column(13, 16, 12, decimal_format)
        worksheet.set_column(17, len(output.columns) - 1, 11, number_format)

    if publisher_filter is None:
        print(f"Saved Target NOC rolling report: {output_path}")
    else:
        print(f"Saved Target NOC rolling report for {publisher_filter}: {output_path}")
    if output_folder is None and publisher_filter is None:
        save_partner_reports(sync_metadata=False)
    return output_path


def save_partner_reports(sync_metadata: bool = True) -> int:
    saved_count = 0
    for publisher, folder in process_paths.TARGET_NOC_DP_FOLDERS.items():
        output_path = build_report(
            output_folder=folder,
            publisher_filter=publisher,
            sync_metadata=sync_metadata,
        )
        if output_path is not None:
            saved_count += 1
    return saved_count


def verticalize_cache(output_folder: Path | None = None) -> Path:
    bootstrap_from_skeleton()
    weekly = pd.read_parquet(SALES_CACHE)
    yearly = pd.read_parquet(YEARLY_CACHE)
    metadata = pd.read_parquet(METADATA_CACHE)
    folder = output_folder or process_paths.TARGET_NOC_CACHE_DIR
    folder.mkdir(parents=True, exist_ok=True)
    output = folder / "target_noc_verticalized_history.xlsx"

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        weekly.merge(metadata, on=["DPCI", "ISBN"], how="left").to_excel(
            writer, sheet_name="Weekly Sales", index=False
        )
        yearly.merge(metadata, on=["DPCI", "ISBN"], how="left").to_excel(
            writer, sheet_name="Yearly Sales", index=False
        )
        if INVENTORY_CACHE.exists():
            pd.read_parquet(INVENTORY_CACHE).to_excel(writer, sheet_name="Inventory", index=False)
    print(f"Saved verticalized Target NOC history: {output}")
    return output


def copy_latest_to_sample() -> None:
    sources = resolve_source_files(use_samples=False)
    if sources.sales.exists() and sources.sales != LOCAL_SAMPLE_SALES:
        shutil.copy2(sources.sales, LOCAL_SAMPLE_SALES)
    if sources.inventory.exists() and sources.inventory != LOCAL_SAMPLE_INVENTORY:
        shutil.copy2(sources.inventory, LOCAL_SAMPLE_INVENTORY)


def run_menu() -> None:
    while True:
        print("\nTarget NOC Rolling Reports")
        print("    1. Update Target NOC Rolling Reports (Full Refresh)")
        print("    2. Build rolling workbook from current cache")
        print("    3. Bootstrap/rebuild cache from skeleton")
        print("    4. Save verticalized history workbook")
        print("    5. Repair cached DPCI/ISBN mapping from current sales file")
        print("    6. Show current source/cache status")
        print("    7. Back to main menu")
        choice = input("\nChoose an option: ").strip().lower()

        try:
            if choice == "1":
                refresh_cache()
                build_report(sync_metadata=False)
            elif choice == "2":
                build_report()
            elif choice == "3":
                bootstrap_from_skeleton(force=True)
            elif choice == "4":
                verticalize_cache()
            elif choice == "5":
                repair_cache_from_sales()
            elif choice == "6":
                show_status()
            elif choice in {"7", "b", "back", "return", "menu"}:
                return
            else:
                print("Invalid choice. Please select a valid option.")
                continue
        except Exception as exc:
            print(f"Target NOC process failed: {exc}")
        input("\nPress Enter to return to the Target NOC menu...")


def show_status() -> None:
    sources = resolve_source_files(use_samples=False)
    print("\nTarget NOC source files:")
    for label, path in {"Sales": sources.sales, "Inventory": sources.inventory}.items():
        if path.exists():
            modified = datetime.fromtimestamp(path.stat().st_mtime)
            print(f"  {label}: {path} (modified {modified:%Y-%m-%d %I:%M %p})")
        else:
            print(f"  {label}: {path} (missing)")

    print("\nTarget NOC cache:")
    for label, path in {
        "Metadata": METADATA_CACHE,
        "Weekly sales": SALES_CACHE,
        "Yearly sales": YEARLY_CACHE,
        "Inventory": INVENTORY_CACHE,
    }.items():
        if path.exists():
            modified = datetime.fromtimestamp(path.stat().st_mtime)
            print(f"  {label}: {path} (modified {modified:%Y-%m-%d %I:%M %p})")
        else:
            print(f"  {label}: {path} (not created yet)")
    if SALES_CACHE.exists():
        weekly = pd.read_parquet(SALES_CACHE)
        if not weekly.empty:
            latest = pd.to_datetime(weekly["Week"]).max()
            print(f"\n  Last cached week: {latest:%B %d, %Y}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Target NOC rolling report process")
    parser.add_argument(
        "command",
        nargs="?",
        default="menu",
        choices=[
            "menu",
            "bootstrap",
            "refresh",
            "build",
            "verticalize",
            "status",
            "repair-cache",
            "refresh-metadata",
        ],
    )
    parser.add_argument("--force", action="store_true", help="Force skeleton cache rebuild for bootstrap.")
    parser.add_argument("--sample", action="store_true", help="Use local sample files instead of latest source-folder files.")
    parser.add_argument("--yes", action="store_true", help="Skip interactive week-ending confirmation.")
    parser.add_argument("--local-output", action="store_true", help="Write report output to target_rolling_report/output for testing.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_folder = LOCAL_DIR / "output" if args.local_output else None

    if args.command == "menu":
        run_menu()
    elif args.command == "bootstrap":
        bootstrap_from_skeleton(force=args.force)
    elif args.command == "refresh":
        refresh_cache(use_samples=args.sample, assume_yes=args.yes)
    elif args.command == "build":
        build_report(output_folder=output_folder)
    elif args.command == "verticalize":
        verticalize_cache(output_folder=output_folder)
    elif args.command == "status":
        show_status()
    elif args.command == "repair-cache":
        repair_cache_from_sales(use_samples=args.sample)
    elif args.command == "refresh-metadata":
        refresh_title_metadata()


if __name__ == "__main__":
    main()
