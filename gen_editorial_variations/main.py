from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    from paths import process_paths
except ImportError:
    import sys

    sys.path.append(str(Path(__file__).resolve().parents[1]))
    from paths import process_paths


SOURCE_SHEET = "Schedule by Publ Group"
PUBLISHER_FILTER = "Chronicle"
SOURCE_COLUMNS = ["ISBN", "Title", "Sea", "Task Name", "Due Date", "Release Date", "Price"]
CACHE_COLUMNS = ["CacheDate", *SOURCE_COLUMNS]
ISBN_FIELDS = ["Sea", "Release Date", "Price"]
TASK_FIELDS = ["Due Date"]


def normalize_isbn(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    digits = "".join(char for char in text if char.isdigit())
    if not digits:
        return ""
    if len(digits) < 13:
        return digits.zfill(13)
    if len(digits) > 13 and digits.startswith("0"):
        return digits[-13:]
    return digits[:13]


def normalize_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def normalize_date(value: object) -> pd.Timestamp | pd.NaT:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return pd.NaT
    return pd.Timestamp(parsed).normalize()


def normalize_price(value: object) -> float | pd.NA:
    parsed = pd.to_numeric(value, errors="coerce")
    if pd.isna(parsed):
        return pd.NA
    return float(parsed)


def today_cache_date() -> pd.Timestamp:
    return pd.Timestamp.today().normalize()


def is_business_day(value: pd.Timestamp) -> bool:
    return value.weekday() < 5


def load_source_snapshot(cache_date: pd.Timestamp | None = None) -> pd.DataFrame:
    source = process_paths.GEN_EDITORIAL_SOURCE_WORKBOOK
    if not source.exists():
        raise FileNotFoundError(f"General Editorial source workbook not found: {source}")

    df = pd.read_excel(source, sheet_name=SOURCE_SHEET, dtype=object)
    missing = [column for column in ["Publisher", *SOURCE_COLUMNS] if column not in df.columns]
    if missing:
        raise ValueError(f"Source workbook is missing columns: {', '.join(missing)}")

    filtered = df[df["Publisher"].astype("string").str.strip().eq(PUBLISHER_FILTER)].copy()
    snapshot = filtered[SOURCE_COLUMNS].copy()
    snapshot["ISBN"] = snapshot["ISBN"].map(normalize_isbn)
    snapshot["Title"] = snapshot["Title"].map(normalize_text)
    snapshot["Sea"] = snapshot["Sea"].map(normalize_text)
    snapshot["Task Name"] = snapshot["Task Name"].map(normalize_text)
    snapshot["Due Date"] = snapshot["Due Date"].map(normalize_date)
    snapshot["Release Date"] = snapshot["Release Date"].map(normalize_date)
    snapshot["Price"] = snapshot["Price"].map(normalize_price)
    snapshot = snapshot[snapshot["ISBN"].ne("")].copy()
    snapshot.insert(0, "CacheDate", cache_date or today_cache_date())
    snapshot["CacheDate"] = pd.to_datetime(snapshot["CacheDate"]).dt.normalize()
    return snapshot[CACHE_COLUMNS]


def load_cache() -> pd.DataFrame:
    cache_file = process_paths.GEN_EDITORIAL_CACHE_FILE
    if not cache_file.exists():
        return pd.DataFrame(columns=CACHE_COLUMNS)
    cache = pd.read_parquet(cache_file)
    for column in CACHE_COLUMNS:
        if column not in cache.columns:
            cache[column] = pd.NA
    cache["CacheDate"] = pd.to_datetime(cache["CacheDate"], errors="coerce").dt.normalize()
    cache["ISBN"] = cache["ISBN"].map(normalize_isbn)
    cache["Title"] = cache["Title"].map(normalize_text)
    cache["Sea"] = cache["Sea"].map(normalize_text)
    cache["Task Name"] = cache["Task Name"].map(normalize_text)
    cache["Due Date"] = cache["Due Date"].map(normalize_date)
    cache["Release Date"] = cache["Release Date"].map(normalize_date)
    cache["Price"] = pd.to_numeric(cache["Price"], errors="coerce")
    return cache[CACHE_COLUMNS]


def save_cache(cache: pd.DataFrame) -> None:
    cache_file = process_paths.GEN_EDITORIAL_CACHE_FILE
    cache_file.parent.mkdir(parents=True, exist_ok=True)
    cache.to_parquet(cache_file, index=False)


def archive_snapshot(cache_date: pd.Timestamp | None = None, allow_weekend: bool = False) -> pd.DataFrame:
    cache_date = cache_date or today_cache_date()
    cache_date = pd.Timestamp(cache_date).normalize()
    if not allow_weekend and not is_business_day(cache_date):
        print(f"{cache_date:%Y-%m-%d} is not a business day. No snapshot archived.")
        return load_cache()

    snapshot = load_source_snapshot(cache_date)
    cache = load_cache()
    if not cache.empty:
        cache = cache[~cache["CacheDate"].eq(cache_date)].copy()
    combined = snapshot if cache.empty else pd.concat([cache, snapshot], ignore_index=True)
    combined = combined.sort_values(["CacheDate", "ISBN", "Task Name"]).reset_index(drop=True)
    save_cache(combined)
    print(f"Archived {len(snapshot):,} General Editorial rows for {cache_date:%Y-%m-%d}.")
    print(f"Cache rows: {len(combined):,}")
    return combined


def archive_manual_snapshot() -> pd.DataFrame:
    return archive_snapshot(allow_weekend=True)


def comparable_value(value: object, field: str):
    if field in {"Due Date", "Release Date"}:
        parsed = pd.to_datetime(value, errors="coerce")
        return None if pd.isna(parsed) else pd.Timestamp(parsed).date().isoformat()
    if field == "Price":
        parsed = pd.to_numeric(value, errors="coerce")
        return None if pd.isna(parsed) else round(float(parsed), 2)
    text = normalize_text(value)
    return text or None


def display_value(value: object, field: str) -> object:
    if field in {"Due Date", "Release Date"}:
        parsed = pd.to_datetime(value, errors="coerce")
        return pd.NaT if pd.isna(parsed) else pd.Timestamp(parsed).normalize()
    if field == "Price":
        parsed = pd.to_numeric(value, errors="coerce")
        return pd.NA if pd.isna(parsed) else float(parsed)
    return normalize_text(value)


def latest_title(group: pd.DataFrame) -> str:
    values = group.sort_values("CacheDate")["Title"].dropna().astype(str)
    values = values[values.str.strip().ne("")]
    return values.iloc[-1] if not values.empty else ""


def last_non_blank(series: pd.Series):
    values = series.dropna()
    if values.empty:
        return pd.NA
    non_blank = values[values.astype(str).str.strip().ne("")]
    if not non_blank.empty:
        return non_blank.iloc[-1]
    return values.iloc[-1]


def build_variation_rows(cache: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    if cache.empty:
        return pd.DataFrame()

    isbn_level = (
        cache[["CacheDate", "ISBN", "Title", *ISBN_FIELDS]]
        .sort_values(["CacheDate", "ISBN"])
        .groupby(["CacheDate", "ISBN"], as_index=False)
        .agg({"Title": last_non_blank, "Sea": last_non_blank, "Release Date": last_non_blank, "Price": last_non_blank})
    )
    for (isbn, field), group in isbn_level.melt(
        id_vars=["CacheDate", "ISBN", "Title"],
        value_vars=ISBN_FIELDS,
        var_name="Field",
        value_name="Value",
    ).groupby(["ISBN", "Field"], dropna=False):
        group = group.sort_values("CacheDate")
        previous_key = None
        previous_value = None
        previous_date = None
        for row in group.itertuples(index=False):
            current_key = comparable_value(row.Value, field)
            if previous_key is not None and current_key != previous_key:
                rows.append(
                    {
                        "ISBN": isbn,
                        "Title": latest_title(group),
                        "Variation Type": field,
                        "Task Name": "",
                        "Previous Cache Date": previous_date,
                        "Previous Value": display_value(previous_value, field),
                        "Current Cache Date": row.CacheDate,
                        "Current Value": display_value(row.Value, field),
                    }
                )
            previous_key = current_key
            previous_value = row.Value
            previous_date = row.CacheDate

    task_level = (
        cache[["CacheDate", "ISBN", "Title", "Task Name", *TASK_FIELDS]]
        .sort_values(["CacheDate", "ISBN", "Task Name"])
        .groupby(["CacheDate", "ISBN", "Task Name"], as_index=False)
        .agg({"Title": last_non_blank, "Due Date": last_non_blank})
    )
    for (isbn, task_name, field), group in task_level.melt(
        id_vars=["CacheDate", "ISBN", "Title", "Task Name"],
        value_vars=TASK_FIELDS,
        var_name="Field",
        value_name="Value",
    ).groupby(["ISBN", "Task Name", "Field"], dropna=False):
        group = group.sort_values("CacheDate")
        previous_key = None
        previous_value = None
        previous_date = None
        for row in group.itertuples(index=False):
            current_key = comparable_value(row.Value, field)
            if previous_key is not None and current_key != previous_key:
                rows.append(
                    {
                        "ISBN": isbn,
                        "Title": latest_title(group),
                        "Variation Type": field,
                        "Task Name": task_name,
                        "Previous Cache Date": previous_date,
                        "Previous Value": display_value(previous_value, field),
                        "Current Cache Date": row.CacheDate,
                        "Current Value": display_value(row.Value, field),
                    }
                )
            previous_key = current_key
            previous_value = row.Value
            previous_date = row.CacheDate

    if not rows:
        return pd.DataFrame(
            columns=[
                "ISBN",
                "Title",
                "Variation Type",
                "Task Name",
                "Previous Cache Date",
                "Previous Value",
                "Current Cache Date",
                "Current Value",
            ]
        )
    variations = pd.DataFrame(rows)
    variations = variations.sort_values(["Current Cache Date", "ISBN", "Variation Type", "Task Name"])
    return variations


def build_current_snapshot(cache: pd.DataFrame) -> pd.DataFrame:
    if cache.empty:
        return pd.DataFrame(columns=CACHE_COLUMNS)
    latest_date = cache["CacheDate"].max()
    return cache[cache["CacheDate"].eq(latest_date)].copy()


def save_variation_report(output_file: Path | None = None) -> Path:
    cache = load_cache()
    output_file = output_file or process_paths.GEN_EDITORIAL_REPORT_FILE
    output_file.parent.mkdir(parents=True, exist_ok=True)

    variations = build_variation_rows(cache)
    current = build_current_snapshot(cache)
    summary = pd.DataFrame(
        [
            {"Metric": "Cache dates", "Value": cache["CacheDate"].nunique() if not cache.empty else 0},
            {"Metric": "Cache rows", "Value": len(cache)},
            {"Metric": "Variation rows", "Value": len(variations)},
            {
                "Metric": "Latest cache date",
                "Value": cache["CacheDate"].max() if not cache.empty else pd.NaT,
            },
            {"Metric": "Source workbook", "Value": str(process_paths.GEN_EDITORIAL_SOURCE_WORKBOOK)},
        ]
    )

    with pd.ExcelWriter(output_file, engine="xlsxwriter", datetime_format="m/d/yyyy", date_format="m/d/yyyy") as writer:
        variations.to_excel(writer, sheet_name="Variations", index=False)
        current.to_excel(writer, sheet_name="Current Snapshot", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

        workbook = writer.book
        date_format = workbook.add_format({"num_format": "m/d/yyyy"})
        money_format = workbook.add_format({"num_format": "$#,##0.00"})
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, worksheet.dim_rowmax or 0, worksheet.dim_colmax or 0)
        if "Variations" in writer.sheets:
            worksheet = writer.sheets["Variations"]
            worksheet.set_column("A:A", 15)
            worksheet.set_column("B:B", 42)
            worksheet.set_column("C:C", 16)
            worksheet.set_column("D:D", 28)
            worksheet.set_column("E:E", 18, date_format)
            worksheet.set_column("F:F", 18)
            worksheet.set_column("G:G", 18, date_format)
            worksheet.set_column("H:H", 18)
        if "Current Snapshot" in writer.sheets:
            worksheet = writer.sheets["Current Snapshot"]
            worksheet.set_column("A:A", 12, date_format)
            worksheet.set_column("B:B", 15)
            worksheet.set_column("C:C", 42)
            worksheet.set_column("D:D", 13)
            worksheet.set_column("E:E", 28)
            worksheet.set_column("F:G", 12, date_format)
            worksheet.set_column("H:H", 10, money_format)
        if "Summary" in writer.sheets:
            writer.sheets["Summary"].set_column("A:B", 34)

    print(f"Saved General Editorial Data Variations report: {output_file}")
    return output_file


def run_process(allow_weekend: bool = False) -> Path:
    archive_snapshot(allow_weekend=allow_weekend)
    return save_variation_report()


def print_status() -> None:
    source = process_paths.GEN_EDITORIAL_SOURCE_WORKBOOK
    cache_file = process_paths.GEN_EDITORIAL_CACHE_FILE
    report_file = process_paths.GEN_EDITORIAL_REPORT_FILE
    print("\nGeneral Editorial Data Variations")
    print(f"  Source workbook: {source}")
    if source.exists():
        print(f"  Source modified: {datetime.fromtimestamp(source.stat().st_mtime):%Y-%m-%d %I:%M %p}")
    else:
        print("  Source modified: missing")
    print(f"  Cache file:      {cache_file}")
    print(f"  Report file:     {report_file}")
    cache = load_cache()
    if cache.empty:
        print("  Cache status:    No snapshots archived yet")
        return
    print(f"  Cache rows:      {len(cache):,}")
    print(f"  Cache dates:     {cache['CacheDate'].nunique():,}")
    print(f"  Earliest date:   {cache['CacheDate'].min():%Y-%m-%d}")
    print(f"  Latest date:     {cache['CacheDate'].max():%Y-%m-%d}")
    print(f"  Latest rows:     {len(build_current_snapshot(cache)):,}")


def run_menu() -> None:
    while True:
        print("\nGeneral Editorial Data Variations")
        print("    1. Run daily archive + variation report")
        print("    2. Manual archive today's snapshot only")
        print("    3. Build variation report from cache")
        print("    4. Show cache/source status")
        print("    5. Back to previous menu")
        choice = input("\nChoose an option: ").strip().lower()

        try:
            if choice == "1":
                run_process()
            elif choice == "2":
                archive_manual_snapshot()
            elif choice == "3":
                save_variation_report()
            elif choice == "4":
                print_status()
            elif choice in {"5", "b", "back", "return", "menu"}:
                return
            else:
                print("Invalid choice. Please select a valid option.")
                continue
        except Exception as exc:
            print(f"General Editorial Data Variations failed: {exc}")
        input("\nPress Enter to return to the General Editorial menu...")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="General Editorial data variation tracking.")
    parser.add_argument(
        "command",
        nargs="?",
        default="menu",
        choices=["menu", "run", "archive", "manual-archive", "report", "status"],
    )
    parser.add_argument("--allow-weekend", action="store_true", help="Allow archive runs on weekends.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.command == "menu":
        run_menu()
    elif args.command == "run":
        run_process(allow_weekend=args.allow_weekend)
    elif args.command == "archive":
        archive_snapshot(allow_weekend=args.allow_weekend)
    elif args.command == "manual-archive":
        archive_manual_snapshot()
    elif args.command == "report":
        save_variation_report()
    elif args.command == "status":
        print_status()


if __name__ == "__main__":
    main()
