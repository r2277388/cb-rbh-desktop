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
SOURCE_COLUMNS = ["ISBN", "Title", "Sea", "Pub Grp", "Task Name", "Due Date", "Release Date", "Price"]
CACHE_COLUMNS = ["CacheDate", *SOURCE_COLUMNS]
ISBN_FIELDS = ["Sea", "Release Date", "Price"]
TASK_FIELDS = ["Due Date"]
PUB_GROUP_FILTERS = ("ENT", "ART", "GAM", "FWN", "LIF", "CPA")
TASK_INTERVAL_WEEKS = {
    "Cover Due": 4,
    "Title memo to DES": 4,
    "1st galley due": 2,
    "MS to DES": 2,
    "Print files out": 2,
    "On Sale Date": 1,
}
INTERVAL_ANCHOR_DATE = pd.Timestamp("2026-05-04")


def default_report_file() -> Path:
    date_prefix = datetime.now().strftime("%Y_%m_%d")
    return process_paths.GEN_EDITORIAL_REPORT_FILE.with_name(
        f"{date_prefix}_GenEd_Data_Deltas.xlsx"
    )


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


def future_seasons(reference_date: pd.Timestamp | None = None) -> list[str]:
    reference_date = pd.Timestamp(reference_date or today_cache_date()).normalize()
    year = reference_date.year
    season = "Spring" if reference_date.month <= 6 else "Fall"
    seasons: list[str] = []
    if season == "Spring":
        year_season = (year, "Fall")
    else:
        year_season = (year + 1, "Spring")
    for _ in range(3):
        season_year, season_name = year_season
        seasons.append(f"{season_year}-{season_name}")
        if season_name == "Spring":
            year_season = (season_year, "Fall")
        else:
            year_season = (season_year + 1, "Spring")
    return seasons


def is_business_day(value: pd.Timestamp) -> bool:
    return value.weekday() < 5


def is_archive_day(value: pd.Timestamp) -> bool:
    return value.weekday() == 0


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
    snapshot["Pub Grp"] = snapshot["Pub Grp"].map(normalize_text)
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
    cache["Pub Grp"] = cache["Pub Grp"].map(normalize_text)
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
    if not allow_weekend and not is_archive_day(cache_date):
        print(f"{cache_date:%Y-%m-%d} is not a Monday archive day. No snapshot archived.")
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


def variation_columns() -> list[str]:
    return [
        "ISBN",
        "Title",
        "Sea",
        "Pub Grp",
        "Variation Type",
        "Task Name",
        "Previous Cache Date",
        "Previous Value",
        "Current Cache Date",
        "Current Value",
    ]


def filtered_activity_cache(cache: pd.DataFrame, reference_date: pd.Timestamp | None = None) -> pd.DataFrame:
    seasons = future_seasons(reference_date)
    return cache[
        cache["Sea"].isin(seasons)
        & cache["Pub Grp"].isin(PUB_GROUP_FILTERS)
        & cache["Task Name"].isin(TASK_INTERVAL_WEEKS)
    ].copy()


def is_interval_date(cache_date: pd.Timestamp, interval_weeks: int) -> bool:
    cache_date = pd.Timestamp(cache_date).normalize()
    days_since_anchor = (cache_date - INTERVAL_ANCHOR_DATE).days
    interval_days = interval_weeks * 7
    return days_since_anchor >= interval_days and days_since_anchor % interval_days == 0


def interval_label(interval_weeks: int) -> str:
    if interval_weeks == 1:
        return "Every 1 Week - a weekly comparison"
    return f"Every {interval_weeks} Weeks - a {interval_weeks} week comparison"


def build_stipulations(reference_date: pd.Timestamp | None = None) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    rows.extend({"Stipulation": "Pub Grp", "Value": value, "Update Variation Interval": ""} for value in PUB_GROUP_FILTERS)
    rows.extend({"Stipulation": "Sea", "Value": value, "Update Variation Interval": ""} for value in future_seasons(reference_date))
    rows.extend(
        {
            "Stipulation": "Task Name",
            "Value": task_name,
            "Update Variation Interval": interval_label(interval_weeks),
        }
        for task_name, interval_weeks in TASK_INTERVAL_WEEKS.items()
    )
    return pd.DataFrame(rows)


def build_variation_rows(cache: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    if cache.empty:
        return pd.DataFrame(columns=variation_columns())

    cache = filtered_activity_cache(cache)
    if cache.empty:
        return pd.DataFrame(columns=variation_columns())

    task_level = (
        cache[["CacheDate", "ISBN", "Title", "Sea", "Pub Grp", "Task Name", *TASK_FIELDS]]
        .sort_values(["CacheDate", "ISBN", "Task Name"])
        .groupby(["CacheDate", "ISBN", "Task Name"], as_index=False)
        .agg({"Title": last_non_blank, "Sea": last_non_blank, "Pub Grp": last_non_blank, "Due Date": last_non_blank})
    )

    for (isbn, task_name), group in task_level.groupby(["ISBN", "Task Name"], dropna=False):
        group = group.sort_values("CacheDate")
        interval_weeks = TASK_INTERVAL_WEEKS.get(task_name)
        if interval_weeks is None:
            continue
        interval_days = interval_weeks * 7
        for _, row in group.iterrows():
            current_date = pd.Timestamp(row["CacheDate"]).normalize()
            if not is_interval_date(current_date, interval_weeks):
                continue
            previous_group = group[group["CacheDate"].le(current_date - pd.Timedelta(days=interval_days))]
            if previous_group.empty:
                continue
            previous_row = previous_group.iloc[-1]
            previous_key = comparable_value(previous_row["Due Date"], "Due Date")
            current_key = comparable_value(row["Due Date"], "Due Date")
            if current_key != previous_key:
                rows.append(
                    {
                        "ISBN": isbn,
                        "Title": row["Title"],
                        "Sea": row["Sea"],
                        "Pub Grp": row["Pub Grp"],
                        "Variation Type": "Due Date",
                        "Task Name": task_name,
                        "Previous Cache Date": previous_row["CacheDate"],
                        "Previous Value": display_value(previous_row["Due Date"], "Due Date"),
                        "Current Cache Date": current_date,
                        "Current Value": display_value(row["Due Date"], "Due Date"),
                    }
                )

    if not rows:
        return pd.DataFrame(columns=variation_columns())
    variations = pd.DataFrame(rows)
    variations = variations.sort_values(["Current Cache Date", "ISBN", "Task Name"])
    return variations[variation_columns()]


def build_current_snapshot(cache: pd.DataFrame) -> pd.DataFrame:
    if cache.empty:
        return pd.DataFrame(columns=CACHE_COLUMNS)
    latest_date = cache["CacheDate"].max()
    return cache[cache["CacheDate"].eq(latest_date)].copy()


def save_variation_report(output_file: Path | None = None) -> Path:
    cache = load_cache()
    output_file = output_file or default_report_file()
    output_file.parent.mkdir(parents=True, exist_ok=True)

    variations = build_variation_rows(cache)
    stipulations = build_stipulations()

    with pd.ExcelWriter(output_file, engine="xlsxwriter", datetime_format="m/d/yyyy", date_format="m/d/yyyy") as writer:
        variations.to_excel(writer, sheet_name="Variations", index=False)
        stipulations.to_excel(writer, sheet_name="Stipulations", index=False)

        workbook = writer.book
        date_format = workbook.add_format({"num_format": "m/d/yyyy"})
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, worksheet.dim_rowmax or 0, worksheet.dim_colmax or 0)
        if "Variations" in writer.sheets:
            worksheet = writer.sheets["Variations"]
            worksheet.set_column("A:A", 15)
            worksheet.set_column("B:B", 42)
            worksheet.set_column("C:C", 14)
            worksheet.set_column("D:D", 10)
            worksheet.set_column("E:E", 16)
            worksheet.set_column("F:F", 28)
            worksheet.set_column("G:G", 18, date_format)
            worksheet.set_column("H:H", 18)
            worksheet.set_column("I:I", 18, date_format)
            worksheet.set_column("J:J", 18)
        if "Stipulations" in writer.sheets:
            worksheet = writer.sheets["Stipulations"]
            worksheet.set_column("A:A", 16)
            worksheet.set_column("B:B", 28)
            worksheet.set_column("C:C", 38)

    print(f"Saved General Editorial Data Variations report: {output_file}")
    print(f"  Cache dates:       {cache['CacheDate'].nunique() if not cache.empty else 0:,}")
    print(f"  Cache rows:        {len(cache):,}")
    print(f"  Variation rows:    {len(variations):,}")
    latest_cache_date = cache["CacheDate"].max() if not cache.empty else pd.NaT
    if pd.isna(latest_cache_date):
        print("  Latest cache date: none")
    else:
        print(f"  Latest cache date: {latest_cache_date:%Y-%m-%d}")
    print(f"  Source workbook:   {process_paths.GEN_EDITORIAL_SOURCE_WORKBOOK}")
    return output_file


def run_process(allow_weekend: bool = False) -> Path:
    archive_snapshot(allow_weekend=allow_weekend)
    return save_variation_report()


def print_status() -> None:
    source = process_paths.GEN_EDITORIAL_SOURCE_WORKBOOK
    cache_file = process_paths.GEN_EDITORIAL_CACHE_FILE
    report_file = default_report_file()
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


def print_process_description() -> None:
    print("\nProcess Description")
    print()
    print("General Editorial Data Variations tracks changes in Chronicle editorial schedule data.")
    print("This process was requested by Meghan Clarke on April 30, 2026.")
    print()
    print("Source:")
    print(f"  Workbook: {process_paths.GEN_EDITORIAL_SOURCE_WORKBOOK}")
    print("  Filter:   Publisher = Chronicle")
    print("  Fields:   ISBN, Title, Sea, Pub Grp, Task Name, Due Date, Release Date, Price")
    print()
    print("Schedule:")
    print(f"  Automatic archive/report run: {process_paths.GEN_EDITORIAL_SCHEDULE_DESCRIPTION}")
    print("  Manual archive: available from this menu when an extra Monday comparison point is needed.")
    print()
    print("Cache:")
    print(f"  File: {process_paths.GEN_EDITORIAL_CACHE_FILE}")
    print("  Each archived row includes CacheDate, which identifies the snapshot date.")
    print()
    print("Report:")
    print(f"  File: {default_report_file()}")
    print("  The Variations sheet shows task due date changes over time by ISBN.")
    print(f"  Pub Grp filter: {', '.join(PUB_GROUP_FILTERS)}")
    print(f"  Sea filter: {', '.join(future_seasons())}")
    print(f"  Task Name filter: {', '.join(TASK_INTERVAL_WEEKS)}")


def run_menu() -> None:
    while True:
        print("\nGeneral Editorial Data Variations")
        print("    1. Run Monday archive + variation report")
        print("    2. Manual archive today's snapshot only")
        print("    3. Build variation report from cache")
        print("    4. Show cache/source status")
        print("    5. Process Description")
        print("    6. Back to previous menu")
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
            elif choice == "5":
                print_process_description()
            elif choice in {"6", "b", "back", "return", "menu"}:
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
        choices=["menu", "run", "archive", "manual-archive", "report", "status", "description"],
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
    elif args.command == "description":
        print_process_description()


if __name__ == "__main__":
    main()
