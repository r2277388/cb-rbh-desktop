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
EDITORIAL_DETAIL_COLUMNS = [
    "Managing Editor",
    "Project Editor",
    "Designer",
    "Prod Developer",
    "Prod Designer",
]
SOURCE_COLUMNS = [
    "ISBN",
    "Title",
    *EDITORIAL_DETAIL_COLUMNS,
    "Sea",
    "Pub Grp",
    "Task Name",
    "Due Date",
    "Release Date",
    "Price",
]
CACHE_COLUMNS = ["CacheDate", *SOURCE_COLUMNS]
ISBN_FIELDS = ["Sea", "Release Date", "Price"]
TASK_FIELDS = ["Due Date"]
PUB_GROUP_FILTERS = ("ENT", "ART", "GAM", "FWN", "LIF", "CPA")
TASK_CHANGE_THRESHOLD_WEEKS = {
    "Cover Due": 4,
    "Title memo to DES": 4,
    "1st galley due": 2,
    "MS to DES": 2,
    "Print files out": 2,
    "On Sale Date": 1,
}
RECENT_VARIATION_WEEK_COUNT = 6
HEADER_FILL_COLOR = "#BFBFBF"
ISBN_GROUP_FILL_COLORS = ("#DCE6F1", "#E4DFEC")


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
    for column in EDITORIAL_DETAIL_COLUMNS:
        snapshot[column] = snapshot[column].map(normalize_text)
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
    for column in EDITORIAL_DETAIL_COLUMNS:
        cache[column] = cache[column].map(normalize_text)
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
        "Current Cache Date",
        "Previous Value",
        "Current Value",
        "Schedule Movement",
        *EDITORIAL_DETAIL_COLUMNS,
    ]


def filtered_activity_cache(cache: pd.DataFrame, reference_date: pd.Timestamp | None = None) -> pd.DataFrame:
    seasons = future_seasons(reference_date)
    return cache[
        cache["Sea"].isin(seasons)
        & cache["Pub Grp"].isin(PUB_GROUP_FILTERS)
        & cache["Task Name"].isin(TASK_CHANGE_THRESHOLD_WEEKS)
    ].copy()


def threshold_label(threshold_weeks: int) -> str:
    if threshold_weeks == 1:
        return "Weekly comparison; show if change is greater than 1 week"
    return f"Weekly comparison; show if change is greater than {threshold_weeks} weeks"


def schedule_movement_label(previous_value: object, current_value: object) -> tuple[int, str] | None:
    previous_date = pd.to_datetime(previous_value, errors="coerce")
    current_date = pd.to_datetime(current_value, errors="coerce")
    if pd.isna(previous_date) or pd.isna(current_date):
        return None

    delta_days = int((pd.Timestamp(current_date).normalize() - pd.Timestamp(previous_date).normalize()).days)
    if delta_days <= 0:
        return None

    weeks, days = divmod(delta_days, 7)
    if days:
        label = f"{weeks:02d} Weeks {days:02d} Days"
    else:
        label = f"{weeks:02d} Weeks"
    return delta_days, label


def sort_variations(variations: pd.DataFrame) -> pd.DataFrame:
    if variations.empty:
        return variations
    sorted_variations = variations.copy()
    sorted_variations["_CurrentValueSort"] = pd.to_datetime(
        sorted_variations["Current Value"], errors="coerce"
    )
    sorted_variations = sorted_variations.sort_values(
        ["ISBN", "_CurrentValueSort", "Current Value", "Task Name"],
        kind="stable",
        na_position="last",
    )
    return sorted_variations.drop(columns="_CurrentValueSort")


def build_stipulations(reference_date: pd.Timestamp | None = None) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    rows.extend({"Stipulation": "Pub Grp", "Value": value, "Update Variation Threshold": ""} for value in PUB_GROUP_FILTERS)
    rows.extend({"Stipulation": "Sea", "Value": value, "Update Variation Threshold": ""} for value in future_seasons(reference_date))
    rows.extend(
        {
            "Stipulation": "Task Name",
            "Value": task_name,
            "Update Variation Threshold": threshold_label(threshold_weeks),
        }
        for task_name, threshold_weeks in TASK_CHANGE_THRESHOLD_WEEKS.items()
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
        cache[["CacheDate", "ISBN", "Title", *EDITORIAL_DETAIL_COLUMNS, "Sea", "Pub Grp", "Task Name", *TASK_FIELDS]]
        .sort_values(["CacheDate", "ISBN", "Task Name"])
        .groupby(["CacheDate", "ISBN", "Task Name"], as_index=False)
        .agg(
            {
                "Title": last_non_blank,
                **{column: last_non_blank for column in EDITORIAL_DETAIL_COLUMNS},
                "Sea": last_non_blank,
                "Pub Grp": last_non_blank,
                "Due Date": last_non_blank,
            }
        )
    )

    for (isbn, task_name), group in task_level.groupby(["ISBN", "Task Name"], dropna=False):
        group = group.sort_values("CacheDate")
        threshold_weeks = TASK_CHANGE_THRESHOLD_WEEKS.get(task_name)
        if threshold_weeks is None:
            continue
        threshold_days = threshold_weeks * 7
        for index, (_, row) in enumerate(group.iterrows()):
            if index == 0:
                continue
            current_date = pd.Timestamp(row["CacheDate"]).normalize()
            previous_row = group.iloc[index - 1]
            previous_key = comparable_value(previous_row["Due Date"], "Due Date")
            current_key = comparable_value(row["Due Date"], "Due Date")
            movement = schedule_movement_label(previous_row["Due Date"], row["Due Date"])
            if current_key != previous_key and movement is not None and movement[0] > threshold_days:
                rows.append(
                    {
                        "ISBN": isbn,
                        "Title": row["Title"],
                        **{column: row[column] for column in EDITORIAL_DETAIL_COLUMNS},
                        "Sea": row["Sea"],
                        "Pub Grp": row["Pub Grp"],
                        "Variation Type": "Due Date",
                        "Task Name": task_name,
                        "Previous Cache Date": previous_row["CacheDate"],
                        "Current Cache Date": current_date,
                        "Previous Value": display_value(previous_row["Due Date"], "Due Date"),
                        "Current Value": display_value(row["Due Date"], "Due Date"),
                        "Schedule Movement": movement[1],
                    }
                )

    if not rows:
        return pd.DataFrame(columns=variation_columns())
    variations = pd.DataFrame(rows)
    variations = sort_variations(variations)
    return variations[variation_columns()]


def recent_cache_date_pairs(cache: pd.DataFrame, week_count: int = RECENT_VARIATION_WEEK_COUNT) -> list[tuple[pd.Timestamp, pd.Timestamp]]:
    if cache.empty:
        return []
    dates = (
        pd.to_datetime(cache["CacheDate"], errors="coerce")
        .dropna()
        .dt.normalize()
        .drop_duplicates()
        .sort_values()
        .tolist()
    )
    pairs = [(pd.Timestamp(dates[index - 1]), pd.Timestamp(dates[index])) for index in range(1, len(dates))]
    return pairs[-week_count:]


def variations_for_cache_pairs(
    variations: pd.DataFrame,
    cache_date_pairs: list[tuple[pd.Timestamp, pd.Timestamp]],
) -> pd.DataFrame:
    if variations.empty or not cache_date_pairs:
        return pd.DataFrame(columns=variation_columns())

    normalized = variations.copy()
    normalized["Previous Cache Date"] = pd.to_datetime(
        normalized["Previous Cache Date"], errors="coerce"
    ).dt.normalize()
    normalized["Current Cache Date"] = pd.to_datetime(
        normalized["Current Cache Date"], errors="coerce"
    ).dt.normalize()

    mask = pd.Series(False, index=normalized.index)
    for previous_date, current_date in cache_date_pairs:
        mask = mask | (
            normalized["Previous Cache Date"].eq(previous_date)
            & normalized["Current Cache Date"].eq(current_date)
        )
    filtered = normalized[mask].copy()
    if filtered.empty:
        return pd.DataFrame(columns=variation_columns())
    filtered = sort_variations(filtered)
    return filtered[variation_columns()]


def source_editorial_details() -> pd.DataFrame:
    details = load_source_snapshot()[["ISBN", *EDITORIAL_DETAIL_COLUMNS]].copy()
    if details.empty:
        return details
    return (
        details.sort_values("ISBN")
        .groupby("ISBN", as_index=False)
        .agg({column: last_non_blank for column in EDITORIAL_DETAIL_COLUMNS})
    )


def enrich_variation_details(variations: pd.DataFrame) -> pd.DataFrame:
    if variations.empty:
        return variations[variation_columns()]
    details = source_editorial_details()
    if details.empty:
        return variations[variation_columns()]

    enriched = variations.merge(details, on="ISBN", how="left", suffixes=("", "_Source"))
    for column in EDITORIAL_DETAIL_COLUMNS:
        source_column = f"{column}_Source"
        if source_column not in enriched.columns:
            continue
        current = enriched[column].astype("string").fillna("").str.strip()
        enriched[column] = enriched[column].where(current.ne(""), enriched[source_column])
        enriched = enriched.drop(columns=source_column)
    return enriched[variation_columns()]


def week_sheet_name(cache_date: pd.Timestamp) -> str:
    return f"week_{cache_date:%m_%d_%Y}"


def write_report_table(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    df.to_excel(writer, sheet_name=sheet_name, index=False)


def apply_report_table_format(
    worksheet,
    df: pd.DataFrame,
    header_format,
    isbn_group_formats: tuple[object, object],
) -> None:
    for col_idx, column_name in enumerate(df.columns):
        worksheet.write(0, col_idx, column_name, header_format)

    if df.empty or "ISBN" not in df.columns:
        return

    last_col = len(df.columns) - 1
    group_index = -1
    previous_isbn = object()
    for row_idx, isbn in enumerate(df["ISBN"].astype(str), start=1):
        if isbn != previous_isbn:
            group_index += 1
            previous_isbn = isbn
        worksheet.conditional_format(
            row_idx,
            0,
            row_idx,
            last_col,
            {
                "type": "formula",
                "criteria": "=TRUE",
                "format": isbn_group_formats[group_index % len(isbn_group_formats)],
            },
        )


def set_variation_sheet_columns(worksheet, date_format) -> None:
    widths = {
        "A:A": (15, None),
        "B:B": (42, None),
        "C:C": (14, None),
        "D:D": (10, None),
        "E:E": (16, None),
        "F:F": (28, None),
        "G:H": (18, date_format),
        "I:J": (18, date_format),
        "K:K": (22, None),
        "L:L": (20, None),
        "M:M": (20, None),
        "N:N": (18, None),
        "O:O": (18, None),
        "P:P": (18, None),
    }
    for column_range, (width, cell_format) in widths.items():
        worksheet.set_column(column_range, width, cell_format)


def build_current_snapshot(cache: pd.DataFrame) -> pd.DataFrame:
    if cache.empty:
        return pd.DataFrame(columns=CACHE_COLUMNS)
    latest_date = cache["CacheDate"].max()
    return cache[cache["CacheDate"].eq(latest_date)].copy()


def save_variation_report(output_file: Path | None = None) -> Path:
    cache = load_cache()
    output_file = output_file or default_report_file()
    output_file.parent.mkdir(parents=True, exist_ok=True)

    all_variations = enrich_variation_details(build_variation_rows(cache))
    cache_date_pairs = recent_cache_date_pairs(cache)
    variations = variations_for_cache_pairs(all_variations, cache_date_pairs)
    weekly_variations = [
        (current_date, variations_for_cache_pairs(all_variations, [(previous_date, current_date)]))
        for previous_date, current_date in reversed(cache_date_pairs)
    ]
    criteria = build_stipulations()

    with pd.ExcelWriter(output_file, engine="xlsxwriter", datetime_format="m/d/yyyy", date_format="m/d/yyyy") as writer:
        write_report_table(writer, variations, "Variations")
        for current_date, weekly_df in weekly_variations:
            write_report_table(writer, weekly_df, week_sheet_name(current_date))
        write_report_table(writer, criteria, "Criteria")

        workbook = writer.book
        date_format = workbook.add_format({"num_format": "m/d/yyyy"})
        header_format = workbook.add_format({"bold": True, "bg_color": HEADER_FILL_COLOR})
        isbn_group_formats = tuple(
            workbook.add_format({"bg_color": color}) for color in ISBN_GROUP_FILL_COLORS
        )
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, worksheet.dim_rowmax or 0, worksheet.dim_colmax or 0)
        if "Variations" in writer.sheets:
            worksheet = writer.sheets["Variations"]
            apply_report_table_format(worksheet, variations, header_format, isbn_group_formats)
            set_variation_sheet_columns(worksheet, date_format)
        if "Criteria" in writer.sheets:
            worksheet = writer.sheets["Criteria"]
            apply_report_table_format(worksheet, criteria, header_format, isbn_group_formats)
            worksheet.set_column("A:A", 16)
            worksheet.set_column("B:B", 28)
            worksheet.set_column("C:C", 38)
        for current_date, weekly_df in weekly_variations:
            worksheet = writer.sheets[week_sheet_name(current_date)]
            apply_report_table_format(worksheet, weekly_df, header_format, isbn_group_formats)
            set_variation_sheet_columns(worksheet, date_format)

    print(f"Saved General Editorial Data Variations report: {output_file}")
    print(f"  Cache dates:       {cache['CacheDate'].nunique() if not cache.empty else 0:,}")
    print(f"  Cache rows:        {len(cache):,}")
    print(f"  Variation rows:    {len(variations):,}")
    print(f"  Weekly tabs:       {len(weekly_variations):,}")
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
    print("The primary requestor and recipient is Beth Weber.")
    print()
    print("Source:")
    print(f"  Workbook: {process_paths.GEN_EDITORIAL_SOURCE_WORKBOOK}")
    print("  Filter:   Publisher = Chronicle")
    print(
        "  Fields:   ISBN, Title, Managing Editor, Project Editor, Designer, "
        "Prod Developer, Prod Designer, Sea, Pub Grp, Task Name, Due Date, Release Date, Price"
    )
    print()
    print("Schedule:")
    print(f"  Automatic archive/report run: {process_paths.GEN_EDITORIAL_SCHEDULE_DESCRIPTION}")
    print("  Manual archive: available from this menu when an extra Monday comparison point is needed.")
    print("  Every tracked task is compared to the prior weekly snapshot.")
    print("  On Sale Date: show only if the date changes by more than 1 week.")
    print("  1st galley due, MS to DES, Print files out: show only if the date changes by more than 2 weeks.")
    print("  Cover Due, Title memo to DES: show only if the date changes by more than 4 weeks.")
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
    print(f"  Task Name filter: {', '.join(TASK_CHANGE_THRESHOLD_WEEKS)}")


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
