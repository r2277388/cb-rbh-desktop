import getpass
import importlib.util
import os
import re
import shutil
import subprocess
import sys
import textwrap
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import Tk, filedialog

import pandas as pd

# call the PO archive manager directly
from paths import process_paths
import tools.po_archive_manager as po_archive_manager
from shared.db import fetch_data_from_db, get_connection
from cross_gap.main import run_cross_gap_menu


def get_full_name():
    USER_NAMES = {
        "kbs": "Kate Breiting Schmitz",
        "mjk": "Marlena Kwasnik",
        "sdm": "Sam Mariucci",
        "RBH": "Barrett Hooper",
    }
    username = getpass.getuser()
    return USER_NAMES.get(username, username)


def greet_user():
    current_datetime = datetime.now()
    current_hour = current_datetime.hour
    current_day = current_datetime.strftime("%A")

    if current_hour < 12:
        greeting = "Good morning"
    elif 12 <= current_hour < 17:
        greeting = "Good afternoon"
    else:
        greeting = "Good evening"

    full_name = get_full_name()
    return f"\n{greeting}, {full_name}! Happy {current_day}!"


def get_farewell_message():
    current_hour = datetime.now().hour
    full_name = get_full_name()

    if 12 <= current_hour < 17:
        return f"\nHave a great afternoon, {full_name}!"
    elif 17 <= current_hour < 24:
        return f"\nGood evening, {full_name}!"
    else:
        return f"\nHave a great day, {full_name}!"


def display_options():
    options = [
        "01. Amazon",
        "02. Retailer Rolling Reports",
        "03. Sales / Operational Reports",
        "04. Data & Automation Tools",
        "05. Admin / Utilities",
        "99. Exit",
    ]
    print("\nWhat would you like to run?")
    print()
    for option in options:
        print(f"    {option}")


def display_info(choice):
    info = {
        "1": "Amazon: Opens Amazon PO, PreOrders, Customer Orders, and AMS monthly campaign workflows.",
        "2": "Retailer Rolling Reports: Opens Amazon, Barnes & Noble, Bookscan, Readerlink, Target NOC, and UK rolling-report workflows.",
        "3": "Sales / Operational Reports: Opens Cross Gap, Frontlist, General Editorial variations, Hachette, Monthend, Reprint Indicator, and SSR workflows.",
        "4": "Data & Automation Tools: Opens Automation Processes, Check Table Updates, Inventory Obsolescence Manager, Power BI Reports, and XGBoost Model.",
        "5": "Admin / Utilities: Opens Desk Procedures, requirements installation, and the main venv shell.",
        "101": f"""Amazon (1) PO Archive Manager: Copies the selected Amazon Vendor Central PO CSV to:
        {process_paths.AMAZON_PO_CURRENT_FILE}
        then prints row and cost-column totals. After reviewing the totals, you can choose whether to save an
        unchanged archive copy and a totals summary in:
        {process_paths.AMAZON_PO_DATAWAREHOUSE_ARCHIVE_DIR}""",
        "102": f"""Amazon (2) PO Report: Generates a detailed report based on Amazon Purchase Orders.
        Before running, use Amazon (1) PO Archive Manager to save/archive the latest Vendor Central PO file.
        The PO Report will use the newest archived file matching:
        {process_paths.AMAZON_PO_ARCHIVE_GLOB}
        A PO Report is saved off to: {process_paths.AMAZON_PO_ROOT_FOLDER} folder""",
        "103": "Amazon (3) PreOrders: Generates a report for Amazon NYP PreOrders. Save the relevant data file to the appropriate location before running.",
        "104": "Amazon (4) Customer Orders: Generates a report for Amazon Customer Orders. Save the relevant data file to the appropriate location before running.",
        "105": "Amazon Rolling Reports Weekly Process step 1: Create SQL Sellthrough Upload (XLSX). Runs the amazon_sql_upload workflow to build the SQL upload workbook.",
        "106": "Amazon Rolling Reports Weekly Process step 2: Process Weekly Rolling Report. Builds the weekly Amazon rolling report workbooks.",
        "107": "Amazon Rolling Reports Monthly Process step 1: Add new Monthly file to Cache. Compiles monthly Amazon sales CSVs into the monthly cache parquet.",
        "108": "Amazon Rolling Reports Monthly Process step 2: Run Monthly Rolling Report. Builds the standalone monthly Amazon rolling report workbooks.",
        "209": "Readerlink Rolling Reports: Opens Readerlink cache update, cache totals, and report creation actions.",
        "109": "Amazon AMS Monthly Campaign Summary: Builds the monthly campaign summary workbook from a selected AMS CSV.",
        "94": "Check Table Updates: Runs SQL checks for table freshness and recent weeks for SSR/Amazon/Bookscan tables.",
        "95": "Install Main Venv Requirements: Runs `pip install -r requirements.txt` using the repo's main virtual environment.",
        "96": "Open Main Venv Shell: Opens a PowerShell window with the repo's main virtual environment activated.",
        "97": "Desk Procedures: Opens a menu of desk procedures and run instructions.",
        "99": "Exit: Exits the program.",
    }
    return info.get(choice, "Invalid choice. No information available.")


def normalize_menu_choice(choice: str) -> str:
    choice = choice.strip().lower()
    if choice.isdigit():
        return str(int(choice))
    return choice


def get_latest_matching_file(folder_path: str | Path, pattern: str) -> Path:
    folder = Path(folder_path)
    matches = [path for path in folder.glob(pattern) if not path.name.startswith("~$")]
    if not matches:
        raise FileNotFoundError(f"No files found in {folder} with pattern {pattern}")
    return max(matches, key=os.path.getctime)


def get_latest_matching_file_by_mtime(folder_path: str | Path, pattern: str) -> Path:
    folder = Path(folder_path)
    matches = [path for path in folder.glob(pattern) if not path.name.startswith("~$")]
    if not matches:
        raise FileNotFoundError(f"No files found in {folder} with pattern {pattern}")
    return max(matches, key=lambda path: path.stat().st_mtime)


def get_latest_ams_campaign_csv(folder_path: str | Path) -> Path:
    folder = Path(folder_path)
    matches = [path for path in folder.glob("*.csv") if not path.name.startswith("~$")]
    if not matches:
        raise FileNotFoundError(f"No CSV files found in {folder}")

    def sort_key(path: Path) -> tuple[str, float]:
        match = re.search(r"(20\d{4})", path.stem)
        return (match.group(1) if match else "", path.stat().st_mtime)

    return max(matches, key=sort_key)


def parse_week_end_from_filename(file_path: str | Path) -> datetime | None:
    filename = Path(file_path).name
    match = re.search(r"_(\d{1,2}-\d{1,2}-\d{4})\.csv$", filename)
    if not match:
        return None
    return datetime.strptime(match.group(1), "%m-%d-%Y")


def normalize_to_week_ending_saturday(value: datetime | None) -> datetime | None:
    if value is None:
        return None
    days_to_saturday = 5 - value.weekday()
    return value + timedelta(days=days_to_saturday)


def query_amazon_sellthrough_latest_week() -> datetime | None:
    query = """
    SELECT MAX(CAST([Week] AS date)) AS LatestWeek
    FROM [CBQ2].[cb].[Sellthrough_Amazon];
    """
    try:
        engine = get_connection()
        df = fetch_data_from_db(engine, query)
    except Exception:
        return None

    if df.empty or df.iloc[0, 0] is None:
        return None

    return pd.to_datetime(df.iloc[0, 0]).to_pydatetime()


def confirm_amazon_preorders_files() -> bool:
    catalog_file = get_latest_matching_file_by_mtime(
        process_paths.ATELIER_AMAZON_CATALOG_FOLDER, "*Catalog*csv"
    )
    inventory_file = get_latest_matching_file_by_mtime(
        process_paths.ATELIER_AMAZON_INVENTORY_FOLDER, "*Inventory*csv"
    )

    while True:
        print()
        print("Amazon PreOrders will use these files:")
        print(f"  Catalog:   {catalog_file}")
        print(f"  Inventory: {inventory_file}")
        print("  Note:      Extra inventory columns are ignored; only the required fields are read.")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_customer_orders_files() -> bool:
    script_path = process_paths.AMAZON_CUSTOMER_ORDERS_SCRIPT
    weekly_sales_file = get_latest_matching_file_by_mtime(
        process_paths.ATELIER_AMAZON_BASE_FOLDER / "Sales", "*Sales*Weekly*csv"
    )
    catalog_file = get_latest_matching_file_by_mtime(
        process_paths.ATELIER_AMAZON_CATALOG_FOLDER, "*Catalog*csv"
    )
    inventory_file = get_latest_matching_file_by_mtime(
        process_paths.ATELIER_AMAZON_INVENTORY_FOLDER, "*inventory*csv"
    )
    traffic_file = get_latest_matching_file_by_mtime(
        process_paths.ATELIER_AMAZON_BASE_FOLDER / "Traffic", "*Traffic*csv"
    )
    ypticod_file = process_paths.ORACLE_YPTICOD_FILE
    output_file = process_paths.AMAZON_CUSTOMER_ORDERS_OUTPUT_FILE

    if not ypticod_file.exists():
        raise FileNotFoundError(f"Required file not found: {ypticod_file}")

    while True:
        print()
        print("Amazon Customer Orders will use these files:")
        print(f"  Script:            {script_path}")
        print(f"  Weekly Sales:      {weekly_sales_file}")
        print(f"  Catalog:           {catalog_file}")
        print(f"  Inventory:         {inventory_file}")
        print(f"  Traffic:           {traffic_file}")
        print(f"  Oracle YPTICOD:    {ypticod_file}")
        print(f"  Output workbook:   {output_file}")
        print("  Note:              Extra file columns are ignored when required fields are present.")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_sql_upload_files() -> bool:
    script_path = process_paths.AMAZON_SQL_UPLOAD_SCRIPT
    manual_key_path = process_paths.AMAZON_SQL_UPLOAD_MANUAL_KEY_FILE
    removal_list_path = process_paths.AMAZON_SQL_UPLOAD_REMOVAL_LIST_FILE

    sales_file = get_latest_matching_file(
        process_paths.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["sales"], "*.csv"
    )
    inventory_file = get_latest_matching_file(
        process_paths.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["inventory"], "*.csv"
    )
    traffic_file = get_latest_matching_file(
        process_paths.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["traffic"], "*.csv"
    )
    catalog_file = get_latest_matching_file(
        process_paths.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["catalog"], "*.csv"
    )
    ypticod_file = process_paths.ORACLE_YPTICOD_FILE
    sales_week_end = parse_week_end_from_filename(sales_file)
    sql_latest_week_raw = query_amazon_sellthrough_latest_week()
    sql_latest_week_normalized = normalize_to_week_ending_saturday(sql_latest_week_raw)

    if not ypticod_file.exists():
        raise FileNotFoundError(f"Required file not found: {ypticod_file}")

    while True:
        print()
        print("amazon_sql_upload will use these files:")
        print(f"  Script:            {script_path}")
        print(f"  Sales CSV:         {sales_file}")
        print(f"  Inventory CSV:     {inventory_file}")
        print(f"  Traffic CSV:       {traffic_file}")
        print(f"  Catalog CSV:       {catalog_file}")
        print(f"  Oracle YPTICOD:    {ypticod_file}")
        print(f"  Manual key file:   {manual_key_path}")
        print(f"  Removal list file: {removal_list_path}")
        print("  SQL source:        sql-2-db / CBQ2 (EBS item ISBN key)")
        print(
            "  Sales file week:   "
            + (sales_week_end.strftime("%A, %Y-%m-%d") if sales_week_end else "Could not parse from filename")
        )
        print(
            "  Current SQL max:   "
            + (
                sql_latest_week_raw.strftime("%A, %Y-%m-%d")
                if sql_latest_week_raw
                else "Unavailable from this session"
            )
        )
        if sql_latest_week_normalized and (
            not sql_latest_week_raw or sql_latest_week_normalized.date() != sql_latest_week_raw.date()
        ):
            print(
                "  SQL week (Sat):    "
                + sql_latest_week_normalized.strftime("%A, %Y-%m-%d")
            )
        print(
            f"  Default output:    {process_paths.amazon_sql_upload_output_file(sales_week_end)}"
        )
        print(
            f"  Output folder:     {process_paths.AMAZON_SQL_UPLOAD_OUTPUT_DIR}"
        )
        print(
            f"  Weekly summary copy: {process_paths.amazon_sql_upload_weekly_summary_file(sales_week_end)}"
        )
        print(
            f"  Weekly workbook:   {process_paths.AMAZON_WEEKLY_REPORTS_DIR}\\w##_yyyy_mm_dd.xlsx"
        )
        print("  Weekly cbq upload workbook:   Auto-created from the cleaned six-column dataset")
        print("  Save dialog:       Opens with this default file prefilled")
        print()
        print("    1. Continue to week-ending date")
        print("    2. Return to main menu")
        print()
        preview_choice = input("Choose an option: ").strip().lower()
        print()

        if preview_choice in {"2", "b", "back", "return", "menu"}:
            return False

        if preview_choice not in {"1", "c", "continue"}:
            print("Invalid choice. Please select a valid option.")
            continue

        expected_text = input(
            "Expected week-ending Saturday (MM/DD/YYYY or blank to use the sales file date): "
        ).strip()
        print()

        if expected_text:
            try:
                expected_week_end = datetime.strptime(expected_text, "%m/%d/%Y")
            except ValueError:
                print("Invalid date. Use MM/DD/YYYY, for example 04/04/2026.")
                continue
        else:
            expected_week_end = sales_week_end

        if expected_week_end is None:
            print("Could not determine the expected week-ending date from the sales file. Please enter it explicitly.")
            continue

        print(f"  Expected week:     {expected_week_end.strftime('%A, %Y-%m-%d')}")
        prior_expected_week_end = expected_week_end - timedelta(days=7)
        print(f"  Prior SQL week:    {prior_expected_week_end.strftime('%A, %Y-%m-%d')}")
        if expected_week_end.weekday() != 5:
            print("  Warning:           The expected week-ending date is not a Saturday.")
        if sales_week_end and sales_week_end.date() != expected_week_end.date():
            print(
                "  Warning:           The sales filename does not match the expected week-ending date."
            )
        if sql_latest_week_raw and sql_latest_week_raw.weekday() != 5:
            print(
                "  Warning:           The current max [Week] in [CBQ2].[cb].[Sellthrough_Amazon] is not a Saturday."
            )
        if (
            sql_latest_week_normalized
            and sql_latest_week_normalized.date() != prior_expected_week_end.date()
        ):
            print(
                "  Warning:           The normalized current SQL week does not match the prior expected loaded week."
            )
        if (
            sql_latest_week_normalized
            and sql_latest_week_raw
            and sql_latest_week_normalized.date() != sql_latest_week_raw.date()
        ):
            print(
                "  Note:              SQL contains a non-Saturday [Week] value, but it normalizes to the Saturday shown above."
            )
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_rolling_reports_check_files() -> bool:
    script_path = process_paths.AMAZON_ROLLING_CHECK_SCRIPT
    sql_file = process_paths.AMAZON_ROLLING_SQL_FILE

    while True:
        print()
        print("Amazon Rolling Reports check will use these files:")
        print(f"  Script:            {script_path}")
        print(f"  SQL file:          {sql_file}")
        print("  SQL source:        sql-2-db / CBQ2")
        print("  Table checked:     [CBQ2].[cb].[Sellthrough_Amazon]")
        print()
        print("    1. Continue")
        print("    2. Return to previous menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_rolling_reports_run_files() -> bool:
    script_path = process_paths.AMAZON_ROLLING_REPORTS_SCRIPT
    latest_po_file = get_latest_matching_file(process_paths.AMAZON_PO_FOLDER, "*.xlsx")
    customer_orders_pickle = process_paths.AMAZON_ROLLING_CUSTOMER_ORDERS_PICKLE
    units_shipped_pickle = process_paths.AMAZON_ROLLING_UNITS_SHIPPED_PICKLE
    po_pickle = process_paths.AMAZON_ROLLING_PO_PICKLE
    customer_orders_sql = process_paths.AMAZON_ROLLING_CUSTOMER_ORDERS_QUERY
    units_shipped_sql = process_paths.AMAZON_ROLLING_UNITS_SHIPPED_QUERY
    date_check_sql = process_paths.AMAZON_ROLLING_DATE_CHECK_QUERY
    output_folder = process_paths.AMAZON_ROLLING_OUTPUT_FOLDER

    while True:
        print()
        print("Amazon Rolling Reports will use these files:")
        print(f"  Script:                  {script_path}")
        print(f"  Latest Amazon PO file:   {latest_po_file}")
        print(f"  PO pickle:               {po_pickle}")
        print(f"  Customer Orders pickle:  {customer_orders_pickle}")
        print(f"  Units Shipped pickle:    {units_shipped_pickle}")
        print(f"  Customer Orders query:   {customer_orders_sql}")
        print(f"  Units Shipped query:     {units_shipped_sql}")
        print(f"  Date check query:        {date_check_sql}")
        print("  SQL source:              sql-2-db / CBQ2 (Faire only)")
        print(f"  Main output folder:      {output_folder}")
        print("  DP output folders:       Configured in amazon_rolling_reports/paths.py")
        print()
        print("    1. Continue")
        print("    2. Return to previous menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_monthly_sales_files() -> bool:
    script_path = process_paths.AMAZON_MONTHLY_SALES_SCRIPT
    source_root = process_paths.AMAZON_MONTHLY_SALES_ROOT
    fallback_root = process_paths.AMAZON_MONTHLY_SALES_FALLBACK_ROOT
    active_root = source_root if source_root.exists() else fallback_root
    output_file = active_root / "cache" / process_paths.AMAZON_MONTHLY_SALES_PARQUET_NAME

    if not active_root.exists():
        raise FileNotFoundError(f"Monthly sales folder not found: {source_root} or {fallback_root}")

    while True:
        print()
        print("Amazon Monthly Sales compile will use these files:")
        print(f"  Script:             {script_path}")
        print(f"  Source folder:      {active_root}")
        print(f"  Output parquet:     {output_file}")
        print(f"  Oracle YPTICOD:     {process_paths.ORACLE_YPTICOD_FILE}")
        print(f"  Latest catalog:     {process_paths.ATELIER_AMAZON_CATALOG_FOLDER}")
        print("  SQL source:         sql-2-db / CBQ2 (EBS item ISBN key)")
        print()
        print("    1. Continue")
        print("    2. Return to previous menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def load_module_from_path(module_name: str, file_path: Path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load module from {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def interpreter_has_module(python_executable: str, module_name: str) -> bool:
    try:
        result = subprocess.run(
            [
                python_executable,
                "-c",
                (
                    "import importlib.util, sys; "
                    f"sys.exit(0 if importlib.util.find_spec('{module_name}') else 1)"
                ),
            ],
            capture_output=True,
            check=False,
        )
    except OSError:
        return False

    return result.returncode == 0


def get_excel_automation_python() -> str:
    candidates: list[str] = []

    if sys.executable:
        candidates.append(sys.executable)

    path_python = shutil.which("python")
    if path_python:
        candidates.append(path_python)

    candidates.append(r"C:\Python310\python.exe")

    seen: set[str] = set()
    for candidate in candidates:
        normalized = os.path.normcase(os.path.abspath(candidate))
        if normalized in seen:
            continue
        seen.add(normalized)
        if interpreter_has_module(candidate, "pythoncom"):
            return candidate

    raise RuntimeError(
        "Could not find a Python interpreter with pywin32/pythoncom for Excel automation."
    )


def parse_schtasks_list_output(output: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for raw_line in output.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, value = line.split(":", 1)
        values[key.strip()] = value.strip()
    return values


def get_title_lookup_task_details() -> dict[str, str] | None:
    result = subprocess.run(
        [
            "schtasks",
            "/Query",
            "/TN",
            process_paths.TITLE_LOOKUP_TASK_NAME,
            "/FO",
            "LIST",
            "/V",
        ],
        capture_output=True,
        check=False,
        text=True,
    )
    if result.returncode != 0:
        return None
    return parse_schtasks_list_output(result.stdout)


def run_title_lookup_refresh_now() -> None:
    print("Running the weekly Title Lookup force refresh... Please wait.")
    subprocess.run(
        [
            get_excel_automation_python(),
            str(process_paths.EXCEL_REFRESH_SCRIPT),
            str(process_paths.TITLE_LOOKUP_WORKBOOK),
            "--connection",
            process_paths.TITLE_LOOKUP_CONNECTION_NAME,
            "--table",
            process_paths.TITLE_LOOKUP_TABLE_NAME,
        ],
        check=True,
    )
    print("The weekly Title Lookup workbook is now refreshed.")


def enable_title_lookup_weekly_automation() -> None:
    print("Creating or updating the weekly Title Lookup scheduled task... Please wait.")
    subprocess.run(
        ["venv/Scripts/python", str(process_paths.repo_path("tools", "title_lookup_automation.py")), "register"],
        check=True,
    )
    print("Weekly Title Lookup automation is configured.")


def show_title_lookup_automation_details() -> None:
    subprocess.run(
        ["venv/Scripts/python", str(process_paths.repo_path("tools", "title_lookup_automation.py")), "status"],
        check=True,
    )


def disable_title_lookup_weekly_automation() -> None:
    print("Disabling the weekly Title Lookup scheduled task... Please wait.")
    subprocess.run(
        ["venv/Scripts/python", str(process_paths.repo_path("tools", "title_lookup_automation.py")), "disable"],
        check=True,
    )
    print("Weekly Title Lookup automation is now disabled.")


def get_cross_gap_task_details() -> dict[str, str] | None:
    result = subprocess.run(
        [
            "schtasks",
            "/Query",
            "/TN",
            process_paths.CROSS_GAP_TASK_NAME,
            "/FO",
            "LIST",
            "/V",
        ],
        capture_output=True,
        check=False,
        text=True,
    )
    if result.returncode != 0:
        return None
    return parse_schtasks_list_output(result.stdout)


def run_cross_gap_report_now() -> None:
    print("Running the Cross Gap report... Please wait.")
    subprocess.run(
        ["venv/Scripts/python", str(process_paths.CROSS_GAP_SCRIPT), "run"],
        check=True,
    )
    print("The Cross Gap report is now ready.")


def enable_cross_gap_weekly_automation() -> None:
    print("Creating or updating the weekly Cross Gap scheduled task... Please wait.")
    subprocess.run(
        ["venv/Scripts/python", str(process_paths.repo_path("tools", "cross_gap_automation.py")), "register"],
        check=True,
    )
    print("Weekly Cross Gap automation is configured.")


def show_cross_gap_automation_details() -> None:
    subprocess.run(
        ["venv/Scripts/python", str(process_paths.repo_path("tools", "cross_gap_automation.py")), "status"],
        check=True,
    )


def disable_cross_gap_weekly_automation() -> None:
    print("Disabling the weekly Cross Gap scheduled task... Please wait.")
    subprocess.run(
        ["venv/Scripts/python", str(process_paths.repo_path("tools", "cross_gap_automation.py")), "disable"],
        check=True,
    )
    print("Weekly Cross Gap automation is now disabled.")


def run_title_lookup_refresh_menu() -> None:
    workbook_path = process_paths.TITLE_LOOKUP_WORKBOOK

    if not workbook_path.exists():
        raise FileNotFoundError(f"Required workbook not found: {workbook_path}")

    while True:
        task_details = get_title_lookup_task_details()
        print()
        print("Weekly Title Lookup Refresh")
        print(f"  Workbook:          {workbook_path}")
        print(f"  Connection:        {process_paths.TITLE_LOOKUP_CONNECTION_NAME}")
        print(f"  Output table:      {process_paths.TITLE_LOOKUP_TABLE_NAME}")
        print("  Manual action:     Force refresh the workbook right now.")
        print("  Automation type:   Windows Task Scheduler")
        print(f"  Task name:         {process_paths.TITLE_LOOKUP_TASK_NAME}")
        print(f"  Task location:     {process_paths.TITLE_LOOKUP_TASK_LOCATION}")
        print(f"  Default schedule:  {process_paths.TITLE_LOOKUP_SCHEDULE_DESCRIPTION}")
        if task_details is None:
            print("  Current status:    Not currently scheduled")
            print("  Next run:          Not available")
        else:
            print("  Current status:    Scheduled")
            print(f"  Next run:          {task_details.get('Next Run Time', 'Unknown')}")
            print(f"  Last run:          {task_details.get('Last Run Time', 'Unknown')}")
            print(f"  Last result:       {task_details.get('Last Result', 'Unknown')}")
        print("  Notes:             This automation refreshes the workbook in place.")
        print(
            textwrap.fill(
                "                     Excel automation is most reliable when you are logged into Windows, because the scheduled task opens Excel to run the refresh.",
                width=100,
                subsequent_indent="                     ",
            )
        )
        print()
        print("    1. Force Refresh Now")
        print("    2. Enable or Update Weekly Automation")
        print("    3. Disable Weekly Automation")
        print("    4. Show Automation Details")
        print("    5. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "force", "refresh", "run"}:
            try:
                run_title_lookup_refresh_now()
            except (subprocess.CalledProcessError, RuntimeError) as e:
                print(f"An error occurred while refreshing the weekly Title Lookup workbook: {e}")
            continue

        if choice in {"2", "enable", "schedule", "automation"}:
            try:
                enable_title_lookup_weekly_automation()
            except subprocess.CalledProcessError as e:
                print(f"An error occurred while configuring weekly automation: {e}")
            continue

        if choice in {"3", "disable", "off"}:
            try:
                disable_title_lookup_weekly_automation()
            except subprocess.CalledProcessError as e:
                print(f"An error occurred while disabling weekly automation: {e}")
            continue

        if choice in {"4", "details", "status", "info"}:
            try:
                show_title_lookup_automation_details()
            except subprocess.CalledProcessError as e:
                print(f"Unable to show automation details: {e}")
            input("\nPress Enter to return to the Title Lookup menu...")
            continue

        if choice in {"5", "b", "back", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def run_cross_gap_automation_menu() -> None:
    while True:
        task_details = get_cross_gap_task_details()
        print()
        print("Weekly Cross Gap Report")
        print(f"  Script:            {process_paths.CROSS_GAP_SCRIPT}")
        print(f"  Output folder:     {process_paths.CROSS_GAP_OUTPUT_DIR}")
        print("  Automation type:   Windows Task Scheduler")
        print(f"  Task name:         {process_paths.CROSS_GAP_TASK_NAME}")
        print(f"  Task location:     {process_paths.CROSS_GAP_TASK_LOCATION}")
        print(f"  Default schedule:  {process_paths.CROSS_GAP_SCHEDULE_DESCRIPTION}")
        if task_details is None:
            print("  Current status:    Not currently scheduled")
            print("  Next run:          Not available")
        else:
            print("  Current status:    Scheduled")
            print(f"  Next run:          {task_details.get('Next Run Time', 'Unknown')}")
            print(f"  Last run:          {task_details.get('Last Run Time', 'Unknown')}")
            print(f"  Last result:       {task_details.get('Last Result', 'Unknown')}")
        print("  Notes:             This automation writes a dated Cross Gap workbook to the shared reports folder.")
        print()
        print("    1. Run Cross Gap Now")
        print("    2. Enable or Update Weekly Automation")
        print("    3. Disable Weekly Automation")
        print("    4. Show Automation Details")
        print("    5. Return to Automation Processes")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "run", "now"}:
            try:
                run_cross_gap_report_now()
            except subprocess.CalledProcessError as e:
                print(f"An error occurred while running the Cross Gap report: {e}")
            continue

        if choice in {"2", "enable", "schedule", "automation"}:
            try:
                enable_cross_gap_weekly_automation()
            except subprocess.CalledProcessError as e:
                print(f"An error occurred while configuring weekly Cross Gap automation: {e}")
            continue

        if choice in {"3", "disable", "off"}:
            try:
                disable_cross_gap_weekly_automation()
            except subprocess.CalledProcessError as e:
                print(f"An error occurred while disabling weekly Cross Gap automation: {e}")
            continue

        if choice in {"4", "details", "status", "info"}:
            try:
                show_cross_gap_automation_details()
            except subprocess.CalledProcessError as e:
                print(f"Unable to show Cross Gap automation details: {e}")
            input("\nPress Enter to return to the Cross Gap automation menu...")
            continue

        if choice in {"5", "b", "back", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def run_automation_processes_menu() -> None:
    while True:
        print("\nAutomation Processes")
        print()
        print("    1. Title Lookup Refresh (weekly)")
        print("    2. Cross Gap Report (weekly)")
        print("    3. General Editorial Data Variations (weekly Monday)")
        print("    4. Back to main menu")
        print()
        try:
            subchoice = input("Choose an option: ").strip().lower()
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if subchoice == "1":
            try:
                run_title_lookup_refresh_menu()
            except FileNotFoundError as e:
                print(f"Unable to locate the Title Lookup workbook: {e}")
            continue

        if subchoice == "2":
            run_cross_gap_automation_menu()
            continue

        if subchoice == "3":
            run_python_process(
                "General Editorial Data Variations Automation Status",
                process_paths.repo_path("tools", "gen_editorial_automation.py"),
                extra_args=["status"],
            )
            input("\nPress Enter to return to Automation Processes...")
            continue

        if subchoice in {"4", "b", "back", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def confirm_amazon_ams_files() -> bool:
    manager_script = process_paths.AMAZON_AMS_MANAGER_SCRIPT
    process_script = process_paths.AMAZON_AMS_PROCESS_SCRIPT
    mapping_file = process_paths.CHRONICLE_ASIN_MAPPING_FILE
    campaign_folder = process_paths.AMAZON_AMS_MONTHLY_CAMPAIGN_FOLDER
    latest_campaign_file = get_latest_ams_campaign_csv(campaign_folder)

    while True:
        print()
        print("Amazon AMS Manager will use these files:")
        print(f"  Manager script:         {manager_script}")
        print(f"  Report script:          {process_script}")
        print(f"  Campaign folder:        {campaign_folder}")
        print(f"  Latest CSV:             {latest_campaign_file}")
        print(f"  ASIN mapping file:      {mapping_file}")
        print(f"  Oracle YPTICOD:         {process_paths.ORACLE_YPTICOD_FILE}")
        print("  SQL source:             sql-2-db / CBQ2 (AMS item metadata)")
        print(f"  History parquet:        {process_paths.AMAZON_AMS_MONTHLY_CAMPAIGN_HISTORY_PARQUET}")
        print(f"  Final report folder:    {process_paths.AMAZON_AMS_FINAL_REPORTS_FOLDER}")
        print("  Output workbooks:       yyyymm_AMS_Performance_by_ASIN_ALL/PWP.xlsx")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def confirm_frontlist_supercharged_files() -> bool:
    script_path = process_paths.FRONTLIST_SUPERCHARGED_SCRIPT
    output_dir = process_paths.FRONTLIST_SUPERCHARGED_OUTPUT_DIR
    from FLTracking_Supercharged.processes.amazon_sellthrough import (
        get_amazon_sellthrough_source_metadata,
    )
    from FLTracking_Supercharged.processes.barnes_noble_weekly import (
        get_barnes_noble_source_metadata,
    )
    from FLTracking_Supercharged.processes.frontlist_main import (
        resolve_frontlist_tracking_path,
    )

    frontlist_file = resolve_frontlist_tracking_path(process_paths.FRONTLIST_TRACKING_FOLDER)
    ingram_file = get_latest_matching_file(
        process_paths.INGRAM_DAILY_REPORT_FOLDER, "Daily Report*.xlsx"
    )
    inventory_file = next(
        process_paths.INVENTORY_DAILY_FINANCE_ONLY_FOLDER.glob("Inventory*.xlsx"), None
    )
    amazon_preorders_file = process_paths.CURRENT_AMAZON_PREORDERS_FILE
    faire_qty_sql = process_paths.FRONTLIST_FAIRE_QTY_SQL
    faire_orders_sql = process_paths.FRONTLIST_FAIRE_ORDERS_SQL
    barnes_noble_metadata = get_barnes_noble_source_metadata()
    amazon_sellthrough_metadata = get_amazon_sellthrough_source_metadata()

    if inventory_file is None:
        raise FileNotFoundError(
            f"No files found in {process_paths.INVENTORY_DAILY_FINANCE_ONLY_FOLDER} "
            "with pattern Inventory*.xlsx"
        )
    if not amazon_preorders_file.exists():
        raise FileNotFoundError(f"Required file not found: {amazon_preorders_file}")

    while True:
        print()
        print("Frontlist Supercharged Data will use these files:")
        print(f"  Script:                  {script_path}")
        print(f"  Frontlist Tracking:      {frontlist_file}")
        print(f"  Inventory Detail:        {inventory_file}")
        print(f"  Amazon Preorders:        {amazon_preorders_file}")
        print(f"  Ingram Daily Report:     {ingram_file}")
        print_frontlist_barnes_noble_source(barnes_noble_metadata)
        print_frontlist_amazon_sellthrough_source(amazon_sellthrough_metadata)
        print(f"  Faire Qty SQL:           {faire_qty_sql}")
        print(f"  Faire Orders SQL:        {faire_orders_sql}")
        print("  SQL source:              sql-2-db / CBQ2")
        print(f"  Cache/output folder:     {output_dir}")
        print("  Final workbook:          Chosen in save dialog, defaulting to FLTracking_Supercharged/output")
        print()
        print("    1. Continue")
        print("    2. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return True

        if choice in {"2", "b", "back", "return", "menu"}:
            return False

        print("Invalid choice. Please select a valid option.")


def print_frontlist_barnes_noble_source(metadata: dict[str, object]) -> None:
    if metadata["source_type"] == "parquet":
        print(
            "  B&N Sales Parquet:       "
            f"{metadata['sales_path']} "
            f"(updated {metadata['sales_modified_date']})"
        )
        print(
            "  B&N Inventory Parquet:   "
            f"{metadata['inventory_path']} "
            f"(updated {metadata['inventory_modified_date']})"
        )
    else:
        print(
            "  B&N Weekly Excel Fallback:"
            f" {metadata['source_path']} "
            f"(updated {metadata['modified_date']})"
        )
        print(f"  B&N fallback reason:     {metadata['fallback_reason']}")


def print_frontlist_amazon_sellthrough_source(metadata: dict[str, object]) -> None:
    if metadata["source_type"] == "cache":
        print(
            "  Amazon Sellthrough Cache:"
            f" {metadata['cache_path']} "
            f"(updated {metadata['modified_date']})"
        )
    else:
        print(
            "  Amazon SQL Fallback:     "
            f"{metadata['sql_path']} "
            f"(updated {metadata['modified_date']})"
        )
        print(f"  Amazon fallback reason:  {metadata['fallback_reason']}")


def choose_bn_raw_folder_with_window(initial_dir: Path) -> Path | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askdirectory(
            initialdir=str(initial_dir),
            title="Choose the Barnes & Noble raw files folder",
        )
    finally:
        root.destroy()

    if not selected:
        return None

    path = Path(selected)
    if not path.exists():
        raise FileNotFoundError(f"Raw folder not found: {path}")
    if not path.is_dir():
        raise NotADirectoryError(f"Raw folder path is not a directory: {path}")
    return path


def confirm_bn_rolling_reports_files() -> Path | None:
    script_path = process_paths.BN_ROLLING_REPORTS_SCRIPT
    raw_base_folder = process_paths.BN_RAW_BASE_FOLDER
    raw_folders = sorted(
        [
            path
            for path in raw_base_folder.iterdir()
            if path.is_dir() and path.name.lower().endswith("_raw_files")
        ],
        key=lambda path: path.name,
    )
    if not raw_folders:
        raise FileNotFoundError(
            f"No Barnes & Noble raw folders were found under {raw_base_folder}"
        )

    selected_raw_folder = raw_folders[-1]

    def format_modified(path: Path) -> str:
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %I:%M:%S %p")

    while True:
        pos_csvs = sorted(
            [path for path in selected_raw_folder.iterdir() if path.is_file() and path.suffix.lower() == ".csv"],
            key=lambda path: path.name.lower(),
        )
        sales_files = sorted(
            [
                path
                for path in selected_raw_folder.iterdir()
                if path.is_file()
                and path.suffix.lower() == ".xlsx"
                and path.name.lower().startswith("sales")
                and not path.name.startswith("~$")
            ],
            key=lambda path: path.name.lower(),
        )
        inventory_files = sorted(
            [
                path
                for path in selected_raw_folder.iterdir()
                if path.is_file()
                and path.suffix.lower() == ".xlsx"
                and path.name.lower().startswith("inventory")
                and not path.name.startswith("~$")
            ],
            key=lambda path: path.name.lower(),
        )

        print()
        print("Barnes & Noble Rolling Reports will use these files:")
        print(f"  Script:              {script_path}")
        print(f"  Raw base folder:     {raw_base_folder}")
        print(f"  Default raw folder:  {selected_raw_folder}")
        print("  POS CSV files:")
        if pos_csvs:
            for file_path in pos_csvs:
                print(f"    {file_path}")
        else:
            print("    None found")
        if sales_files:
            print("  Sales workbook(s):")
            for file_path in sales_files:
                print(f"    {file_path}")
                print(f"      Last modified: {format_modified(file_path)}")
        else:
            print("  Sales workbook(s):   None found in default raw folder")
        if inventory_files:
            print("  Inventory workbook(s):")
            for file_path in inventory_files:
                print(f"    {file_path}")
                print(f"      Last modified: {format_modified(file_path)}")
        else:
            print("  Inventory workbook(s): None found in default raw folder")
        print("  Output files:        Saved in the selected raw folder during submenu actions")
        print()
        print("    1. Continue")
        print("    2. Change raw folder")
        print("    3. Return to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice in {"1", "c", "continue"}:
            return selected_raw_folder

        if choice in {"2", "change", "folder", "browse"}:
            updated_folder = choose_bn_raw_folder_with_window(raw_base_folder)
            if updated_folder is None:
                print("No folder was selected.")
                continue
            selected_raw_folder = updated_folder
            continue

        if choice in {"3", "b", "back", "return", "menu"}:
            return None

        print("Invalid choice. Please select a valid option.")


def run_program(choice):
    if choice == "1":
        run_amazon_menu()
        return

    if choice == "2":
        run_retailer_rolling_reports_menu()
        return

    if choice == "3":
        run_sales_operational_reports_menu()
        return

    if choice == "4":
        run_data_automation_tools_menu()
        return

    if choice == "5":
        run_admin_utilities_menu()
        return

    if choice == "99":
        print(get_farewell_message())
        return

    print("Invalid choice. Please select a valid option.")


def run_python_process(
    report_name: str,
    script_path: str | Path,
    *,
    python_executable: str | Path = "venv/Scripts/python",
    extra_args: list[str] | None = None,
    skipped_returncodes: set[int] | None = None,
) -> bool:
    print(f"Running the {report_name}... Please wait.")
    command = [str(python_executable), str(script_path)]
    if extra_args:
        command.extend(extra_args)
    try:
        subprocess.run(command, check=True, cwd=process_paths.REPO_ROOT)
        print(f"The {report_name} is now ready.")
        return True
    except subprocess.CalledProcessError as exc:
        if skipped_returncodes and exc.returncode in skipped_returncodes:
            return False
        print(f"An error occurred while running {script_path}.")
        return False


def confirm_refresh_all_amazon_rolling_caches() -> bool:
    while True:
        choice = input("Do you want to update all Amazon Rolling caches now? (y/n): ").strip().lower()
        if choice in {"y", "yes"}:
            return True
        if choice in {"n", "no"}:
            return False
        print("Invalid choice. Please enter y or n.")


def run_readerlink_rolling_reports() -> None:
    while True:
        print("")
        print("Readerlink Rolling Reports")
        print("    1. Add a new week's data to cache")
        print("    2. Show totals for the last 4 cached Readerlink weeks")
        print("    3. Create the Readerlink Rolling Report")
        print("    0. Back")
        choice = input("Choose an action: ").strip().lower()

        if choice in {"0", "b", "back"}:
            return
        if choice == "1":
            run_python_process(
                "Readerlink Cache Update",
                process_paths.repo_path("readerlink_rolling_reports", "build_cache.py"),
                extra_args=["--add-new-week"],
                skipped_returncodes={10},
            )
            continue
        if choice == "2":
            run_python_process(
                "Readerlink Cache Totals",
                process_paths.repo_path("readerlink_rolling_reports", "build_cache.py"),
                extra_args=["--show-last-weeks"],
            )
            continue
        if choice == "3":
            run_python_process(
                "Readerlink Rolling Reports",
                process_paths.repo_path("readerlink_rolling_reports", "main.py"),
            )
            continue

        print("Invalid choice. Please select a valid option.")


def install_main_venv_requirements() -> None:
    print("Installing requirements into the main venv... Please wait.")
    try:
        subprocess.run(
            ["venv/Scripts/python", "-m", "pip", "install", "-r", "requirements.txt"],
            check=True,
        )
        print("Main venv requirements are up to date.")
    except subprocess.CalledProcessError:
        print("An error occurred while installing requirements.txt into the main venv.")


def open_main_venv_shell() -> None:
    activate_script = Path("venv") / "Scripts" / "Activate.ps1"
    if not activate_script.exists():
        print(f"Main venv activation script not found: {activate_script}")
        return
    print("Opening a PowerShell window with the main venv activated.")
    try:
        subprocess.Popen(
            [
                "powershell",
                "-NoExit",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                f"& '{activate_script.resolve()}'",
            ]
        )
    except OSError as e:
        print(f"Unable to open the main venv shell: {e}")


def run_retailer_rolling_reports_menu() -> None:
    while True:
        print("\nRetailer Rolling Reports")
        print()
        print("Amazon Rolling Reports (Weekly Process)")
        print("    01. Create SQL Sellthrough Upload (XLSX) (step 1)")
        print("    02. Process Weekly Rolling Report (step 2)")
        print()
        print("Amazon Rolling Reports (Monthly Process)")
        print("    03. Add new Monthly file to Cache (monthly step 1)")
        print("    04. Run Monthly Rolling Report (monthly step 2)")
        print()
        print("Other Retailer Rolling Reports")
        print("    05. Barnes & Noble Rolling Reports")
        print("    06. Bookscan Rolling Reports")
        print("    07. Target NOC Rolling Reports")
        print("    08. Abrams & Chronicle UK Rolling Reports")
        print("    09. Readerlink Rolling Reports")
        print()
        print("    99. Back to main menu")
        print()
        try:
            choice = normalize_menu_choice(input("Choose an option: "))
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if choice == "1":
            try:
                if not confirm_amazon_sql_upload_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Create SQL Sellthrough Upload (XLSX) source files: {e}")
                continue
            run_python_process("Amazon Create SQL Sellthrough Upload (XLSX)", "amazon_sql_upload/main.py")
            continue

        if choice == "2":
            try:
                if not confirm_amazon_rolling_reports_run_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Rolling Reports source files: {e}")
                continue
            extra_args = ["--refresh-all-caches"] if confirm_refresh_all_amazon_rolling_caches() else None
            run_python_process(
                "Amazon Weekly Rolling Reports",
                "amazon_rolling_reports/main.py",
                extra_args=extra_args,
            )
            continue

        if choice == "3":
            try:
                if not confirm_amazon_monthly_sales_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Monthly Sales files: {e}")
                continue
            run_python_process("Amazon Monthly Sales Cache", "amazon_rolling_reports/monthly_sales.py")
            continue

        if choice == "4":
            try:
                if not confirm_amazon_monthly_sales_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Monthly Sales files: {e}")
                continue
            run_python_process("Amazon Monthly Rolling Reports", "amazon_rolling_reports/monthly_rolling_reports.py")
            continue

        if choice == "5":
            try:
                selected_bn_raw_folder = confirm_bn_rolling_reports_files()
                if selected_bn_raw_folder is None:
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Barnes & Noble Rolling Reports files: {e}")
                continue
            run_python_process(
                "Barnes & Noble Rolling Reports",
                process_paths.repo_path("bn_rolling_reports", "main.py"),
                extra_args=["--default-raw-folder", str(selected_bn_raw_folder)],
            )
            continue

        if choice == "6":
            run_python_process("Bookscan Rolling Reports", process_paths.BOOKSCAN_ROLLING_REPORTS_SCRIPT)
            continue

        if choice == "7":
            run_python_process("Target NOC Rolling Reports", process_paths.repo_path("target_rolling_report", "main.py"))
            continue

        if choice == "8":
            run_python_process(
                "Abrams & Chronicle UK Rolling Reports",
                process_paths.repo_path("Abrams_Chronicle_rollling_reports", "main.py"),
            )
            continue

        if choice == "9":
            run_readerlink_rolling_reports()
            continue

        if choice in {"99", "back", "b", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def run_sales_operational_reports_menu() -> None:
    while True:
        print("\nSales / Operational Reports")
        print()
        print("    01. Cross Gap")
        print("    02. Frontlist Supercharged Data")
        print("    03. Hachette Orders - Shipping Estimates")
        print("    04. Monthend Reports")
        print("    05. Reprint Indicator Report Updater")
        print("    06. SSR Daily Summary")
        print("    07. General Editorial Data Variations")
        print("    08. Monthly Top Customers")
        print()
        print("    99. Back to main menu")
        print()
        try:
            choice = normalize_menu_choice(input("Choose an option: "))
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if choice == "1":
            run_cross_gap_menu()
            continue

        if choice == "2":
            try:
                if not confirm_frontlist_supercharged_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Frontlist Supercharged source files: {e}")
                continue
            run_python_process("Frontlist Supercharged Data", process_paths.repo_path("FLTracking_Supercharged", "main.py"))
            continue

        if choice == "3":
            run_python_process("Hachette Orders - Shipping Estimates", process_paths.repo_path("hachette_orders", "main.py"))
            continue

        if choice == "4":
            run_python_process("Monthend Reports", process_paths.repo_path("monthend", "main.py"))
            continue

        if choice == "5":
            run_python_process(
                "Reprint Indicator Report Updater",
                process_paths.REPRINT_INDICATOR_AUTOMATION_SCRIPT,
                python_executable=get_excel_automation_python(),
            )
            continue

        if choice == "6":
            run_ssr_daily_summary_menu()
            continue

        if choice == "7":
            run_python_process(
                "General Editorial Data Variations",
                process_paths.GEN_EDITORIAL_VARIATIONS_SCRIPT,
            )
            continue

        if choice == "8":
            run_python_process(
                "Monthly Top Customers",
                process_paths.repo_path("Monthly_Top_Customers", "main.py"),
            )
            continue

        if choice in {"99", "back", "b", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def run_data_automation_tools_menu() -> None:
    while True:
        print("\nData & Automation Tools")
        print()
        print("    01. Automation Processes")
        print("    02. Check Table Updates")
        print("    03. Inventory Obsolescence Manager")
        print("    04. Power BI Reports")
        print("    05. XGBoost Model")
        print()
        print("    99. Back to main menu")
        print()
        try:
            choice = normalize_menu_choice(input("Choose an option: "))
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if choice == "1":
            run_automation_processes_menu()
            continue

        if choice == "2":
            run_check_table_updates_menu()
            continue

        if choice == "3":
            run_inventory_obsolescence_manager_menu()
            continue

        if choice == "4":
            run_python_process("Power BI Reports", process_paths.POWER_BI_REPORTS_SCRIPT)
            continue

        if choice == "5":
            run_python_process("XGBoost Model", process_paths.repo_path("xgboost_model", "main.py"))
            continue

        if choice in {"99", "back", "b", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def run_inventory_obsolescence_manager_menu() -> None:
    while True:
        print("\nInventory Obsolescence Manager")
        print()
        print("    01. Consolidate Inventory Manager")
        print("    02. HBG vs Oracle Inventory Comparison")
        print()
        print("    99. Back to previous menu")
        print()
        try:
            choice = normalize_menu_choice(input("Choose an option: "))
        except KeyboardInterrupt:
            print("\nReturning to previous menu.")
            return

        if choice == "1":
            run_python_process(
                "Consolidate Inventory Manager",
                process_paths.CONSOLIDATED_INVENTORY_VERTICALIZATION_SCRIPT,
            )
            continue

        if choice == "2":
            run_python_process(
                "HBG vs Oracle Inventory Comparison",
                process_paths.HBG_ORACLE_INVENTORY_COMPARISON_SCRIPT,
            )
            continue

        if choice in {"99", "back", "b", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def run_admin_utilities_menu() -> None:
    while True:
        print("\nAdmin / Utilities")
        print()
        print("    01. Desk Procedures")
        print("    02. Install Main Venv Requirements")
        print("    03. Open Main Venv Shell")
        print()
        print("    99. Back to main menu")
        print()
        try:
            choice = normalize_menu_choice(input("Choose an option: "))
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if choice == "1":
            run_python_process("Desk Procedures", process_paths.repo_path("desk_procedures", "main.py"))
            continue

        if choice == "2":
            install_main_venv_requirements()
            continue

        if choice == "3":
            open_main_venv_shell()
            continue

        if choice in {"99", "back", "b", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def run_amazon_menu():
    while True:
        print("\nAmazon")
        print()
        print("Purchase Order Report")
        print("    01. PO Archive Manager")
        print("    02. PO Report")
        print()
        print("PreOrders Report")
        print("    03. PreOrders")
        print()
        print("Customer Order Reports")
        print("    04. Customer Orders")
        print()
        print("AMS Manager")
        print("    05. AMS Manager (monthly)")
        print()
        print("    99. Back to main menu")
        print()
        try:
            subchoice = normalize_menu_choice(input("Choose an option: "))
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if subchoice == "1":
            try:
                po_archive_manager.main()
            except Exception as e:
                print(f"An error occurred while running PO Archive Manager: {e}")
            continue

        amazon_reports = {
            "2": ("Amazon (2) PO Report", "amazon_po/main.py"),
            "3": ("Amazon (3) PreOrders", "amazon_preorders/main.py"),
            "4": ("Amazon (4) Customer Orders", "amazon_customer_orders/main.py"),
            "5": ("Amazon AMS Manager (monthly)", "amazon_ams/manage_ams.py"),
        }

        if subchoice in amazon_reports:
            report_name, script_path = amazon_reports[subchoice]
            if subchoice == "3":
                try:
                    if not confirm_amazon_preorders_files():
                        continue
                except FileNotFoundError as e:
                    print(f"Unable to locate the Amazon (3) PreOrders source files: {e}")
                    continue
            if subchoice == "4":
                try:
                    if not confirm_amazon_customer_orders_files():
                        continue
                except FileNotFoundError as e:
                    print(f"Unable to locate the Amazon (4) Customer Orders source files: {e}")
                    continue
            if subchoice == "5":
                try:
                    if not confirm_amazon_ams_files():
                        continue
                except (FileNotFoundError, ImportError, AttributeError) as e:
                    print(f"Unable to locate the Amazon AMS Manager source files: {e}")
                    continue

            print(f"Running the {report_name}... Please wait.")
            try:
                subprocess.run(["venv/Scripts/python", script_path], check=True)
                print(f"The {report_name} is now ready.")
            except subprocess.CalledProcessError:
                print(f"An error occurred while running {script_path}.")
            continue

        if subchoice in {"99", "back", "b", "return", "menu"}:
            return

        print("Invalid choice. Please select a valid option.")


def run_amazon_rolling_reports_menu():
    while True:
        print("\nAmazon Rolling Reports")
        print()
        print("Weekly Process")
        print("    1. Check Amazon upload table (last 10 weeks)")
        print("    2. Full refresh + all weekly reports")
        print("    3. Rebuild all weekly reports from current pickles")
        print("    4. Full refresh + main two weekly reports only")
        print("    5. Rebuild main two weekly reports only from current pickles")
        print()
        print("Monthly Process")
        print("    6. Add new Monthly file to Cache")
        print("    7. Run Monthly Rolling Report")
        print("    8. Back to main menu")
        print()
        try:
            subchoice = input("Choose an option: ").strip().lower()
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if subchoice == "1":
            try:
                if not confirm_amazon_rolling_reports_check_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Rolling Reports check files: {e}")
                continue
            print("Running SQL check for the latest 10 weeks... Please wait.")
            try:
                subprocess.run(
                    [
                        "venv/Scripts/python",
                        "amazon_rolling_reports/check_last_10_weeks.py",
                    ],
                    check=True,
                )
            except subprocess.CalledProcessError:
                print("The SQL check failed.")
            continue

        if subchoice == "2":
            try:
                if not confirm_amazon_rolling_reports_run_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Rolling Reports source files: {e}")
                continue
            print("Running full refresh + all Amazon Rolling Reports... Please wait.")
            try:
                subprocess.run(
                    ["venv/Scripts/python", "amazon_rolling_reports/main.py", "--refresh-all-caches"],
                    check=True,
                )
                print("The Amazon Rolling Reports are now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running amazon_rolling_reports/main.py.")
            return

        if subchoice == "3":
            try:
                if not confirm_amazon_rolling_reports_run_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Rolling Reports source files: {e}")
                continue
            print("Rebuilding all Amazon Rolling Reports from current pickles... Please wait.")
            try:
                subprocess.run(
                    ["venv/Scripts/python", "amazon_rolling_reports/main.py", "--report-only"],
                    check=True,
                )
                print("The Amazon Rolling Reports are now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running amazon_rolling_reports/main.py --report-only.")
            return

        if subchoice == "4":
            try:
                if not confirm_amazon_rolling_reports_run_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Rolling Reports source files: {e}")
                continue
            print("Running full refresh + main-two Amazon Rolling Reports only... Please wait.")
            try:
                subprocess.run(
                    [
                        "venv/Scripts/python",
                        "amazon_rolling_reports/main.py",
                        "--main-only",
                        "--refresh-all-caches",
                    ],
                    check=True,
                )
                print("The main Amazon Rolling Reports are now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running amazon_rolling_reports/main.py --main-only.")
            return

        if subchoice == "5":
            try:
                if not confirm_amazon_rolling_reports_run_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Rolling Reports source files: {e}")
                continue
            print("Rebuilding the main-two Amazon Rolling Reports from current pickles... Please wait.")
            try:
                subprocess.run(
                    [
                        "venv/Scripts/python",
                        "amazon_rolling_reports/main.py",
                        "--main-only",
                        "--report-only",
                    ],
                    check=True,
                )
                print("The main Amazon Rolling Reports are now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running amazon_rolling_reports/main.py --main-only --report-only.")
            return

        if subchoice == "6":
            try:
                if not confirm_amazon_monthly_sales_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Monthly Sales files: {e}")
                continue
            print("Compiling Amazon monthly sales parquet... Please wait.")
            try:
                subprocess.run(
                    ["venv/Scripts/python", "amazon_rolling_reports/monthly_sales.py"],
                    check=True,
                )
                print("The Amazon monthly sales parquet is now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running amazon_rolling_reports/monthly_sales.py.")
            continue

        if subchoice == "7":
            try:
                if not confirm_amazon_monthly_sales_files():
                    continue
            except FileNotFoundError as e:
                print(f"Unable to locate the Amazon Monthly Sales files: {e}")
                continue
            print("Building standalone Amazon monthly rolling reports... Please wait.")
            try:
                subprocess.run(
                    ["venv/Scripts/python", "amazon_rolling_reports/monthly_rolling_reports.py"],
                    check=True,
                )
                print("The Amazon monthly rolling reports are now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running amazon_rolling_reports/monthly_rolling_reports.py.")
            continue

        if subchoice in ["8", "back", "b", "exit", "quit", "q"]:
            return

        print("Invalid choice. Please select a valid option.")


def run_check_table_updates_menu():
    while True:
        print("\nCheck Table Updates")
        print()
        print("    1. All Updates")
        print("    2. Tables for SSR Summary")
        print("    3. Ebs.Sales Prior 5 Days")
        print("    4. Amazon")
        print("    5. Bookscan")
        print("    6. Barnes & Noble")
        print("    7. Freight Costs")
        print("    8. Back to main menu")
        print()
        try:
            subchoice = input("Choose an option: ").strip().lower()
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if subchoice in ["1", "2", "3", "4", "5", "6", "7"]:
            print("Running table-update SQL check... Please wait.")
            try:
                subprocess.run(
                    [
                        "venv/Scripts/python",
                        "table_check/check_table_updates.py",
                        subchoice,
                    ],
                    check=True,
                )
            except subprocess.CalledProcessError:
                print("The table-update SQL check failed.")
            continue

        if subchoice in ["8", "back", "b", "exit", "quit", "q"]:
            return

        print("Invalid choice. Please select a valid option.")


def run_ssr_daily_summary_menu():
    while True:
        print("\nSSR Daily Summary")
        print()
        print("    1. Run SSR Daily Reporting")
        print("    2. Run SSR Aggregate Totals")
        print("    3. Run SSR Visualization")
        print("    4. Ebs.Sales Prior 5 Days")
        print("    5. Back to main menu")
        print()
        try:
            subchoice = input("Choose an option: ").strip().lower()
        except KeyboardInterrupt:
            print("\nReturning to main menu.")
            return

        if subchoice == "4":
            print("Running table-update SQL check... Please wait.")
            try:
                subprocess.run(
                    [
                        "venv/Scripts/python",
                        "table_check/check_table_updates.py",
                        "3",
                    ],
                    check=True,
                )
            except subprocess.CalledProcessError:
                print("The table-update SQL check failed.")
            continue

        if subchoice == "1":
            print("Running SSR Daily Reporting... Please wait.")
            try:
                subprocess.run(
                    ["venv/Scripts/python", "ssr_daily_summary/ssr_preparation.py"],
                    check=True,
                )
                print("SSR Daily Reporting is now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running ssr_daily_summary/ssr_preparation.py.")
            continue

        if subchoice == "2":
            print("Running SSR Aggregate Totals... Please wait.")
            try:
                subprocess.run(
                    ["venv/Scripts/python", "ssr_daily_summary/ssr_summary.py"],
                    check=True,
                )
                print("SSR Aggregate Totals are now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running ssr_daily_summary/ssr_summary.py.")
            continue

        if subchoice == "3":
            print("Running SSR Visualization... Please wait.")
            try:
                subprocess.run(
                    ["venv/Scripts/python", "ssr_daily_summary/ssr_visualizations.py"],
                    check=True,
                )
                print("SSR Visualization is now ready.")
            except subprocess.CalledProcessError:
                print("An error occurred while running ssr_daily_summary/ssr_visualizations.py.")
            continue

        if subchoice in ["5", "back", "b", "exit", "quit", "q"]:
            return

        print("Invalid choice. Please select a valid option.")


def main():
    print(greet_user())

    while True:
        display_options()
        try:
            choice = normalize_menu_choice(input(
                "\nPlease enter the number of your choice (or type 'info' to learn more): "
            ))
        except KeyboardInterrupt:
            print(get_farewell_message())
            break

        if choice.lower() == "info":
            choice_info = normalize_menu_choice(input(
                "\nEnter the number of the option you want to learn more about: "
            ))
            print(display_info(choice_info))
            continue

        if choice.lower() in ["99", "exit", "quit"]:
            run_program("99")
            break

        run_program(choice)


if __name__ == "__main__":
    main()
