from datetime import datetime
from pathlib import Path
from shared.bookscan_calendar import bookscan_week

REPO_ROOT = Path(__file__).resolve().parent.parent


def repo_path(*parts: str) -> Path:
    return (REPO_ROOT / Path(*parts)).resolve()


# Shared external locations
DOWNLOADS_FOLDER = Path(r"G:\SALES\Amazon\RBH\DOWNLOADED_FILES")
ORACLE_YPTICOD_FILE = Path(r"J:\Metadata Reports\Oracle YPTICOD.xlsx")
DATAWAREHOUSE_SHAREPOINT_FOLDERS = {
    "Sam": Path(r"C:\Users\sdm\OneDrive - chroniclebooks.com\Finance Department - Documents\DataWarehouse"),
    "Barrett": Path(r"C:\Users\rbh\OneDrive - chroniclebooks.com\Finance Department - Documents\DataWarehouse"),
}
ATELIER_AMAZON_BASE_FOLDER = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon")
ATELIER_AMAZON_CATALOG_FOLDER = ATELIER_AMAZON_BASE_FOLDER / "Catalog"
ATELIER_AMAZON_INVENTORY_FOLDER = ATELIER_AMAZON_BASE_FOLDER / "Inventory"
AMAZON_WEEKLY_BASE_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Amazon"
)
AMAZON_WEEKLY_REPORTS_DIR = AMAZON_WEEKLY_BASE_FOLDER
USER_DESKTOP = Path(r"C:\Users\rbh\Desktop")
AMAZON_SQL_UPLOAD_SOURCE_FOLDERS = {
    "sales": ATELIER_AMAZON_BASE_FOLDER / "Sales",
    "inventory": ATELIER_AMAZON_BASE_FOLDER / "Inventory",
    "traffic": ATELIER_AMAZON_BASE_FOLDER / "Traffic",
    "catalog": ATELIER_AMAZON_BASE_FOLDER / "Catalog",
}
AMAZON_PO_FOLDER = Path(r"G:\SALES\Amazon\PURCHASE ORDERS\2026")
AMAZON_PO_ROOT_FOLDER = AMAZON_PO_FOLDER.parent
AMAZON_PO_ANALYSIS_INPUT_FILE = (
    AMAZON_PO_ROOT_FOLDER / "atelier" / "po_analysis" / "PurchaseOrderItems.csv"
)
AMAZON_PO_CURRENT_FILE = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\po_analysis\current_amaz_po_file.csv"
)
AMAZON_PO_ARCHIVE_GLOB = "POItemExport_*.csv"
AMAZON_PO_CURRENT_PREORDERS_FILE = (
    AMAZON_PO_ROOT_FOLDER / "atelier" / "po_analysis" / "current_amaz_preorders.csv"
)
AMAZON_PO_DATAWAREHOUSE_ANALYSIS_FILE = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\po_analysis\current_amaz_preorders.csv"
)
AMAZON_PO_DATAWAREHOUSE_DUMP_FILE = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\po_analysis\amazon_order_py_dump.xlsx"
)
AMAZON_PO_DATAWAREHOUSE_ARCHIVE_DIR = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\po_analysis\po_archive"
)
AMAZON_ROLLING_OUTPUT_FOLDER = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Amazon"
)
FRONTLIST_TRACKING_FOLDER = Path(r"G:\SALES\2026 Sales Reports\Frontlist Tracking")
INGRAM_DAILY_REPORT_FOLDER = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Ingram"
)
BN_WEEKLY_REPORT_FOLDER = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Barnes & Noble"
)
TARGET_NOC_OUTPUT_FOLDER = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Target NOC"
)
TARGET_NOC_DP_FOLDERS = {
    "Quadrille Publishing Limited": Path(
        r"G:\SALES\Distribution_Partners\Quadrille\QD REPORTS\Sell Through Reporting\Target NOC"
    ),
    "Galison": Path(r"G:\SALES\Distribution_Partners\Galison\GA REPORTS\Target NOC"),
}
TARGET_NOC_SALES_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Target NOC\TargetNOC_Sales"
)
TARGET_NOC_INVENTORY_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Target NOC\TargetNOC_Inventory"
)
TARGET_NOC_CACHE_DIR = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier TargetNOC\cache")
BOOKSCAN_WEEKLY_REPORT_FOLDER = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Bookscan"
)
EDELWEISS_WEEKLY_REPORT_FOLDER = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Edelweiss"
)
AWBC_WEEKLY_REPORT_FOLDER = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\AWBC"
)
POWER_BI_REPORTS_FOLDER = Path(
    r"\\sfx\SFNY-Files\SF\Groups\Visual Dashboards"
)
POWER_BI_BARRETT_REPORT_FOLDERS = {
    "Faire Sales": Path(r"\\sfx\sfny-files\SF\Groups\Tableau Dashboards\Faire Sales"),
    "Inventory": Path(r"\\sfx\sfny-files\SF\Groups\Tableau Dashboards\Inventory"),
    "Production": Path(r"\\sfx\sfny-files\SF\Groups\Tableau Dashboards\Production"),
}
CONSOLIDATED_INVENTORY_VERTICALIZATION_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\consolidated_inventory"
)
INVENTORY_OBSOLESCENCE_FOLDER = repo_path("Inventory_Obsolescence")
HBG_MOON_HEADER_WORKBOOK = INVENTORY_OBSOLESCENCE_FOLDER / "hbg_inventory_moon_headers_only.xlsx"
HBG_ORACLE_COMPARISON_OUTPUT_FOLDER = CONSOLIDATED_INVENTORY_VERTICALIZATION_FOLDER / "hbg_oracle_comparison"
GENOPS_BASE_FOLDER = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier GenOps")
GEN_EDITORIAL_SOURCE_WORKBOOK = Path(r"J:\SCHEDULE\SchedPubGrpAll.xlsx")
GEN_EDITORIAL_OUTPUT_FOLDER = Path(r"J:\SCHEDULE\Schedule Changes\2026")
GEN_EDITORIAL_CACHE_DIR = GENOPS_BASE_FOLDER / "cache"
GEN_EDITORIAL_CACHE_FILE = GEN_EDITORIAL_CACHE_DIR / "cache_gen_editorial.parquet"
GEN_EDITORIAL_REPORT_FILE = GEN_EDITORIAL_OUTPUT_FOLDER / "General Editorial Data Variations.xlsx"
INVENTORY_DAILY_FINANCE_ONLY_FOLDER = Path(r"G:\OPS\Inventory\Daily\Finance_Only")
CURRENT_AMAZON_PREORDERS_FILE = Path(
    r"G:\SALES\Amazon\PREORDERS\2026\current_amaz_preorders.xlsx"
)
AMAZON_PREORDERS_OUTPUT_FOLDER = CURRENT_AMAZON_PREORDERS_FILE.parent
BN_RAW_BASE_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Barnes & Noble"
)
CHRONICLE_ASIN_MAPPING_FILE = DOWNLOADS_FOLDER / "Chronicle-AsinMapping.xlsx"
AMAZON_CUSTOMER_ORDERS_OUTPUT_FILE = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\weekly_customer_order\amazon_weekly_customer_order_py.xlsx"
)
AMAZON_SQL_UPLOAD_OUTPUT_DIR = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\vc_weekly_summary"
)
AMAZON_SQL_UPLOAD_WEEKLY_SUMMARIES_DIR = (
    DATAWAREHOUSE_SHAREPOINT_FOLDERS["Barrett"] / "Amazon" / "WeeklySummaries"
)
UK_ROLLING_SOURCE_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier UK\Script Files"
)
UK_ROLLING_CACHE_DIR = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier UK\cache")
UK_ROLLING_OUTPUT_FOLDER = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Abrams & Chronicle"
)
UK_ROLLING_OUTPUT_FILE = UK_ROLLING_OUTPUT_FOLDER / "Title Sales Week ##.xlsx"
SSR_QUERY_OUTPUT_FILE = Path(
    r"G:\SALES\2026 Sales Reports\SSR\SSR_Template\rbh_daily_py.xlsx"
)
SSR_VIZ_OUTPUT_FILE = Path(
    r"G:\SALES\2026 Sales Reports\SSR\SSR_Template\ssr_summary_chart.html"
)
AMAZON_ROLLING_DP_FOLDERS = {
    "Galison": Path(r"G:\SALES\Distribution_Partners\Galison\GA REPORTS\Amazon\2026"),
    "Hardie Grant": Path(
        r"G:\SALES\Distribution_Partners\Hardie Grant\HG REPORTS\Sell Through Reporting\Amazon\2026"
    ),
    "Laurence King": Path(
        r"G:\SALES\Distribution_Partners\Laurence King\LK REPORTS\Amazon\2026"
    ),
    "Levine Querido": Path(
        r"G:\SALES\Distribution_Partners\Levine Querido\LQ REPORTS\Sell Through Reporting\Amazon\2026"
    ),
    "Paperblanks": Path(
        r"G:\SALES\Distribution_Partners\Paperblanks\PB REPORTS\Sell-Through Reporting\Amazon\2026"
    ),
    "Quadrille": Path(
        r"G:\SALES\Distribution_Partners\Quadrille\QD REPORTS\Sell Through Reporting\Amazon\2026"
    ),
    "Sierra Club": Path(
        r"G:\SALES\Distribution_Partners\Sierra Club\SC REPORTS\Amazon\2026"
    ),
    "Tourbillon": Path(
        r"G:\SALES\Distribution_Partners\Twirl\TW-REPORTS\Sell Through Reporting\Amazon\2026"
    ),
    "Creative Company": Path(
        r"G:\SALES\Distribution_Partners\Creative Company\CC REPORTS\Amazon\2026"
    ),
}
BN_ROLLING_DP_FOLDERS = {
    "Galison": Path(
        r"G:\Sales\Distribution_Partners\Galison\GA REPORTS\Barnes & Noble\2026"
    ),
    "Hardie Grant": Path(
        r"G:\SALES\Distribution_Partners\Hardie Grant\HG REPORTS\Sell Through Reporting\Barnes & Noble\2026"
    ),
    "Laurence King": Path(
        r"G:\SALES\Distribution_Partners\Laurence King\LK REPORTS\Barnes & Noble\2026"
    ),
    "Levine Querido": Path(
        r"G:\Sales\Distribution_Partners\Levine Querido\LQ REPORTS\Sell Through Reporting\Barnes & Noble\2026"
    ),
    "Paperblanks": Path(
        r"G:\Sales\Distribution_Partners\Paperblanks\PB REPORTS\Sell-Through Reporting\Barnes & Noble\2026"
    ),
    "Quadrille": Path(
        r"G:\Sales\Distribution_Partners\Quadrille\QD REPORTS\Sell Through Reporting\Barnes & Noble\2026"
    ),
    "Sierra Club": Path(
        r"G:\Sales\Distribution_Partners\Sierra Club\SC REPORTS\Barnes & Noble\2026"
    ),
    "Tourbillon": Path(
        r"G:\Sales\Distribution_Partners\Twirl\TW-REPORTS\Sell Through Reporting\Barnes & Noble\2026"
    ),
    "Creative Company": Path(
        r"G:\Sales\Distribution_Partners\Creative Company\CC REPORTS\Barnes & Noble\2026"
    ),
}
BOOKSCAN_ROLLING_DP_FOLDERS = {
    "Galison": Path(r"G:\Sales\Distribution_Partners\Galison\GA REPORTS\Bookscan\2026"),
    "Hardie Grant": Path(
        r"G:\SALES\Distribution_Partners\Hardie Grant\HG REPORTS\Sell Through Reporting\Bookscan\2026"
    ),
    "Laurence King": Path(
        r"G:\SALES\Distribution_Partners\Laurence King\LK REPORTS\Bookscan\2026"
    ),
    "Levine Querido": Path(
        r"G:\Sales\Distribution_Partners\Levine Querido\LQ REPORTS\Sell Through Reporting\Bookscan\2026"
    ),
    "Paperblanks": Path(
        r"G:\Sales\Distribution_Partners\Paperblanks\PB REPORTS\Sell-Through Reporting\Bookscan\2026"
    ),
    "Quadrille": Path(
        r"G:\Sales\Distribution_Partners\Quadrille\QD REPORTS\Sell Through Reporting\Bookscan\2026"
    ),
    "Sierra Club": Path(
        r"G:\Sales\Distribution_Partners\Sierra Club\SC REPORTS\Bookscan\2026"
    ),
    "Tourbillon": Path(
        r"G:\Sales\Distribution_Partners\Twirl\TW-REPORTS\Sell Through Reporting\Bookscan\2026"
    ),
    "Creative Company": Path(
        r"G:\Sales\Distribution_Partners\Creative Company\CC REPORTS\Bookscan\2026"
    ),
}

# Launcher-owned local files
AMAZON_CUSTOMER_ORDERS_SCRIPT = repo_path("amazon_customer_orders", "main.py")
AMAZON_SQL_UPLOAD_SCRIPT = repo_path("amazon_sql_upload", "main.py")
AMAZON_SQL_UPLOAD_MANUAL_KEY_FILE = repo_path("amazon_sql_upload", "asin_manual_key.py")
AMAZON_SQL_UPLOAD_REMOVAL_LIST_FILE = repo_path(
    "amazon_sql_upload", "asin_removal_list.py"
)
AMAZON_ROLLING_CHECK_SCRIPT = repo_path(
    "amazon_rolling_reports", "check_last_10_weeks.py"
)
AMAZON_ROLLING_SQL_FILE = repo_path(
    "shared", "sql", "amazon_rolling_reports", "last_10_weeks.sql"
)
AMAZON_ROLLING_REPORTS_SCRIPT = repo_path("amazon_rolling_reports", "main.py")
AMAZON_MONTHLY_ROLLING_REPORTS_SCRIPT = repo_path(
    "amazon_rolling_reports", "monthly_rolling_reports.py"
)
AMAZON_MONTHLY_SALES_ROOT = ATELIER_AMAZON_BASE_FOLDER / "Sales_Monthly"
AMAZON_MONTHLY_SALES_FALLBACK_ROOT = ATELIER_AMAZON_BASE_FOLDER / "Sales" / "sales_monthly"
AMAZON_MONTHLY_SALES_SCRIPT = repo_path("amazon_rolling_reports", "monthly_sales.py")
AMAZON_MONTHLY_SALES_PARQUET_NAME = "amazon_monthly_sales.parquet"
CONSOLIDATED_INVENTORY_VERTICALIZATION_SCRIPT = repo_path(
    "consolidate_inventory_verticalization", "main.py"
)
HBG_ORACLE_INVENTORY_COMPARISON_SCRIPT = repo_path(
    "Inventory_Obsolescence", "hbg_oracle_inventory_comparison.py"
)
AMAZON_ROLLING_CUSTOMER_ORDERS_PICKLE = repo_path("rr_customer_orders.pkl")
AMAZON_ROLLING_UNITS_SHIPPED_PICKLE = repo_path("rr_units_shipped.pkl")
AMAZON_ROLLING_PO_PICKLE = repo_path("latest_amazon_po.pkl")
AMAZON_ROLLING_CUSTOMER_ORDERS_QUERY = repo_path(
    "amazon_rolling_reports", "query_co.py"
)
AMAZON_ROLLING_UNITS_SHIPPED_QUERY = repo_path("amazon_rolling_reports", "query_us.py")
AMAZON_ROLLING_DATE_CHECK_QUERY = repo_path(
    "amazon_rolling_reports", "query_datecheck.py"
)
AMAZON_AMS_MANAGER_SCRIPT = repo_path("amazon_ams", "manage_ams.py")
AMAZON_AMS_PROCESS_SCRIPT = repo_path("amazon_ams", "main.py")
AMAZON_AMS_MONTHLY_CAMPAIGN_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\AMS_Monthly_Campaign"
)
AMAZON_AMS_MONTHLY_CAMPAIGN_HISTORY_PARQUET = (
    AMAZON_AMS_MONTHLY_CAMPAIGN_FOLDER / "cache" / "ams_monthly_campaigns.parquet"
)
AMAZON_AMS_FINAL_REPORTS_FOLDER = Path(
    r"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2026"
)
AMAZON_AMS_CONFIG_FILE = repo_path("amazon_ams", "UPDATE_ams_config.py")
AMAZON_AMS_OUTPUT_PICKLE = repo_path("amazon_ams", "combined_amazon_ads_by_asin.pkl")
AMAZON_AMS_OUTPUT_EXCEL = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\ams_summaries\combined_amazon_ads_by_asin.xlsx"
)
AMAZON_AMS_ERROR_LOG = repo_path("amazon_ams", "processing_errors.log")
FRONTLIST_SUPERCHARGED_SCRIPT = repo_path("FLTracking_Supercharged", "main.py")
FRONTLIST_SUPERCHARGED_OUTPUT_DIR = repo_path("FLTracking_Supercharged", "output")
FRONTLIST_AMAZON_SELLTHROUGH_SQL = repo_path(
    "FLTracking_Supercharged", "sql", "amazon_sellthrough_latest.sql"
)
FRONTLIST_FAIRE_QTY_SQL = repo_path("FLTracking_Supercharged", "sql", "faire_qty.sql")
FRONTLIST_FAIRE_ORDERS_SQL = repo_path(
    "FLTracking_Supercharged", "sql", "faire_orders.sql"
)
BN_ROLLING_REPORTS_SCRIPT = repo_path("bn_rolling_reports", "main.py")
BOOKSCAN_ROLLING_REPORTS_SCRIPT = repo_path("bookscan_rolling_reports", "main.py")
EDELWEISS_ROLLING_REPORTS_SCRIPT = repo_path("edelweiss_rolling_reports", "main.py")
AWBC_ROLLING_REPORTS_SCRIPT = repo_path("awbc_rolling_reports", "main.py")
POWER_BI_REPORTS_SCRIPT = repo_path("power_bi_reports", "main.py")
CROSS_GAP_SCRIPT = repo_path("cross_gap", "main.py")
GEN_EDITORIAL_VARIATIONS_SCRIPT = repo_path("gen_editorial_variations", "main.py")
CROSS_GAP_CONFIG_FILE = repo_path("cross_gap", "title_groups.json")
CROSS_GAP_OUTPUT_DIR = Path(
    r"\\sfx\sfny-files\SF\Groups\Sales\2026 Sales Reports\Reports\Cross Gap"
)
CROSS_GAP_CACHE_DIR = CROSS_GAP_OUTPUT_DIR / "cache"
CROSS_GAP_TASK_NAME = "Chronicle Weekly Cross Gap Report"
CROSS_GAP_TASK_LOCATION = "Task Scheduler Library"
CROSS_GAP_SCHEDULE_DAY = "SUN"
CROSS_GAP_SCHEDULE_TIME = "12:00"
CROSS_GAP_SCHEDULE_DESCRIPTION = "Every Sunday at 12:00 PM"
REPRINT_INDICATOR_AUTOMATION_SCRIPT = repo_path(
    "reprint_indicator_automation", "main.py"
)
EXCEL_REFRESH_SCRIPT = repo_path("tools", "excel_refresh.py")
TITLE_LOOKUP_WORKBOOK = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Z-Archive\1 - title lookup.xlsx"
)
TITLE_LOOKUP_CONNECTION_NAME = "Query - title_query"
TITLE_LOOKUP_TABLE_NAME = "title_query"
TITLE_LOOKUP_TASK_NAME = "Chronicle Weekly Title Lookup Refresh"
TITLE_LOOKUP_TASK_LOCATION = "Task Scheduler Library"
TITLE_LOOKUP_SCHEDULE_DAY = "SUN"
TITLE_LOOKUP_SCHEDULE_TIME = "07:00"
TITLE_LOOKUP_SCHEDULE_DESCRIPTION = "Every Sunday at 7:00 AM"
GEN_EDITORIAL_TASK_NAME = "Chronicle Weekly General Editorial Data Variations"
GEN_EDITORIAL_TASK_LOCATION = "Task Scheduler Library"
GEN_EDITORIAL_SCHEDULE_DAYS = "MON"
GEN_EDITORIAL_SCHEDULE_TIME = "09:30"
GEN_EDITORIAL_SCHEDULE_DESCRIPTION = "Every Monday at 9:30 AM"
EXPORT_REPORT_WORKBOOKS = (
    Path(r"G:\SALES\2026 Sales Reports\Reports\Export\Export_DataDump.xlsx"),
    Path(r"G:\SALES\2026 Sales Reports\Reports\Export\UK Detail Report_2025_2026.xlsx"),
)
EXPORT_REPORT_REFRESH_SCRIPT = repo_path("tools", "export_reports_refresh.py")
EXPORT_REPORT_TASK_NAME = "Chronicle Weekly Export Reports Refresh"
EXPORT_REPORT_SCHEDULE_DAY = "SUN"
EXPORT_REPORT_SCHEDULE_TIME = "10:30"
EXPORT_REPORT_SCHEDULE_DESCRIPTION = "Every Sunday at 10:30 AM"


def amazon_sql_upload_output_file(for_date: datetime | None = None) -> Path:
    date_value = for_date or datetime.now()
    return (
        AMAZON_SQL_UPLOAD_OUTPUT_DIR
        / f"amazon_update__{date_value.strftime('%Y_%m_%d')}.xlsx"
    )


def amazon_sql_upload_weekly_summary_file(for_date: datetime | None = None) -> Path:
    date_value = for_date or datetime.now()
    return (
        AMAZON_SQL_UPLOAD_WEEKLY_SUMMARIES_DIR
        / f"amazon_update__{date_value.strftime('%Y_%m_%d')}.xlsx"
    )


def amazon_weekly_report_file(for_date: datetime) -> Path:
    week = bookscan_week(for_date).week
    return (
        AMAZON_WEEKLY_REPORTS_DIR
        / f"w{week:02d}__{for_date.strftime('%Y_%m_%d')}.xlsx"
    )

