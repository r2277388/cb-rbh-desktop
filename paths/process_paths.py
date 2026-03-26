from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent


def repo_path(*parts: str) -> Path:
    return (REPO_ROOT / Path(*parts)).resolve()


# Shared external locations
DOWNLOADS_FOLDER = Path(r"G:\SALES\Amazon\RBH\DOWNLOADED_FILES")
ORACLE_YPTICOD_FILE = Path(r"J:\Metadata Reports\Oracle YPTICOD.xlsx")
AMAZON_WEEKLY_BASE_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Amazon"
)
USER_DESKTOP = Path(r"C:\Users\rbh\Desktop")
AMAZON_SQL_UPLOAD_SOURCE_FOLDERS = {
    "sales": AMAZON_WEEKLY_BASE_FOLDER / "Sales",
    "inventory": AMAZON_WEEKLY_BASE_FOLDER / "Inventory",
    "traffic": AMAZON_WEEKLY_BASE_FOLDER / "Traffic",
    "catalog": AMAZON_WEEKLY_BASE_FOLDER / "Catalog",
}
AMAZON_PO_FOLDER = Path(r"G:\SALES\Amazon\PURCHASE ORDERS\2026")
AMAZON_PO_ROOT_FOLDER = AMAZON_PO_FOLDER.parent
AMAZON_PO_ANALYSIS_INPUT_FILE = (
    AMAZON_PO_ROOT_FOLDER / "atelier" / "po_analysis" / "PurchaseOrderItems.csv"
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
CONSOLIDATED_INVENTORY_VERTICALIZATION_FOLDER = Path(
    r"F:\RBH_Finternal\consolidated_inventory"
)
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
    r"G:\SALES\Amazon\RBH\weekly_customer_order\atelier\amazon_weekly_customer_order_py.xlsx"
)
UK_ROLLING_SOURCE_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Weekly reports\2026\Abrams & Chronicle\Script Files"
)
UK_ROLLING_OUTPUT_FILE = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Abrams & Chronicle\Title Sales Week ##.xlsx"
)
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
CONSOLIDATED_INVENTORY_VERTICALIZATION_SCRIPT = repo_path(
    "consolidate_inventory_verticalization", "main.py"
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
AMAZON_AMS_CONFIG_FILE = repo_path("amazon_ams", "UPDATE_ams_config.py")
AMAZON_AMS_OUTPUT_PICKLE = repo_path("amazon_ams", "combined_amazon_ads_by_asin.pkl")
AMAZON_AMS_OUTPUT_EXCEL = repo_path("amazon_ams", "combined_amazon_ads_by_asin.xlsx")
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
