from .process_paths import *

amz_weekly_sales = str(AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["sales"])
amz_weekly_inventory = str(AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["inventory"])
amz_weekly_traffic = str(AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["traffic"])
amz_catalog = str(AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["catalog"])
ypticod = str(ORACLE_YPTICOD_FILE)

amazon_po_folder = str(AMAZON_PO_FOLDER)
amazon_rolling_folder = str(AMAZON_ROLLING_OUTPUT_FOLDER)
dp_folders = {name: str(path) for name, path in AMAZON_ROLLING_DP_FOLDERS.items()}

folder_path = str(UK_ROLLING_SOURCE_FOLDER)
output_path = str(UK_ROLLING_OUTPUT_FILE)


def saved_query_location():
    return str(SSR_QUERY_OUTPUT_FILE)


def saved_viz_location():
    return str(SSR_VIZ_OUTPUT_FILE)
