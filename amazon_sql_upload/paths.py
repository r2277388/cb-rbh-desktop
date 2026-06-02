from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
import sys


def _load_shared_paths():
    repo_root = Path(__file__).resolve().parents[1]
    shared_path = repo_root / "paths" / "process_paths.py"
    repo_root_text = str(repo_root)
    if repo_root_text not in sys.path:
        sys.path.insert(0, repo_root_text)
    spec = spec_from_file_location("_shared_process_paths", shared_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load shared process paths from {shared_path}")
    module = module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_shared = _load_shared_paths()

amz_weekly_sales = str(_shared.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["sales"])
amz_weekly_inventory = str(_shared.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["inventory"])
amz_weekly_traffic = str(_shared.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["traffic"])
amz_catalog = str(_shared.AMAZON_SQL_UPLOAD_SOURCE_FOLDERS["catalog"])
ypticod = str(_shared.ORACLE_YPTICOD_FILE)

AMAZON_SQL_UPLOAD_OUTPUT_DIR = _shared.AMAZON_SQL_UPLOAD_OUTPUT_DIR
AMAZON_SQL_UPLOAD_WEEKLY_SUMMARIES_DIR = _shared.AMAZON_SQL_UPLOAD_WEEKLY_SUMMARIES_DIR
AMAZON_WEEKLY_REPORTS_DIR = _shared.AMAZON_WEEKLY_REPORTS_DIR


def amazon_sql_upload_output_file(for_date=None):
    return _shared.amazon_sql_upload_output_file(for_date)


def amazon_sql_upload_weekly_summary_file(for_date=None):
    return _shared.amazon_sql_upload_weekly_summary_file(for_date)


def amazon_weekly_report_file(for_date):
    return _shared.amazon_weekly_report_file(for_date)
