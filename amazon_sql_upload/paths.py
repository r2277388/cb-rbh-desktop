from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path


def _load_shared_paths():
    shared_path = Path(__file__).resolve().parents[1] / "paths" / "process_paths.py"
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
