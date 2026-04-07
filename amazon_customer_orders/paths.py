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

DOWNLOADS_FOLDER = _shared.DOWNLOADS_FOLDER
ATELIER_AMAZON_SALES_FOLDER = _shared.ATELIER_AMAZON_BASE_FOLDER / "Sales"
ATELIER_AMAZON_INVENTORY_FOLDER = _shared.ATELIER_AMAZON_INVENTORY_FOLDER
ATELIER_AMAZON_TRAFFIC_FOLDER = _shared.ATELIER_AMAZON_BASE_FOLDER / "Traffic"
ATELIER_AMAZON_CATALOG_FOLDER = _shared.ATELIER_AMAZON_CATALOG_FOLDER
ORACLE_YPTICOD_FILE = _shared.ORACLE_YPTICOD_FILE
AMAZON_CUSTOMER_ORDERS_OUTPUT_FILE = _shared.AMAZON_CUSTOMER_ORDERS_OUTPUT_FILE
