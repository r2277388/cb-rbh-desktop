from __future__ import annotations

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

bookscan_rolling_folder = Path(_shared.BOOKSCAN_WEEKLY_REPORT_FOLDER)
bookscan_dp_folders = {
    name: Path(path) for name, path in _shared.BOOKSCAN_ROLLING_DP_FOLDERS.items()
}
inventory_detail_workbook = _shared.INVENTORY_DAILY_FINANCE_ONLY_FOLDER / "Inventory Detail.xlsx"
cache_dir = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Bookscan\cache")
local_review_dir = Path(__file__).resolve().parent / "review_output"
sales_cache_file = cache_dir / "bookscan_sales.parquet"
metadata_cache_file = cache_dir / "bookscan_metadata.parquet"
inventory_cache_file = cache_dir / "bookscan_inventory_detail.parquet"
manual_missing_weeks_file = cache_dir / "bookscan_manual_missing_weeks.parquet"
