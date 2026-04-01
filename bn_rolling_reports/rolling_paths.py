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

bn_rolling_folder = Path(_shared.BN_WEEKLY_REPORT_FOLDER)
bn_dp_folders = {name: Path(path) for name, path in _shared.BN_ROLLING_DP_FOLDERS.items()}
cache_dir = Path(__file__).resolve().parent / "cache"
sales_cache_file = cache_dir / "bn_customer_sales.parquet"
inventory_cache_file = cache_dir / "bn_inventory_snapshots.parquet"
