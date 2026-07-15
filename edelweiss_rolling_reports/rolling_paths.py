from __future__ import annotations

from pathlib import Path


edelweiss_rolling_folder = Path(
    r"G:\SALES\2026 Sales Reports\Sell-Through Reporting\Edelweiss"
)
cache_dir = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Edelweiss\cache")
edelweiss_source_folder = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Edelweiss\Edelweiss_Sales"
)
local_review_dir = Path(__file__).resolve().parent / "review_output"
sample_workbook = Path(__file__).resolve().parents[1] / "Edelweiss" / "Week 25 - 2026 Rolling Edelweiss (062726).xlsx"

sales_cache_file = cache_dir / "edelweiss_sales.parquet"
metadata_cache_file = cache_dir / "edelweiss_metadata.parquet"
manual_missing_weeks_file = cache_dir / "edelweiss_manual_missing_weeks.parquet"
