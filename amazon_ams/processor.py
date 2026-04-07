from __future__ import annotations

import importlib
import shutil
import time
from pathlib import Path

import pandas as pd

from loader_asin_mapping import load_asin_mapping
from loader_item import upload_item
from loader_monthly_reports import load_monthly_data
import UPDATE_ams_config as ams_config
from paths import process_paths


OUTPUT_PICKLE = Path(__file__).with_name("combined_amazon_ads_by_asin.pkl")
OUTPUT_EXCEL = process_paths.AMAZON_AMS_OUTPUT_EXCEL
ERROR_LOG = Path(__file__).with_name("processing_errors.log")
ARCHIVE_DIR = Path(__file__).with_name("archive")


def archive_current_outputs(reason: str) -> list[Path]:
    targets = [OUTPUT_PICKLE, OUTPUT_EXCEL, ERROR_LOG]
    existing = [path for path in targets if path.exists()]
    if not existing:
        return []

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    archive_dir = ARCHIVE_DIR / f"{timestamp}_{reason}"
    archive_dir.mkdir(parents=True, exist_ok=True)

    archived_paths: list[Path] = []
    for path in existing:
        destination = archive_dir / path.name
        shutil.copy2(path, destination)
        archived_paths.append(destination)

    return archived_paths


def available_months() -> list[str]:
    return sorted(_config().month_list)


def get_month_entry(month: str) -> dict[str, object]:
    config = _config()
    if month not in config.tab_dict:
        raise KeyError(f"Month not found in configuration: {month}")
    return config.tab_dict[month]


def _config():
    global ams_config
    ams_config = importlib.reload(ams_config)
    return ams_config


def _load_shared_inputs() -> tuple[pd.DataFrame, pd.DataFrame]:
    asin_mapping = load_asin_mapping()
    item_df = upload_item()
    return asin_mapping, item_df


def _merge_item_data(df: pd.DataFrame, item_df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or item_df.empty:
        return df
    return pd.merge(df, item_df, on="ISBN", how="left")


def process_single_month(month: str, asin_mapping: pd.DataFrame, item_df: pd.DataFrame) -> pd.DataFrame:
    df_month = load_monthly_data(get_month_entry(month), asin_mapping, month)
    return _merge_item_data(df_month, item_df)


def build_full_dataset(months: list[str] | None = None) -> tuple[pd.DataFrame, list[str]]:
    asin_mapping, item_df = _load_shared_inputs()
    selected_months = months or available_months()

    combined_df = pd.DataFrame()
    errors: list[str] = []

    for month in selected_months:
        try:
            print(f"Processing {month}...")
            df_month = process_single_month(month, asin_mapping, item_df)
            combined_df = pd.concat([combined_df, df_month], ignore_index=True)
        except Exception as exc:
            error_msg = f"Failed for {month}: {exc}"
            print(error_msg)
            errors.append(error_msg)

    if not combined_df.empty and "period" in combined_df.columns:
        combined_df = combined_df.sort_values(["period", "ASIN"]).reset_index(drop=True)

    return combined_df, errors


def save_outputs(df: pd.DataFrame, errors: list[str]) -> None:
    if not df.empty:
        OUTPUT_EXCEL.parent.mkdir(parents=True, exist_ok=True)
        df.to_pickle(OUTPUT_PICKLE)
        df.to_excel(OUTPUT_EXCEL, index=False)
        print(f"Saved pickle: {OUTPUT_PICKLE}")
        print(f"Saved Excel: {OUTPUT_EXCEL}")
    else:
        print("No data was successfully combined.")

    if errors:
        ERROR_LOG.write_text("\n".join(errors) + "\n", encoding="utf-8")
        print(f"Saved error log: {ERROR_LOG}")
    elif ERROR_LOG.exists():
        ERROR_LOG.unlink()


def run_full_rebuild(months: list[str] | None = None, archive_reason: str = "full_rebuild") -> tuple[pd.DataFrame, list[str], list[Path]]:
    archived = archive_current_outputs(archive_reason)
    combined_df, errors = build_full_dataset(months=months)
    save_outputs(combined_df, errors)
    return combined_df, errors, archived


def run_incremental_update(month: str | None = None) -> tuple[pd.DataFrame, list[Path]]:
    selected_month = month or (available_months()[-1] if available_months() else None)
    if selected_month is None:
        raise RuntimeError("No months are available for incremental processing.")

    asin_mapping, item_df = _load_shared_inputs()
    df_month = process_single_month(selected_month, asin_mapping, item_df)

    if OUTPUT_PICKLE.exists():
        existing = pd.read_pickle(OUTPUT_PICKLE)
        if "period" in existing.columns:
            existing = existing[existing["period"] != selected_month]
        combined_df = pd.concat([existing, df_month], ignore_index=True)
    else:
        combined_df = df_month

    if not combined_df.empty and "period" in combined_df.columns:
        combined_df = combined_df.sort_values(["period", "ASIN"]).reset_index(drop=True)

    archived = archive_current_outputs(f"incremental_{selected_month.replace('-', '')}")
    save_outputs(combined_df, [])
    return combined_df, archived
