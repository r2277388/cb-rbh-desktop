from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from config import (
    BASE_FOLDER,
    NUMERIC_COLUMNS,
    OUTPUT_PREFIX,
    OUTPUT_COLUMN_RENAMES,
    RAW_FOLDER_SUFFIX,
    REQUIRED_COLUMNS,
    REQUIRED_FILE_KEYWORDS,
    TEXT_COLUMNS,
)
from isbn_utils import normalize_isbn_series


RAW_FOLDER_PATTERN = re.compile(r"^(?P<date>\d{4}_\d{2}_\d{2})_raw_files$", re.IGNORECASE)


@dataclass
class PosBuildResult:
    dataframe: pd.DataFrame
    raw_folder: Path
    output_file: Path
    source_files: list[Path]
    rows_before_dedup: int
    rows_after_dedup: int
    duplicate_rows_removed: int
    rows_with_missing_ean: int


def parse_week_ending(folder_name: str) -> datetime:
    match = RAW_FOLDER_PATTERN.match(folder_name)
    if not match:
        raise ValueError(
            f"Folder name must match yyyy_mm_dd{RAW_FOLDER_SUFFIX}. Received: {folder_name}"
        )
    return datetime.strptime(match.group("date"), "%Y_%m_%d")


def format_output_filename(week_ending: datetime) -> str:
    return f"{OUTPUT_PREFIX}_{week_ending:%Y%m%d}.xlsx"


def get_candidate_raw_folders(base_folder: Path = BASE_FOLDER) -> list[Path]:
    if not base_folder.exists():
        return []
    folders = [path for path in base_folder.iterdir() if path.is_dir() and RAW_FOLDER_PATTERN.match(path.name)]
    return sorted(folders, key=lambda path: path.name)


def resolve_raw_folder(raw_folder: str | Path | None = None) -> Path:
    if raw_folder:
        path = Path(raw_folder)
        if not path.exists():
            raise FileNotFoundError(f"Raw folder not found: {path}")
        if not path.is_dir():
            raise NotADirectoryError(f"Raw folder path is not a directory: {path}")
        return path

    folders = get_candidate_raw_folders()
    if not folders:
        raise FileNotFoundError(
            f"No raw folders matching yyyy_mm_dd{RAW_FOLDER_SUFFIX} were found under {BASE_FOLDER}"
        )
    return folders[-1]


def find_pos_source_files(raw_folder: Path) -> list[Path]:
    files_by_keyword: list[Path] = []
    lower_name_to_path = {path.name.lower(): path for path in raw_folder.iterdir() if path.is_file()}

    for keyword in REQUIRED_FILE_KEYWORDS:
        matches = [
            path
            for lower_name, path in lower_name_to_path.items()
            if lower_name.endswith(".csv") and "pos" in lower_name and keyword in lower_name
        ]
        if not matches:
            raise FileNotFoundError(
                f"Could not find a POS CSV containing '{keyword}' in {raw_folder}"
            )
        files_by_keyword.append(sorted(matches, key=lambda path: path.name.lower())[0])

    return files_by_keyword


def read_pos_csv(file_path: Path) -> pd.DataFrame:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            df = pd.read_csv(
                file_path,
                dtype={column: "string" for column in TEXT_COLUMNS},
                usecols=REQUIRED_COLUMNS,
                encoding=encoding,
            )
            return df
        except Exception as exc:  # pragma: no cover - fallback path
            last_error = exc
    raise RuntimeError(f"Unable to read {file_path}: {last_error}")


def clean_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    for column in NUMERIC_COLUMNS:
        df[column] = pd.to_numeric(
            df[column].astype("string").str.replace(",", "", regex=False),
            errors="coerce",
        )
    return df


def build_combined_pos(raw_folder: str | Path | None = None, output_file: str | Path | None = None) -> PosBuildResult:
    resolved_raw_folder = resolve_raw_folder(raw_folder)
    week_ending = parse_week_ending(resolved_raw_folder.name)
    source_files = find_pos_source_files(resolved_raw_folder)

    frames = []
    for file_path in source_files:
        df = read_pos_csv(file_path)
        df["_source_file"] = file_path.name
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    combined["EAN"] = normalize_isbn_series(combined["EAN"])
    rows_with_missing_ean = int(combined["EAN"].isna().sum())
    combined = combined.loc[combined["EAN"].notna()].copy()
    combined = clean_numeric_columns(combined)

    rows_before_dedup = len(combined)
    combined = combined.drop_duplicates(subset=["EAN"], keep="first").reset_index(drop=True)
    rows_after_dedup = len(combined)
    duplicate_rows_removed = rows_before_dedup - rows_after_dedup
    combined.drop(columns=["_source_file"], inplace=True)
    combined.rename(columns=OUTPUT_COLUMN_RENAMES, inplace=True)

    output_path = Path(output_file) if output_file else resolved_raw_folder / format_output_filename(week_ending)
    save_to_excel(combined, output_path)

    return PosBuildResult(
        dataframe=combined,
        raw_folder=resolved_raw_folder,
        output_file=output_path,
        source_files=source_files,
        rows_before_dedup=rows_before_dedup,
        rows_after_dedup=rows_after_dedup,
        duplicate_rows_removed=duplicate_rows_removed,
        rows_with_missing_ean=rows_with_missing_ean,
    )


def save_to_excel(df: pd.DataFrame, output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        sheet_name = "pos_combined"
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        header_format = workbook.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#D8E4BC",
                "border": 1,
            }
        )
        text_format = workbook.add_format({"num_format": "@"})
        integer_format = workbook.add_format({"num_format": "#,##0"})

        for col_idx, column in enumerate(df.columns):
            worksheet.write(0, col_idx, column, header_format)
            if column in TEXT_COLUMNS:
                worksheet.set_column(col_idx, col_idx, 18, text_format)
            else:
                worksheet.set_column(col_idx, col_idx, 14, integer_format)

        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
        worksheet.freeze_panes(1, 0)


def preview_dataframe(df: pd.DataFrame, rows: int = 25) -> str:
    if df.empty:
        return "No rows to display."
    return df.head(rows).to_string(index=False)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Combine Barnes & Noble POS non-book files.")
    parser.add_argument(
        "--raw-folder",
        help="Full path to a yyyy_mm_dd_raw_files folder. Defaults to the latest matching folder.",
    )
    parser.add_argument(
        "--output-file",
        help="Optional full output path for the combined POS Excel file.",
    )
    parser.add_argument(
        "--preview-rows",
        type=int,
        default=25,
        help="Number of rows to print to the terminal after building the file.",
    )
    return parser


def print_result_summary(result: PosBuildResult, preview_rows: int = 25) -> None:
    print()
    print(f"Raw folder: {result.raw_folder}")
    print("Source files:")
    for file_path in result.source_files:
        print(f"  - {file_path.name}")
    print(f"DataFrame shape: {result.dataframe.shape}")
    print(f"Rows before de-duplication: {result.rows_before_dedup:,}")
    print(f"Duplicate EAN rows removed: {result.duplicate_rows_removed:,}")
    print(f"Rows dropped for invalid/missing EAN: {result.rows_with_missing_ean:,}")
    print(f"Rows after de-duplication: {result.rows_after_dedup:,}")
    print(f"Saved file: {result.output_file}")
    print()
    print(f"First {preview_rows} rows:")
    print(preview_dataframe(result.dataframe, rows=preview_rows))


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    result = build_combined_pos(raw_folder=args.raw_folder, output_file=args.output_file)
    print_result_summary(result, preview_rows=args.preview_rows)


if __name__ == "__main__":
    main()
