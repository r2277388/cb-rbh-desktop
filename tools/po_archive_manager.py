import pathlib
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

from paths import process_paths

PO_ANALYSIS = process_paths.AMAZON_PO_CURRENT_FILE.parent
ARCHIVE = process_paths.AMAZON_PO_DATAWAREHOUSE_ARCHIVE_DIR
ORDER_DATE_COLUMN = "Order date"
SUMMARY_TOTAL_COLUMNS = (
    "Total requested cost",
    "Total accepted cost",
    "Total received cost",
    "Total cancelled cost",
)
DATE_LIKE_PATTERN = re.compile(
    r"^\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}|\d{1,2}-[A-Za-z]{3}-\d{2,4})"
)


def make_archive_name(src_path: pathlib.Path) -> pathlib.Path:
    return ARCHIVE / src_path.name


def unique_path(p: pathlib.Path) -> pathlib.Path:
    if not p.exists():
        return p
    base = p.stem
    suffix = p.suffix
    i = 1
    while True:
        candidate = p.with_name(f"{base}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def numeric_series(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype("string")
        .str.strip()
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def date_series(series: pd.Series) -> pd.Series:
    cleaned = series.astype("string").str.strip()
    date_like = cleaned.where(cleaned.str.match(DATE_LIKE_PATTERN, na=False))
    return pd.to_datetime(
        date_like,
        errors="coerce",
        format="mixed",
    ).dt.date


def format_numeric_total(total: float) -> str:
    if total.is_integer():
        return f"{int(total):,}"
    return f"{total:,.2f}"


def summarize_csv_contents(csv_path: pathlib.Path) -> str:
    df = pd.read_csv(csv_path, dtype="string", low_memory=False)
    lines = [f"CSV summary for {csv_path.name}", f"Rows: {len(df):,}"]

    missing_total_columns = [
        column_name for column_name in SUMMARY_TOTAL_COLUMNS if column_name not in df.columns
    ]
    total_columns = [
        (column_name, numeric_series(df[column_name]))
        for column_name in SUMMARY_TOTAL_COLUMNS
        if column_name in df.columns
    ]

    if not total_columns:
        lines.append("Requested cost columns found: none")
        if missing_total_columns:
            lines.append("Missing requested cost columns:")
            for column_name in missing_total_columns:
                lines.append(f"  {column_name}")
        return "\n".join(lines)

    lines.append("Cost column totals:")
    for column_name, converted in total_columns:
        lines.append(f"  {column_name}: {format_numeric_total(float(converted.sum()))}")

    if missing_total_columns:
        lines.append("Missing requested cost columns:")
        for column_name in missing_total_columns:
            lines.append(f"  {column_name}")

    if ORDER_DATE_COLUMN not in df.columns:
        lines.append(f"{ORDER_DATE_COLUMN} column not found; date totals unavailable.")
        return "\n".join(lines)

    summary_df = pd.DataFrame({"__date": date_series(df[ORDER_DATE_COLUMN])})
    for column_name, converted in total_columns:
        summary_df[column_name] = converted

    summary_df = summary_df.dropna(subset=["__date"])
    if summary_df.empty:
        lines.append(f"No parseable {ORDER_DATE_COLUMN} values found; date totals unavailable.")
        return "\n".join(lines)

    grouped = (
        summary_df.groupby("__date", dropna=True)
        .sum(numeric_only=True)
        .sort_index()
    )
    lines.append("")
    lines.append(f"Cost totals by {ORDER_DATE_COLUMN}:")
    for date_value, totals in grouped.iterrows():
        lines.append(f"  {date_value.strftime('%m/%d/%Y')}:")
        for column_name in grouped.columns:
            total = float(totals[column_name])
            lines.append(f"    {column_name}: {format_numeric_total(total)}")

    return "\n".join(lines)


def main():
    # GUI: pick the new vendor file first
    root = tk.Tk()
    root.withdraw()
    src_file = filedialog.askopenfilename(
        title="Select new Amazon Vendor Central PO Report (CSV)",
        initialdir=str(PO_ANALYSIS) if PO_ANALYSIS.exists() else None,
        filetypes=[("CSV Files", "*.csv"), ("All files", "*.*")],
    )
    if not src_file:
        print("No file selected. Exiting.")
        return

    src_path = pathlib.Path(src_file)
    if not src_path.exists():
        messagebox.showerror("Error", f"Selected file does not exist:\n{src_path}")
        return

    # Ensure analysis folder exists
    PO_ANALYSIS.mkdir(parents=True, exist_ok=True)
    ARCHIVE.mkdir(parents=True, exist_ok=True)

    dest = process_paths.AMAZON_PO_CURRENT_FILE
    archive_target = unique_path(make_archive_name(src_path))
    try:
        shutil.copy2(str(src_path), str(dest))
        shutil.copy2(str(src_path), str(archive_target))
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Failed to copy new file into output folders:\n{e}",
        )
        return

    try:
        summary = summarize_csv_contents(src_path)
    except Exception as e:
        summary = f"CSV summary unavailable: {e}"

    print("New file copied to:")
    print(f"  {dest}")
    print(f"  {archive_target}")
    print()
    print(summary)
    print("Complete.")


if __name__ == "__main__":
    main()
