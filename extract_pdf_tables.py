from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
import pdfplumber


PDF_PATH = Path(r"c:\Users\rbh\Downloads\Deal 24669-rights.pdf")
OUTPUT_XLSX = Path(r"c:\Users\rbh\code\Deal 24669-rights_tables.xlsx")


def sanitize_sheet_name(name: str) -> str:
    clean = re.sub(r"[:\\/?*\[\]]", "_", name).strip()
    return clean[:31] or "Table"


def table_to_dataframe(table: list[list[str | None]]) -> pd.DataFrame:
    if not table:
        return pd.DataFrame()
    max_cols = max(len(row) for row in table)
    normalized = [row + [None] * (max_cols - len(row)) for row in table]
    # Remove common icon glyphs and normalize whitespace.
    cleaned = []
    for row in normalized:
        out_row = []
        for cell in row:
            if cell is None:
                out_row.append(None)
                continue
            text = str(cell).replace("\ue409", "").strip()
            out_row.append(text)
        cleaned.append(out_row)
    normalized = cleaned
    return pd.DataFrame(normalized)


def reorder_royalty_rates_rows(df: pd.DataFrame) -> pd.DataFrame:
    text = df.fillna("").astype(str)
    first_col = text.iloc[:, 0].str.strip().str.lower()
    marker_idx = first_col[first_col.str.contains("royalty rates", na=False)].index
    if marker_idx.empty:
        return df

    marker = marker_idx[0]
    expected = {"properties", "formats", "territory", "channel", "language"}
    header_row = None
    for i in range(marker + 1, min(marker + 8, len(text))):
        row_tokens = set(
            token.strip().lower()
            for token in text.iloc[i].tolist()
            if token and token.strip()
        )
        if len(expected.intersection(row_tokens)) >= 3:
            header_row = i
            break
    if header_row is None:
        return df

    # Move meaningful pre-header rows below the header so layout matches what users see.
    pre = df.iloc[:marker].copy()
    if pre.empty:
        return df

    drop_prefixes = (
        "+ contract dates",
        "+ exclusions",
        "+ recoupable allocations",
        "+ fixed fees",
        "+ payments schedule",
    )
    mask_data = pre.apply(
        lambda r: any((str(v).strip() if v is not None else "") for v in r.tolist()),
        axis=1,
    )
    mask_section = (
        pre.fillna("")
        .astype(str)
        .iloc[:, 0]
        .str.strip()
        .str.lower()
        .str.startswith(drop_prefixes)
    )
    data_rows = pre[mask_data & ~mask_section]
    if data_rows.empty:
        return df

    marker_row = df.iloc[[marker]]
    header = df.iloc[[header_row]]
    after_header = df.iloc[header_row + 1 :]
    return pd.concat([marker_row, header, data_rows, after_header], ignore_index=True)


def main() -> None:
    extracted: list[tuple[str, pd.DataFrame]] = []
    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "intersection_tolerance": 6,
        "snap_tolerance": 3,
    }

    with pdfplumber.open(PDF_PATH) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_tables = page.extract_tables(table_settings=table_settings)
            for table_idx, table in enumerate(page_tables, start=1):
                df = table_to_dataframe(table)
                if df.empty:
                    continue
                if df.replace("", pd.NA).dropna(how="all").empty:
                    continue
                df = reorder_royalty_rates_rows(df)
                sheet_name = sanitize_sheet_name(f"p{page_num}_t{table_idx}")
                extracted.append((sheet_name, df))

    if not extracted:
        raise RuntimeError("No tables were detected in the PDF.")

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        used: set[str] = set()
        for base_name, df in extracted:
            name = base_name
            suffix = 1
            while name in used:
                suffix += 1
                name = sanitize_sheet_name(f"{base_name}_{suffix}")
            used.add(name)
            df.to_excel(writer, sheet_name=name, index=False, header=False)

    print(f"Saved {len(extracted)} tables to: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
