from __future__ import annotations

from datetime import datetime

import pandas as pd


def build_active_asins(cleaned: pd.DataFrame, as_of_date: datetime) -> pd.DataFrame:
    start_dates = pd.to_datetime(
        cleaned["Sell-through start date"], errors="coerce", format="mixed"
    )
    end_dates = pd.to_datetime(
        cleaned["Sell-through end date"], errors="coerce", format="mixed"
    )
    as_of = pd.Timestamp(as_of_date.date())
    active_mask = (
        cleaned["Offer state"].astype("string").str.strip().str.casefold().eq("active")
        & cleaned["Status"].astype("string").str.strip().str.casefold().eq("accepted")
        & start_dates.le(as_of)
        & end_dates.ge(as_of)
    )
    active = cleaned.loc[
        active_mask,
        [
            "ISBN", "Title", "Publisher", "ASIN",
            "Sell-through start date", "Sell-through end date",
        ],
    ].copy()
    active["Days Left"] = (end_dates.loc[active.index] - as_of).dt.days
    return active.sort_values(
        ["Days Left", "Sell-through end date", "ASIN"], kind="stable"
    ).reset_index(drop=True)


def _header_format(workbook, background_color: str):
    return workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": background_color,
            "border": 1,
            "text_wrap": True,
            "valign": "top",
        }
    )


def write_grouped_status_changes(
    writer: pd.ExcelWriter, changes: pd.DataFrame
) -> None:
    sheet_name = "Status Changes"
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet
    standard_header = _header_format(workbook, "#1F4E78")
    comparison_header = _header_format(workbook, "#E26B0A")
    metadata_header = _header_format(workbook, "#403151")
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
    worksheet.freeze_panes(1, 0)

    status_order = (
        changes["Status"]
        .value_counts(dropna=False)
        .sort_values(kind="stable")
        .index.tolist()
    )
    output_row = 0
    for status in status_order:
        group = (
            changes[changes["Status"].isna()]
            if pd.isna(status)
            else changes[changes["Status"].eq(status)]
        )
        for column_number, column_name in enumerate(changes.columns):
            if column_name in {"ISBN", "Title", "Publisher"}:
                header_format = metadata_header
            elif column_name in {
                "Change Type", "Previous Status", "Previous Status description"
            }:
                header_format = comparison_header
            else:
                header_format = standard_header
            worksheet.write(output_row, column_number, column_name, header_format)
        worksheet.set_row(output_row, 32)
        output_row += 1
        group.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=output_row,
            header=False,
            index=False,
        )
        output_row += len(group) + 1

    for column_number, column_name in enumerate(changes.columns):
        values = changes[column_name].dropna().astype(str)
        sampled_width = max(
            [len(str(column_name)), *(len(value) for value in values.head(500))]
        )
        width = min(max(sampled_width + 2, 11), 45)
        cell_format = date_format if "date" in column_name.casefold() else None
        worksheet.set_column(column_number, column_number, width, cell_format)
