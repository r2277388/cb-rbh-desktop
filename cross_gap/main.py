from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from paths import process_paths
from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db


METADATA_COLUMNS = ["HBG_Num", "HBG", "Rep_Num", "Rep"]
DISPLAY_METADATA_HEADERS = ["HBG Num", "HBG", "Rep Num", "Rep"]
VALUE_COLUMN = "Sales+OO"
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


@dataclass(frozen=True)
class SupplementalSalesColumn:
    label: str
    ip_family_name: str
    start_period: str
    end_period: str
    formats: tuple[str, ...] = ()

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "SupplementalSalesColumn":
        return cls(
            label=str(raw["label"]).strip(),
            ip_family_name=str(raw["ip_family_name"]).strip(),
            start_period=str(raw["start_period"]).strip(),
            end_period=str(raw["end_period"]).strip(),
            formats=tuple(str(value).strip() for value in raw.get("formats", []) if str(value).strip()),
        )

    def validate(self, group_name: str) -> None:
        if not self.label:
            raise ValueError(f"{group_name} has a supplemental sales column without a label.")
        if not self.ip_family_name:
            raise ValueError(f"{group_name} supplemental column {self.label} needs ip_family_name.")
        if not self.start_period or not self.end_period:
            raise ValueError(f"{group_name} supplemental column {self.label} needs start_period and end_period.")


@dataclass(frozen=True)
class TitleGroup:
    name: str
    isbns: tuple[str, ...] = ()
    ip_family_name: str | None = None
    formats: tuple[str, ...] = ()
    supplemental_sales_columns: tuple[SupplementalSalesColumn, ...] = ()

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "TitleGroup":
        name = str(raw["name"]).strip()
        if not name:
            raise ValueError("Every title group needs a non-empty name.")

        return cls(
            name=name,
            isbns=tuple(str(value).strip() for value in raw.get("isbns", []) if str(value).strip()),
            ip_family_name=raw.get("ip_family_name"),
            formats=tuple(str(value).strip() for value in raw.get("formats", []) if str(value).strip()),
            supplemental_sales_columns=tuple(
                SupplementalSalesColumn.from_dict(value)
                for value in raw.get("supplemental_sales_columns", [])
            ),
        )

    def validate(self) -> None:
        if not self.isbns and not self.ip_family_name:
            raise ValueError(f"{self.name} needs either isbns or ip_family_name.")
        for column in self.supplemental_sales_columns:
            column.validate(self.name)


def sql_string(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def sql_in(values: tuple[str, ...]) -> str:
    return ", ".join(sql_string(value) for value in values)


def load_title_groups(config_path: Path) -> list[TitleGroup]:
    with config_path.open("r", encoding="utf-8") as config_file:
        raw_config = json.load(config_file)

    groups = [TitleGroup.from_dict(raw_group) for raw_group in raw_config.get("groups", [])]
    if not groups:
        raise ValueError(f"No title groups found in {config_path}")

    seen = set()
    for group in groups:
        group.validate()
        lowered = group.name.lower()
        if lowered in seen:
            raise ValueError(f"Duplicate title group name: {group.name}")
        seen.add(lowered)
    return groups


def group_to_dict(group: TitleGroup) -> dict[str, Any]:
    raw: dict[str, Any] = {"name": group.name}
    if group.isbns:
        raw["isbns"] = list(group.isbns)
    if group.ip_family_name:
        raw["ip_family_name"] = group.ip_family_name
    if group.formats:
        raw["formats"] = list(group.formats)
    if group.supplemental_sales_columns:
        raw["supplemental_sales_columns"] = [
            {
                "label": column.label,
                "ip_family_name": column.ip_family_name,
                "start_period": column.start_period,
                "end_period": column.end_period,
                **({"formats": list(column.formats)} if column.formats else {}),
            }
            for column in group.supplemental_sales_columns
        ]
    return raw


def save_title_groups(config_path: Path, groups: list[TitleGroup]) -> None:
    raw_config = {"groups": [group_to_dict(group) for group in groups]}
    config_path.write_text(
        json.dumps(raw_config, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def group_condition(group: TitleGroup, item_alias: str = "i") -> str:
    pieces = []
    if group.isbns:
        pieces.append(f"{item_alias}.ISBN IN ({sql_in(group.isbns)})")
    if group.ip_family_name:
        family_condition = f"{item_alias}.IP_FAMILY_NAME = {sql_string(group.ip_family_name)}"
        if group.formats:
            family_condition += f" AND {item_alias}.FORMAT IN ({sql_in(group.formats)})"
        pieces.append(f"({family_condition})")
    return "(" + " OR ".join(pieces) + ")"


def supplemental_sales_condition(
    column: SupplementalSalesColumn,
    item_alias: str = "i",
    sales_alias: str = "sd",
) -> str:
    pieces = [f"{item_alias}.IP_FAMILY_NAME = {sql_string(column.ip_family_name)}"]
    if column.formats:
        pieces.append(f"{item_alias}.FORMAT IN ({sql_in(column.formats)})")
    pieces.append(f"{sales_alias}.PERIOD BETWEEN {sql_string(column.start_period)} AND {sql_string(column.end_period)}")
    return "(" + " AND ".join(pieces) + ")"


def family_case_sql(
    groups: list[TitleGroup],
    item_alias: str = "i",
) -> str:
    lines = ["CASE"]
    for group in groups:
        lines.append(f"        WHEN {group_condition(group, item_alias)} THEN {sql_string(group.name)}")
    lines.append("        ELSE 'Other'")
    lines.append("    END")
    return "\n".join(lines)


def all_group_conditions(groups: list[TitleGroup], item_alias: str = "i") -> str:
    return "\n        OR ".join(group_condition(group, item_alias) for group in groups)


def build_sales_query(groups: list[TitleGroup]) -> str:
    family_case = family_case_sql(groups)
    where_conditions = all_group_conditions(groups)
    return f"""
WITH OSD AS (
    SELECT
        tt.ean13 AS ISBN,
        tt.active_datevalue AS osd
    FROM tmm.cb_Import_Title_Tasks tt
    WHERE
        tt.date_desc = 'On Sale Date'
        AND tt.active_datevalue IS NOT NULL
        AND tt.printingnumber = 1
)
SELECT
    LEFT(billto.PARTYSITENUMBER, 8) AS HBG_Num,
    billto.PARTYSITENAME AS HBG,
    sr.SALESREP_NUMBER AS Rep_Num,
    sr.NAME AS Rep,
    osd.osd AS OSD,
    i.ISBN,
    i.SHORT_TITLE AS Title,
    SUM(sd.salesqty) AS SalesQty,
    {family_case} AS Family,
    'Sales + OO' AS ValueHeader
FROM summary.CustomerTitleMonthlySales sd
INNER JOIN ebs.Item i
    ON sd.ITEM_ID = i.ITEM_ID
INNER JOIN ebs.Customer shipto
    ON shipto.SITE_USE_ID = sd.SHIP_TO_SITE_USE_ID
LEFT JOIN ebs.Customer billto
    ON billto.SITE_USE_ID = shipto.BILL_TO_SITE_USE_ID
INNER JOIN ebs.SalesRep sr
    ON sr.SALESREP_ID = billto.PRIMARY_SALESREP_ID
LEFT JOIN OSD
    ON OSD.ISBN = i.ISBN
WHERE
    sd.SALETYPECODE = 'N'
    AND sd.DISTRIBUTION_DIRECT = 'N'
    AND i.PRODUCT_TYPE IN ('BK', 'FT')
    AND (
        {where_conditions}
    )
    AND sr.SALESREP_NUMBER <> '1061'
GROUP BY
    LEFT(billto.PARTYSITENUMBER, 8),
    billto.PARTYSITENAME,
    sr.SALESREP_NUMBER,
    sr.NAME,
    osd.osd,
    i.ISBN,
    i.SHORT_TITLE,
    {family_case};
""".strip()


def build_supplemental_sales_query(group: TitleGroup, column: SupplementalSalesColumn) -> str:
    condition = supplemental_sales_condition(column)
    return f"""
SELECT
    LEFT(billto.PARTYSITENUMBER, 8) AS HBG_Num,
    billto.PARTYSITENAME AS HBG,
    sr.SALESREP_NUMBER AS Rep_Num,
    sr.NAME AS Rep,
    CAST(NULL AS date) AS OSD,
    {sql_string(column.label)} AS ISBN,
    {sql_string(column.label)} AS Title,
    SUM(sd.salesqty) AS SalesQty,
    {sql_string(group.name)} AS Family,
    'Sales' AS ValueHeader
FROM summary.CustomerTitleMonthlySales sd
INNER JOIN ebs.Item i
    ON sd.ITEM_ID = i.ITEM_ID
INNER JOIN ebs.Customer shipto
    ON shipto.SITE_USE_ID = sd.SHIP_TO_SITE_USE_ID
LEFT JOIN ebs.Customer billto
    ON billto.SITE_USE_ID = shipto.BILL_TO_SITE_USE_ID
INNER JOIN ebs.SalesRep sr
    ON sr.SALESREP_ID = billto.PRIMARY_SALESREP_ID
WHERE
    sd.SALETYPECODE = 'N'
    AND sd.DISTRIBUTION_DIRECT = 'N'
    AND i.PRODUCT_TYPE IN ('BK', 'FT')
    AND {condition}
    AND sr.SALESREP_NUMBER <> '1061'
GROUP BY
    LEFT(billto.PARTYSITENUMBER, 8),
    billto.PARTYSITENAME,
    sr.SALESREP_NUMBER,
    sr.NAME;
""".strip()


def build_orders_query(groups: list[TitleGroup]) -> str:
    family_case = family_case_sql(groups)
    where_conditions = all_group_conditions(groups)
    return f"""
WITH OSD AS (
    SELECT
        tt.ean13 AS ISBN,
        tt.active_datevalue AS osd
    FROM tmm.cb_Import_Title_Tasks tt
    WHERE
        tt.date_desc = 'On Sale Date'
        AND tt.active_datevalue IS NOT NULL
        AND tt.printingnumber = 1
)
SELECT
    ho.AccountNumber AS HBG_Num,
    c.PARTYSITENAME AS HBG,
    sr.SALESREP_NUMBER AS Rep_Num,
    sr.NAME AS Rep,
    osd.osd AS OSD,
    i.ISBN,
    i.short_title AS Title,
    SUM(ho.quantity) AS OrderUnits,
    {family_case} AS Family,
    'Sales + OO' AS ValueHeader
FROM Hachette.HachetteOrders ho
INNER JOIN (
    SELECT DISTINCT
        c2.PARTYSITENUMBER,
        c2.PARTYSITENAME,
        c2.PRIMARY_SALESREP_ID
    FROM ebs.Customer c2
    WHERE
        c2.PARTYSITENUMBER IS NOT NULL
        AND c2.SITE_USE_STATUS = 'A'
        AND c2.SITE_USE_CODE = 'BILL_TO'
) c
    ON ho.AccountNumber = c.PARTYSITENUMBER
INNER JOIN ebs.Item i
    ON ho.isbn = i.ITEM_TITLE
INNER JOIN ebs.SalesRep sr
    ON sr.SALESREP_ID = c.PRIMARY_SALESREP_ID
LEFT JOIN OSD
    ON OSD.ISBN = i.ISBN
WHERE
    i.PRODUCT_TYPE IN ('BK', 'FT')
    AND (
        {where_conditions}
    )
    AND sr.SALESREP_NUMBER <> '1061'
GROUP BY
    ho.AccountNumber,
    c.PARTYSITENAME,
    sr.SALESREP_NUMBER,
    sr.NAME,
    osd.osd,
    i.ISBN,
    i.short_title,
    {family_case};
""".strip()


def safe_cache_stem(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("_")
    return cleaned or "cache"


def supplemental_cache_path(
    group: TitleGroup,
    column: SupplementalSalesColumn,
    cache_dir: Path = process_paths.CROSS_GAP_CACHE_DIR,
) -> Path:
    filename = f"{safe_cache_stem(group.name)}__{safe_cache_stem(column.label)}.csv"
    return cache_dir / filename


def supplemental_cache_columns() -> list[str]:
    return [*METADATA_COLUMNS, "OSD", "ISBN", "Title", "SalesQty", "Family", "ValueHeader"]


def normalize_supplemental_cache_frame(df: pd.DataFrame) -> pd.DataFrame:
    output = df.reindex(columns=supplemental_cache_columns()).copy()
    output["OSD"] = pd.to_datetime(output["OSD"], errors="coerce")
    output["SalesQty"] = pd.to_numeric(output["SalesQty"], errors="coerce").fillna(0)
    for column in METADATA_COLUMNS + ["ISBN", "Title", "Family", "ValueHeader"]:
        output[column] = output[column].fillna("").astype(str)
    return output


def load_or_create_supplemental_cache(
    engine: Any,
    group: TitleGroup,
    column: SupplementalSalesColumn,
    cache_dir: Path = process_paths.CROSS_GAP_CACHE_DIR,
) -> pd.DataFrame:
    cache_file = supplemental_cache_path(group, column, cache_dir)

    if cache_file.exists():
        print(f"Loading cached Cross Gap column: {cache_file}")
        return normalize_supplemental_cache_frame(pd.read_csv(cache_file))

    print(f"Creating cached Cross Gap column: {cache_file}")
    df = fetch_data_from_db(engine, build_supplemental_sales_query(group, column))
    cache_file.parent.mkdir(parents=True, exist_ok=True)
    df = normalize_supplemental_cache_frame(df)
    df.to_csv(cache_file, index=False)
    return df


def fetch_supplemental_sales_data(
    engine: Any,
    groups: list[TitleGroup],
    cache_dir: Path = process_paths.CROSS_GAP_CACHE_DIR,
) -> pd.DataFrame:
    frames = [
        load_or_create_supplemental_cache(engine, group, column, cache_dir)
        for group in groups
        for column in group.supplemental_sales_columns
    ]
    if not frames:
        return pd.DataFrame(columns=supplemental_cache_columns())
    return pd.concat(frames, ignore_index=True)


def supplemental_sales_queries(groups: list[TitleGroup]) -> list[tuple[str, str]]:
    return [
        (f"Cached Supplemental Sales Query - {group.name} - {column.label}", build_supplemental_sales_query(group, column))
        for group in groups
        for column in group.supplemental_sales_columns
    ]


def fetch_cross_gap_data(groups: list[TitleGroup]) -> tuple[pd.DataFrame, pd.DataFrame]:
    engine = get_connection()
    sales_df = fetch_data_from_db(engine, build_sales_query(groups))
    supplemental_sales_df = fetch_supplemental_sales_data(engine, groups)
    if not supplemental_sales_df.empty:
        sales_df = pd.concat([sales_df, supplemental_sales_df], ignore_index=True)
    orders_df = fetch_data_from_db(engine, build_orders_query(groups))
    return sales_df, orders_df


def fetch_title_lookup(isbns: list[str]) -> dict[str, str]:
    if not isbns:
        return {}
    unique_isbns = sorted(set(isbns))
    query = f"""
SELECT
    i.ISBN,
    i.SHORT_TITLE AS Title
FROM ebs.Item i
WHERE i.ISBN IN ({sql_in(tuple(unique_isbns))});
""".strip()
    engine = get_connection()
    df = fetch_data_from_db(engine, query)
    if df.empty:
        return {}
    return {
        str(row.ISBN): str(row.Title)
        for row in df.itertuples(index=False)
        if getattr(row, "ISBN", None) is not None
    }


def print_groupings(config_path: Path = process_paths.CROSS_GAP_CONFIG_FILE) -> None:
    groups = load_title_groups(config_path)
    explicit_isbns = [isbn for group in groups for isbn in group.isbns]
    title_lookup = fetch_title_lookup(explicit_isbns)

    print("\nCross Gap Groupings")
    print(f"Config: {config_path}")
    for index, group in enumerate(groups, start=1):
        print()
        print(f"{index}. {group.name}")
        if group.isbns:
            for isbn in group.isbns:
                title = title_lookup.get(isbn, "Title not found")
                print(f"   {isbn}  {title}")
        if group.ip_family_name:
            format_text = ", ".join(group.formats) if group.formats else "Any"
            print(f"   IP family: {group.ip_family_name}")
            print(f"   Format(s): {format_text}")


def parse_isbn_text(isbn_text: str) -> tuple[str, ...]:
    isbns = [
        piece.strip()
        for piece in re.split(r"[,;\s]+", isbn_text.strip())
        if piece.strip()
    ]
    return tuple(dict.fromkeys(isbns))


def add_grouping_interactive(config_path: Path = process_paths.CROSS_GAP_CONFIG_FILE) -> None:
    groups = load_title_groups(config_path)
    name = input("Grouping name: ").strip()
    if not name:
        print("No grouping name entered.")
        return
    if any(group.name.lower() == name.lower() for group in groups):
        print(f"A grouping named {name} already exists.")
        return

    isbn_text = input("ISBNs (comma, space, or semicolon separated): ").strip()
    isbns = parse_isbn_text(isbn_text)
    if not isbns:
        print("No ISBNs entered.")
        return

    new_group = TitleGroup(name=name, isbns=isbns)
    new_group.validate()
    groups.append(new_group)
    save_title_groups(config_path, groups)
    print(f"Added grouping: {name}")
    for isbn in isbns:
        print(f"  {isbn}")


def remove_grouping_interactive(config_path: Path = process_paths.CROSS_GAP_CONFIG_FILE) -> None:
    groups = load_title_groups(config_path)
    print("\nCross Gap Groupings")
    for index, group in enumerate(groups, start=1):
        print(f"    {index}. {group.name}")
    print()
    choice = input("Enter the grouping number or exact name to remove: ").strip()
    if not choice:
        print("No grouping selected.")
        return

    selected_index = None
    if choice.isdigit():
        selected_index = int(choice) - 1
    else:
        for index, group in enumerate(groups):
            if group.name.lower() == choice.lower():
                selected_index = index
                break

    if selected_index is None or selected_index < 0 or selected_index >= len(groups):
        print("Grouping not found.")
        return

    selected_group = groups[selected_index]
    confirm = input(f"Remove {selected_group.name}? Type YES to confirm: ").strip()
    if confirm != "YES":
        print("Removal cancelled.")
        return

    groups.pop(selected_index)
    save_title_groups(config_path, groups)
    print(f"Removed grouping: {selected_group.name}")


def run_cross_gap_menu(config_path: Path = process_paths.CROSS_GAP_CONFIG_FILE) -> None:
    while True:
        print("\nCross Gap")
        print()
        print("    1. Run Cross Gap report")
        print("    2. Show current groupings")
        print("    3. Add grouping")
        print("    4. Remove grouping")
        print("    5. Back to main menu")
        print()
        choice = input("Choose an option: ").strip().lower()

        if choice == "1":
            run(config_path)
            continue

        if choice == "2":
            try:
                print_groupings(config_path)
            except Exception as exc:
                print(f"Unable to show groupings: {exc}")
            continue

        if choice == "3":
            try:
                add_grouping_interactive(config_path)
            except Exception as exc:
                print(f"Unable to add grouping: {exc}")
            continue

        if choice == "4":
            try:
                remove_grouping_interactive(config_path)
            except Exception as exc:
                print(f"Unable to remove grouping: {exc}")
            continue

        if choice in {"5", "b", "back", "return", "menu", "q", "quit", "exit"}:
            return

        print("Invalid choice. Please select a valid option.")


def normalize_source_frame(df: pd.DataFrame, quantity_column: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[*METADATA_COLUMNS, "OSD", "ISBN", "Title", "Family", "ValueHeader", quantity_column])

    output = df.copy()
    output["ISBN"] = output["ISBN"].astype(str)
    output[quantity_column] = pd.to_numeric(output[quantity_column], errors="coerce").fillna(0)
    output["OSD"] = pd.to_datetime(output["OSD"], errors="coerce")
    for column in METADATA_COLUMNS + ["Title", "Family", "ValueHeader"]:
        output[column] = output[column].fillna("").astype(str)
    return output


def combine_sales_and_orders(sales_df: pd.DataFrame, orders_df: pd.DataFrame) -> pd.DataFrame:
    sales = normalize_source_frame(sales_df, "SalesQty")
    orders = normalize_source_frame(orders_df, "OrderUnits")

    keys = [*METADATA_COLUMNS, "OSD", "ISBN", "Title", "Family", "ValueHeader"]
    combined = sales.merge(orders, on=keys, how="outer")
    combined["SalesQty"] = pd.to_numeric(combined["SalesQty"], errors="coerce").fillna(0)
    combined["OrderUnits"] = pd.to_numeric(combined["OrderUnits"], errors="coerce").fillna(0)
    combined[VALUE_COLUMN] = combined["SalesQty"] + combined["OrderUnits"]
    combined = combined[combined[VALUE_COLUMN] != 0].copy()
    return combined


def sheet_name(name: str, used_names: set[str]) -> str:
    cleaned = INVALID_SHEET_CHARS.sub(" ", name).strip() or "Sheet"
    cleaned = cleaned[:31]
    candidate = cleaned
    index = 2
    while candidate.lower() in used_names:
        suffix = f" {index}"
        candidate = cleaned[: 31 - len(suffix)] + suffix
        index += 1
    used_names.add(candidate.lower())
    return candidate


def family_metadata(group: TitleGroup, group_df: pd.DataFrame) -> pd.DataFrame:
    metadata = (
        group_df[["ISBN", "Title", "OSD", "ValueHeader"]]
        .drop_duplicates(subset=["ISBN"])
        .sort_values(["OSD", "Title", "ISBN"], na_position="last")
    )
    configured_order = [*group.isbns, *(column.label for column in group.supplemental_sales_columns)]
    if not configured_order:
        return metadata

    order_lookup = {column_name: index for index, column_name in enumerate(configured_order)}
    metadata["_ConfiguredOrder"] = metadata["ISBN"].map(order_lookup)
    return (
        metadata.sort_values(
            ["_ConfiguredOrder", "OSD", "Title", "ISBN"],
            na_position="last",
        )
        .drop(columns=["_ConfiguredOrder"])
        .reset_index(drop=True)
    )


def family_matrix(group_df: pd.DataFrame, isbn_order: list[str]) -> pd.DataFrame:
    metadata = (
        group_df.sort_values(["HBG_Num", "HBG", "Rep_Num", "Rep"])
        .groupby("HBG_Num", dropna=False, as_index=False)
        .agg({"HBG": "first", "Rep_Num": "first", "Rep": "first"})
    )
    summarized = (
        group_df.groupby(["HBG_Num", "ISBN"], dropna=False, as_index=False)[VALUE_COLUMN]
        .sum()
    )
    matrix = summarized.pivot_table(
        index="HBG_Num",
        columns="ISBN",
        values=VALUE_COLUMN,
        aggfunc="sum",
        fill_value=0,
    )
    matrix = matrix.reindex(columns=isbn_order, fill_value=0).reset_index()
    matrix = matrix.merge(metadata, on="HBG_Num", how="left")
    matrix = matrix[[*METADATA_COLUMNS, *isbn_order]]
    matrix = matrix[matrix[isbn_order].sum(axis=1) != 0]
    matrix["_RowTotal"] = matrix[isbn_order].sum(axis=1)
    return (
        matrix.sort_values(["_RowTotal", "HBG"], ascending=[False, True])
        .drop(columns=["_RowTotal"])
        .reset_index(drop=True)
    )


def write_group_sheet(
    writer: pd.ExcelWriter,
    workbook: Any,
    group: TitleGroup,
    group_df: pd.DataFrame,
    used_sheet_names: set[str],
) -> None:
    worksheet_name = sheet_name(group.name, used_sheet_names)
    worksheet = workbook.add_worksheet(worksheet_name)
    writer.sheets[worksheet_name] = worksheet

    title_format = workbook.add_format(
        {"font_size": 11, "text_wrap": True, "align": "center", "valign": "vcenter"}
    )
    gray_label_format = workbook.add_format(
        {"font_color": "white", "bg_color": "#A6A6A6", "align": "right", "valign": "vcenter"}
    )
    osd_label_format = workbook.add_format({"bold": True, "align": "right", "valign": "vcenter"})
    metadata_header_format = workbook.add_format(
        {"bg_color": "#B7DEE8", "border": 1, "valign": "vcenter"}
    )
    green_header_format = workbook.add_format(
        {"bg_color": "#B6E3A1", "border": 1, "align": "center", "valign": "vcenter"}
    )
    pink_header_format = workbook.add_format(
        {"bg_color": "#E7C0E5", "border": 1, "align": "center", "valign": "vcenter"}
    )
    isbn_format = workbook.add_format({"align": "center", "valign": "vcenter"})
    integer_format = workbook.add_format({"num_format": "#,##0;-#,##0;-", "align": "right"})
    green_total_format = workbook.add_format(
        {"num_format": "#,##0;-#,##0;-", "bg_color": "#B6E3A1", "align": "right"}
    )
    pink_total_format = workbook.add_format(
        {"num_format": "#,##0;-#,##0;-", "bg_color": "#E7C0E5", "align": "right"}
    )
    date_format = workbook.add_format({"num_format": "m/d/yyyy", "align": "center"})

    if group_df.empty:
        worksheet.write("A1", f"No data found for {group.name}", title_format)
        return

    metadata = family_metadata(group, group_df)
    isbn_order = metadata["ISBN"].tolist()
    matrix = family_matrix(group_df, isbn_order)

    for col_idx in range(len(METADATA_COLUMNS)):
        worksheet.write_blank(0, col_idx, None, gray_label_format)
    worksheet.write(0, 3, "Totals", gray_label_format)
    worksheet.write(2, 3, "OSD", osd_label_format)

    for col_idx, column_name in enumerate(DISPLAY_METADATA_HEADERS):
        worksheet.write(4, col_idx, column_name, metadata_header_format)

    for offset, row in enumerate(metadata.itertuples(index=False), start=len(METADATA_COLUMNS)):
        title_column_index = offset - len(METADATA_COLUMNS)
        total_format = green_total_format if title_column_index % 2 == 0 else pink_total_format
        header_format = green_header_format if title_column_index % 2 == 0 else pink_header_format
        worksheet.write_formula(0, offset, f"=SUM({xl_col(offset)}6:{xl_col(offset)}{len(matrix) + 5})", total_format)
        worksheet.write(1, offset, row.Title, title_format)
        if pd.isna(row.OSD):
            worksheet.write_blank(2, offset, None)
        else:
            worksheet.write_datetime(2, offset, row.OSD.to_pydatetime(), date_format)
        worksheet.write(3, offset, row.ISBN, isbn_format)
        worksheet.write(4, offset, row.ValueHeader, header_format)

    for row_idx, row in enumerate(matrix.itertuples(index=False), start=5):
        for col_idx, value in enumerate(row):
            if col_idx >= len(METADATA_COLUMNS):
                worksheet.write_number(row_idx, col_idx, float(value), integer_format)
            else:
                worksheet.write(row_idx, col_idx, value)

    last_row = max(len(matrix) + 4, 4)
    last_col = len(METADATA_COLUMNS) + len(isbn_order) - 1
    worksheet.autofilter(4, 0, last_row, last_col)
    worksheet.freeze_panes(5, 4)
    worksheet.hide_gridlines(2)
    worksheet.set_row(0, 21)
    worksheet.set_row(1, 61)
    worksheet.set_row(4, 20)
    worksheet.set_column(0, 0, 12)
    worksheet.set_column(1, 1, 28)
    worksheet.set_column(2, 2, 10)
    worksheet.set_column(3, 3, 30)
    worksheet.set_column(4, last_col, 14, integer_format)


def xl_col(col_idx: int) -> str:
    name = ""
    col_num = col_idx + 1
    while col_num:
        col_num, remainder = divmod(col_num - 1, 26)
        name = chr(65 + remainder) + name
    return name


def write_sql_sheet(
    workbook: Any,
    sales_query: str,
    orders_query: str,
    cached_sales_queries: list[tuple[str, str]],
) -> None:
    worksheet = workbook.add_worksheet("SQL")
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAD3"})
    body_format = workbook.add_format({"font_name": "Consolas", "font_size": 9})
    worksheet.set_column("A:A", 140)

    row_idx = 0
    queries = [
        ("Sales Query", sales_query),
        ("Hachette Orders Query", orders_query),
        *cached_sales_queries,
    ]
    for title, query in queries:
        worksheet.write(row_idx, 0, title, header_format)
        row_idx += 1
        for line in query.splitlines():
            worksheet.write(row_idx, 0, line, body_format)
            row_idx += 1
        row_idx += 2


def default_output_path() -> Path:
    process_paths.CROSS_GAP_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    run_date = datetime.now().strftime("%Y_%m_%d")
    return process_paths.CROSS_GAP_OUTPUT_DIR / f"Cross Gap {run_date}.xlsx"


def write_workbook(
    combined_df: pd.DataFrame,
    groups: list[TitleGroup],
    output_path: Path,
    sales_query: str,
    orders_query: str,
    cached_sales_queries: list[tuple[str, str]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        workbook = writer.book
        used_sheet_names: set[str] = set()
        for group in groups:
            group_df = combined_df[combined_df["Family"] == group.name].copy()
            write_group_sheet(writer, workbook, group, group_df, used_sheet_names)
        write_sql_sheet(workbook, sales_query, orders_query, cached_sales_queries)


def run(config_path: Path, output_path: Path | None = None, dry_run: bool = False) -> Path | None:
    groups = load_title_groups(config_path)
    sales_query = build_sales_query(groups)
    orders_query = build_orders_query(groups)
    cached_sales_queries = supplemental_sales_queries(groups)

    if dry_run:
        print("Sales query:")
        print(sales_query)
        print("\nHachette orders query:")
        print(orders_query)
        for title, query in cached_sales_queries:
            print(f"\n{title}:")
            print(query)
        return None

    print("Fetching Cross Gap sales data...")
    sales_df, orders_df = fetch_cross_gap_data(groups)
    print(f"Sales rows: {len(sales_df):,}")
    print(f"Order rows: {len(orders_df):,}")

    combined_df = combine_sales_and_orders(sales_df, orders_df)
    selected_output_path = output_path or default_output_path()
    write_workbook(combined_df, groups, selected_output_path, sales_query, orders_query, cached_sales_queries)
    print(f"Saved Cross Gap workbook: {selected_output_path}")
    return selected_output_path


def main(argv: list[str] | None = None) -> Path | None:
    parser = argparse.ArgumentParser(description="Build the Cross Gap sales and open-order workbook.")
    parser.add_argument(
        "command",
        nargs="?",
        choices=["menu", "run", "list"],
        default="menu",
        help="Use menu for the interactive Cross Gap menu, run to build the workbook, or list to show groupings.",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=process_paths.CROSS_GAP_CONFIG_FILE,
        help="Path to the title group JSON config.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Optional output workbook path. Defaults to Desktop/Cross Gap Reports.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print generated SQL without querying SQL Server or writing Excel.",
    )
    args = parser.parse_args(argv)
    if args.dry_run or args.output:
        return run(args.config, args.output, args.dry_run)
    if args.command == "menu":
        run_cross_gap_menu(args.config)
        return None
    if args.command == "list":
        print_groupings(args.config)
        return None
    return run(args.config, args.output, args.dry_run)


if __name__ == "__main__":
    main()
