from __future__ import annotations

import argparse
import shutil
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from sqlalchemy import create_engine

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

AMAZON_AMS_DIR = REPO_ROOT / "amazon_ams"
if str(AMAZON_AMS_DIR) not in sys.path:
    sys.path.insert(0, str(AMAZON_AMS_DIR))

AMAZON_SQL_UPLOAD_DIR = REPO_ROOT / "amazon_sql_upload"
if str(AMAZON_SQL_UPLOAD_DIR) not in sys.path:
    sys.path.insert(0, str(AMAZON_SQL_UPLOAD_DIR))

from shared.amazon_metadata import resolve_isbn_series  # noqa: E402
from loader_asin_mapping import load_asin_mapping  # noqa: E402
from load_ypticod import load_ypticod  # noqa: E402

CAMPAIGN_FOLDER = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\AMS_Monthly_Campaign"
)
CATALOG_FILE = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\Catalog\Catalog_Manufacturing_UnitedStates.csv"
)
AMS_HISTORY_PICKLE = REPO_ROOT / "amazon_ams" / "combined_amazon_ads_by_asin.pkl"
AMS_HISTORY_PARQUET = CAMPAIGN_FOLDER / "cache" / "ams_monthly_campaigns.parquet"
AMS_AGGREGATE_EXCEL = Path(
    r"F:\ANALYSIS\Finance\DataWarehouse\Atelier Amazon\ams_summaries\combined_amazon_ads_by_asin.xlsx"
)
AMS_AGGREGATE_PICKLE = REPO_ROOT / "amazon_ams" / "combined_amazon_ads_by_asin.pkl"
FINAL_REPORTS_FOLDER = Path(
    r"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2026"
)
DEFAULT_OUTPUT_SUFFIX = "_summary.xlsx"
CAMPAIGN_NAMES = {
    "Chronicle Books MAIN": "CB",
    "Quadrille": "QD",
    "Twirl": "TW",
    "Hardie Grant": "HG",
}
PUBLISHER_DISPLAY_NAMES = {
    "Quadrille Publishing Limited": "Quadrille",
}
PUBLISHER_OVERRIDES = {
    "Princeton": ("Chronicle", "CPA"),
}
PGRP_DISPLAY_NAMES = {
    "SC": "SIE",
    "TW": "TOU",
}
PUBLISHER_REPORTS = {
    "PWP": ("Publisher", "Post Wave"),
    "HG": ("Publisher", "Hardie Grant"),
    "QD": ("Publisher", "Quadrille"),
    "TW": ("pgrp", "TOU"),
}
SOURCE_NUMERIC_COLUMNS = [
    "Impressions",
    "Clicks",
    "Units sold",
    "Total cost",
    "Sales",
]
MISSING_ASIN_DISPLAY_COLUMNS = [
    "Campaign name",
    "Advertised product ID",
    "ASIN",
    *SOURCE_NUMERIC_COLUMNS,
]
OUTPUT_COLUMNS = [
    "ASIN",
    "ISBN",
    "Title",
    "Publisher",
    "pgrp",
    "osd",
    "Impressions",
    "Clicks",
    "CTR",
    "Units sold",
    "CVR",
    "Spend",
    "Sales",
    "ACOS",
    "ROAS",
    "Count of Campaigns",
    "Format",
    "campaign",
]
HISTORY_COLUMNS = ["period", "source_file", *OUTPUT_COLUMNS]
LEGACY_AGGREGATE_COLUMNS = [
    "ASIN",
    "impressions",
    "clicks",
    "spend",
    "14 day total sales",
    "units sold",
    "count of campaigns",
    "ISBN",
    "CTR",
    "CRV",
    "ACOS",
    "ROAS",
    "period",
    "source_file",
    "title",
    "publisher",
    "pgrp",
    "PT",
    "OSD",
]
COMPARISON_METRICS = [
    "Impressions",
    "Clicks",
    "CTR",
    "Units sold",
    "CVR",
    "Spend",
    "Sales",
    "ACOS",
    "ROAS",
    "Count of Campaigns",
]
COMPARISON_VALUE_COLUMNS = ["Publisher", "pgrp", *COMPARISON_METRICS]
CAMPAIGN_ANALYSIS_METRICS = [
    "Impressions",
    "Clicks",
    "CTR",
    "Units sold",
    "CVR",
    "Spend",
    "Sales",
    "ACOS",
    "ROAS",
    "ISBN Count",
]
CAMPAIGN_ANALYSIS_COLUMNS = [
    "Publisher",
    "Campaign ID",
    "Campaign name",
    *CAMPAIGN_ANALYSIS_METRICS,
]
DEFAULT_MISSING_ASIN_OVERRIDES = {
    "TWL | BOOKS | Ultimate Series | SB | BRAND | KW": ["B08RC8VVDF", "B078X9MFG5", "2848019425"],
    "TWL | BOOKS | Top Sellers | SB | BRAND | KW": ["2408012856", "B08CG7LM5B", "2848019425"],
    "TWL | BOOKS | Top Sellers | SB | NON-BRAND | KW 1": ["2408012856", "B08CG7LM5B", "2848019425"],
    "TWL | BOOKS | Ultimate Series | SB | NON-BRAND | KW": ["B08RC8VVDF", "B078X9MFG5", "2848019425"],
    "CB | Books | Construction Site: Garbage Crew to the Rescue | 179722655X | Competitor | CAT | SBV": ["179722655X"],
    "S: Ivy and Bean - B": ["0811864952", "B006P0AIMO", "1452117322"],
    "C: Kids Valentine's - BI": ["1452139970", "1797204300", "1452184895"],
    "S: Greek Myths - B": ["1452178917", "1797201867", "1797207075"],
    "I: Desi Bakes - BV": ["1958417319"],
    "S: Ultimate - B (D)": ["2848019425", "2848019840", "B00T3CV2VW"],
    "C: This is Home, Still, Style - B": ["1743793456", "174379570X", "1743797974"],
    "G: Relationships Birthday - M": [
        "0811870545",
        "1797212478",
        "1452173184",
        "1452154759",
        "1452155380",
        "1797215876",
        "145214124X",
        "1452183007",
        "1797202812",
        "1452177112",
    ],
}


def normalize_asin(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text.zfill(10) if text else ""


def normalize_isbn(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    digits = "".join(char for char in text if char.isdigit())
    if not digits:
        return ""
    if len(digits) < 13:
        return digits.zfill(13)
    if len(digits) > 13 and digits.startswith("0"):
        return digits[-13:]
    return digits[:13]


def clean_numeric(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype("string")
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def parse_period_from_filename(source_file: Path) -> str:
    match = re.search(r"(?P<period>20\d{4})", source_file.stem)
    if not match:
        raise ValueError(f"Could not parse YYYYMM period from file name: {source_file.name}")
    period = match.group("period")
    return f"{period[:4]}-{period[4:]}"


def previous_period(period: str) -> str:
    timestamp = pd.Period(period, freq="M") - 1
    return str(timestamp)


def period_compact(period: str) -> str:
    return period.replace("-", "")


def final_report_file(period: str, suffix: str = "ALL") -> Path:
    return FINAL_REPORTS_FOLDER / f"{period_compact(period)}_AMS_Performance_by_ASIN_{suffix}.xlsx"


def campaign_file_sort_key(path: Path) -> tuple[int, str, float]:
    try:
        return (1, parse_period_from_filename(path), path.stat().st_mtime)
    except ValueError:
        return (0, "", path.stat().st_mtime)


def choose_campaign_file(folder: Path = CAMPAIGN_FOLDER) -> Path:
    files = sorted(folder.glob("*.csv"), key=campaign_file_sort_key, reverse=True)
    if not files:
        raise FileNotFoundError(f"No CSV files found in {folder}")

    print("Available AMS monthly campaign CSV files:")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name}")
    choice = input("Choose the current month's file: ").strip()
    try:
        selected = files[int(choice) - 1]
    except (ValueError, IndexError) as exc:
        raise ValueError("Invalid file selection.") from exc
    return selected


def read_campaign_csv(source_file: Path) -> pd.DataFrame:
    df = pd.read_csv(source_file, dtype=object)
    df.columns = [str(column).strip() for column in df.columns]
    required = [
        "Advertiser account name",
        "Advertised product ID",
        "Campaign ID",
        *SOURCE_NUMERIC_COLUMNS,
    ]
    missing = [column for column in required if column not in df.columns]
    if missing:
        raise ValueError(f"Campaign CSV is missing required columns: {', '.join(missing)}")

    df = df.copy()
    df["ASIN"] = df["Advertised product ID"].map(normalize_asin)
    for column in SOURCE_NUMERIC_COLUMNS:
        df[column] = clean_numeric(df[column])
    return df


def parse_missing_asin_overrides(values: list[str] | None) -> dict[str, list[str]]:
    overrides: dict[str, list[str]] = {
        campaign: [normalize_asin(asin) for asin in asins]
        for campaign, asins in DEFAULT_MISSING_ASIN_OVERRIDES.items()
    }
    for value in values or []:
        if "=" not in value:
            raise ValueError(
                "Missing ASIN override must use Campaign name=ASIN1,ASIN2 format."
            )
        campaign_name, asin_values = value.split("=", 1)
        asins = [normalize_asin(asin) for asin in asin_values.split(",") if asin.strip()]
        if not asins:
            raise ValueError(f"No ASINs supplied for campaign override: {campaign_name}")
        overrides[campaign_name.strip()] = asins
    return overrides


def missing_asins_for_campaign(campaign_rows: pd.DataFrame, campaign_asins: list[str]) -> list[str]:
    campaign_asins = list(dict.fromkeys(asin for asin in campaign_asins if asin))
    represented_asins = set(campaign_rows["ASIN"].astype("string").fillna("").str.strip())
    represented_asins.discard("")
    missing_asins = [asin for asin in campaign_asins if asin not in represented_asins]
    return missing_asins or campaign_asins


def resolve_missing_asins(
    df: pd.DataFrame,
    *,
    prompt: bool,
    overrides: dict[str, list[str]] | None = None,
) -> pd.DataFrame:
    missing_mask = df["ASIN"].eq("")
    if not missing_mask.any():
        return df

    merged_overrides = {
        campaign: [normalize_asin(asin) for asin in asins]
        for campaign, asins in DEFAULT_MISSING_ASIN_OVERRIDES.items()
    }
    merged_overrides.update(overrides or {})
    overrides = merged_overrides
    resolved_rows: list[pd.DataFrame] = []
    rows_to_drop: list[int] = []

    for row_index, row in df[missing_mask].iterrows():
        campaign_name = normalize_text(row["Campaign name"])
        campaign_rows = df[df["Campaign name"].map(normalize_text).eq(campaign_name)].copy()
        print()
        print(f'Missing ASIN found for campaign: "{campaign_name}"')
        print("Rows currently in this campaign:")
        print(campaign_rows[MISSING_ASIN_DISPLAY_COLUMNS].to_string(index=False))
        print()

        replacement_asins = overrides.get(campaign_name)
        if replacement_asins is None:
            if not prompt:
                raise ValueError(
                    f'Missing ASINs for campaign "{campaign_name}". '
                    "Rerun interactively or pass --missing-asin "
                    f'"{campaign_name}=ASIN1,ASIN2".'
                )
            response = input(
                "Enter the campaign ASINs; already represented ASINs will be skipped, comma-separated: "
            ).strip()
            replacement_asins = [
                normalize_asin(value) for value in response.split(",") if value.strip()
            ]

        replacement_asins = missing_asins_for_campaign(campaign_rows, replacement_asins)
        if not replacement_asins:
            raise ValueError(
                f'No missing replacement ASINs remain for campaign "{campaign_name}" '
                "after removing ASINs already represented in that campaign."
            )

        split_count = len(replacement_asins)
        for asin in replacement_asins:
            replacement = row.copy()
            replacement["Advertised product ID"] = asin
            replacement["Advertised product SKU"] = asin
            replacement["ASIN"] = asin
            for column in SOURCE_NUMERIC_COLUMNS:
                replacement[column] = replacement[column] / split_count
            resolved_rows.append(pd.DataFrame([replacement]))
        rows_to_drop.append(row_index)

        print(
            f'Replaced blank ASIN row for "{campaign_name}" with '
            f"{split_count} split row(s): {', '.join(replacement_asins)}"
        )

    output = df.drop(index=rows_to_drop).copy()
    if resolved_rows:
        output = pd.concat([output, *resolved_rows], ignore_index=True)
    return output


def normalize_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def print_advertiser_summary(df: pd.DataFrame) -> None:
    summary = (
        df.groupby("Advertiser account name", dropna=False)
        .agg(
            {
                "Impressions": "sum",
                "Clicks": "sum",
                "Units sold": "sum",
                "Total cost": "sum",
                "Sales": "sum",
                "Campaign ID": "count",
            }
        )
        .rename(columns={"Campaign ID": "Campaign ID Count"})
        .reset_index()
    )
    print()
    print("AMS monthly campaign summary by advertiser:")
    print(summary.to_string(index=False))
    print()


def load_asin_isbn_map() -> pd.DataFrame:
    mapping_parts: list[pd.DataFrame] = []

    ypticod = load_ypticod().copy()
    ypticod["ASIN"] = ypticod["ASIN"].map(normalize_asin)
    ypticod["ISBN"] = ypticod["ISBN"].map(normalize_isbn)
    mapping_parts.append(ypticod[["ASIN", "ISBN"]])

    catalog = load_catalog_asin_map()
    if not catalog.empty:
        mapping_parts.append(catalog[["ASIN", "ISBN"]])

    asin_mapping = load_asin_mapping().copy()
    asin_mapping["ASIN"] = asin_mapping["ASIN"].map(normalize_asin)
    asin_mapping["ISBN"] = asin_mapping["ISBN"].map(normalize_isbn)
    mapping_parts.append(asin_mapping[["ASIN", "ISBN"]])

    mapping = pd.concat(mapping_parts, ignore_index=True)
    mapping = mapping[(mapping["ASIN"] != "") & (mapping["ISBN"] != "")]
    mapping = mapping.drop_duplicates(subset=["ASIN"], keep="first")
    return mapping


def load_catalog_asin_map(catalog_file: Path = CATALOG_FILE) -> pd.DataFrame:
    if not catalog_file.exists():
        print(f"Warning: Amazon catalog not found for ASIN fallback: {catalog_file}")
        return pd.DataFrame(columns=["ASIN", "ISBN"])

    catalog = pd.read_csv(
        catalog_file,
        skiprows=1,
        dtype=object,
        usecols=["ASIN", "EAN", "ISBN", "Model Number"],
    )
    catalog["ASIN"] = catalog["ASIN"].map(normalize_asin)
    for column in ["ISBN", "EAN", "Model Number"]:
        catalog[column] = catalog[column].map(normalize_isbn)

    isbn = catalog["ISBN"].copy()
    for fallback_column in ["EAN", "Model Number"]:
        isbn = isbn.mask(isbn.eq(""), catalog[fallback_column])
    catalog["ISBN"] = isbn
    catalog = catalog[(catalog["ASIN"] != "") & (catalog["ISBN"] != "")]
    return catalog[["ASIN", "ISBN"]].drop_duplicates(subset=["ASIN"], keep="first")


def item_metadata_sql() -> str:
    return """
    SELECT
        i.ISBN
        ,i.SHORT_TITLE Title
        ,i.PUBLISHER_CODE Publisher
        ,CASE WHEN left(i.PUBLISHING_GROUP,3) = 'BAR' THEN 'BAR' ELSE i.PUBLISHING_GROUP END pgrp
        ,convert(char,coalesce(convert(varchar,i.AMORTIZATION_DATE,101),shdt.shipdate),101) [osd]
        ,i.PRODUCT_TYPE_DESCRIPTION [Format]
    FROM
        ebs.item i
        left join (
            SELECT [ISBN],[SHIPDATE]
            FROM [CBQ2].[pm].[ItemInfo] ii
            WHERE ii.IMPRESSION = 1 and ii.SHIPDATE is not null
        ) shdt on shdt.ISBN = i.ISBN
    WHERE i.SHORT_TITLE is not null and i.ISBN is not null
    """


def load_item_metadata() -> pd.DataFrame:
    engine = create_engine("mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server")
    with engine.connect() as connection:
        metadata = pd.read_sql_query(item_metadata_sql(), connection)
    metadata["ISBN"] = metadata["ISBN"].map(normalize_isbn)
    metadata = metadata.drop_duplicates(subset=["ISBN"], keep="first")
    return metadata


def safe_divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    denominator = denominator.replace(0, np.nan)
    return (numerator / denominator).replace([np.inf, -np.inf], np.nan).fillna(0)


def build_summary(
    df: pd.DataFrame,
    mapping: pd.DataFrame | None = None,
    metadata: pd.DataFrame | None = None,
) -> pd.DataFrame:
    mapping = load_asin_isbn_map() if mapping is None else mapping
    metadata = load_item_metadata() if metadata is None else metadata

    working = df.merge(mapping, on="ASIN", how="left")
    resolved = resolve_isbn_series(
        working, metadata, ["ISBN"]
    )
    working["ISBN"] = resolved.fillna(working["ISBN"])
    working["ISBN"] = working["ISBN"].map(normalize_isbn)
    working["ISBN"] = working["ISBN"].replace("", "NO ISBN")
    working["campaign"] = (
        working["Advertiser account name"]
        .map(CAMPAIGN_NAMES)
        .fillna(working["Advertiser account name"].astype("string").str.strip())
    )

    aggregated = (
        working.groupby(["ASIN", "ISBN", "campaign"], dropna=False)
        .agg(
            {
                "Impressions": "sum",
                "Clicks": "sum",
                "Units sold": "sum",
                "Total cost": "sum",
                "Sales": "sum",
                "Campaign ID": "count",
            }
        )
        .rename(columns={"Total cost": "Spend", "Campaign ID": "Count of Campaigns"})
        .reset_index()
    )
    aggregated["CTR"] = safe_divide(aggregated["Clicks"], aggregated["Impressions"])
    aggregated["CVR"] = safe_divide(aggregated["Units sold"], aggregated["Clicks"])
    aggregated["ACOS"] = safe_divide(aggregated["Spend"], aggregated["Sales"])
    aggregated["ROAS"] = safe_divide(aggregated["Sales"], aggregated["Spend"])

    summary = aggregated.merge(metadata, on="ISBN", how="left")
    summary = apply_publisher_overrides(summary)
    summary = summary[OUTPUT_COLUMNS].sort_values(["campaign", "Title", "ASIN"], kind="stable")
    return summary


def build_campaign_detail(
    df: pd.DataFrame,
    mapping: pd.DataFrame | None = None,
    metadata: pd.DataFrame | None = None,
) -> pd.DataFrame:
    mapping = load_asin_isbn_map() if mapping is None else mapping
    metadata = load_item_metadata() if metadata is None else metadata

    working = df.merge(mapping, on="ASIN", how="left")
    resolved = resolve_isbn_series(
        working, metadata, ["ISBN"]
    )
    working["ISBN"] = resolved.fillna(working["ISBN"])
    working["ISBN"] = working["ISBN"].map(normalize_isbn)
    working["ISBN"] = working["ISBN"].replace("", "NO ISBN")
    working = working.merge(metadata, on="ISBN", how="left")
    working = apply_publisher_overrides(working)
    working["Spend"] = working["Total cost"]
    for column in ["Campaign ID", "Campaign name", "Publisher", "pgrp"]:
        if column not in working.columns:
            working[column] = ""
        working[column] = working[column].astype("string").fillna("").str.strip()
    return working[
        [
            "ASIN",
            "ISBN",
            "Publisher",
            "pgrp",
            "Campaign ID",
            "Campaign name",
            "Impressions",
            "Clicks",
            "Units sold",
            "Spend",
            "Sales",
        ]
    ].copy()


def clean_group_value(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def apply_publisher_overrides(summary: pd.DataFrame) -> pd.DataFrame:
    output = summary.copy()
    output["Publisher"] = output["Publisher"].map(
        lambda value: PUBLISHER_DISPLAY_NAMES.get(clean_group_value(value), clean_group_value(value))
    )
    output["pgrp"] = output["pgrp"].map(
        lambda value: PGRP_DISPLAY_NAMES.get(clean_group_value(value), clean_group_value(value))
    )
    for publisher_name, (replacement_publisher, replacement_pgrp) in PUBLISHER_OVERRIDES.items():
        mask = output["Publisher"].astype("string").str.strip().eq(publisher_name)
        output.loc[mask, "Publisher"] = replacement_publisher
        output.loc[mask, "pgrp"] = replacement_pgrp
    quadrille_mask = output["Publisher"].astype("string").str.strip().eq("Quadrille")
    output.loc[quadrille_mask, "pgrp"] = "QD"
    return output


def normalize_publisher_group(publisher: object, pgrp: object) -> tuple[str, str]:
    publisher_text = clean_group_value(publisher)
    pgrp_text = clean_group_value(pgrp)
    publisher_text = PUBLISHER_DISPLAY_NAMES.get(publisher_text, publisher_text)
    pgrp_text = PGRP_DISPLAY_NAMES.get(pgrp_text, pgrp_text)
    publisher_text, pgrp_text = PUBLISHER_OVERRIDES.get(
        publisher_text, (publisher_text, pgrp_text)
    )
    if publisher_text == "Quadrille":
        pgrp_text = "QD"
    return publisher_text, pgrp_text


def comparison_group_label(publisher: object, pgrp: object) -> str:
    publisher_text, pgrp_text = normalize_publisher_group(publisher, pgrp)
    if publisher_text.casefold() == "chronicle":
        return pgrp_text or "Chronicle"
    if publisher_text:
        return f"{publisher_text} - {pgrp_text}" if pgrp_text else publisher_text
    return pgrp_text or "Unknown"


def summarize_for_comparison(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=COMPARISON_VALUE_COLUMNS)

    working = df.copy()
    for column in ["Publisher", "pgrp"]:
        if column not in working.columns:
            working[column] = ""
    normalized_groups = working.apply(
        lambda row: normalize_publisher_group(row["Publisher"], row["pgrp"]),
        axis=1,
        result_type="expand",
    )
    working["Publisher"] = normalized_groups[0]
    working["pgrp"] = normalized_groups[1]
    working["Publisher"] = working.apply(
        lambda row: comparison_group_label(row["Publisher"], row["pgrp"])
        if not row["Publisher"] and row["pgrp"]
        else row["Publisher"] or "Unknown",
        axis=1,
    )
    working["pgrp"] = working.apply(
        lambda row: row["pgrp"] if row["pgrp"] and row["pgrp"] != row["Publisher"] else "",
        axis=1,
    )
    grouped = (
        working.groupby(["Publisher", "pgrp"], dropna=False)
        .agg(
            {
                "Impressions": "sum",
                "Clicks": "sum",
                "Units sold": "sum",
                "Spend": "sum",
                "Sales": "sum",
                "Count of Campaigns": "sum",
            }
        )
        .reset_index()
    )
    grouped["CTR"] = safe_divide(grouped["Clicks"], grouped["Impressions"])
    grouped["CVR"] = safe_divide(grouped["Units sold"], grouped["Clicks"])
    grouped["ACOS"] = safe_divide(grouped["Spend"], grouped["Sales"])
    grouped["ROAS"] = safe_divide(grouped["Sales"], grouped["Spend"])
    return grouped[COMPARISON_VALUE_COLUMNS].sort_values(["Publisher", "pgrp"], kind="stable")


def load_prior_month_summary(period: str, history_file: Path = AMS_HISTORY_PICKLE) -> pd.DataFrame:
    if AMS_HISTORY_PARQUET.exists():
        history = pd.read_parquet(AMS_HISTORY_PARQUET)
        history = history[history["period"].astype(str).eq(period)].copy()
        if not history.empty:
            return summarize_for_comparison(history)

    if not history_file.exists():
        print(f"Warning: AMS history pickle not found for monthly comparison: {history_file}")
        return pd.DataFrame(columns=COMPARISON_VALUE_COLUMNS)

    history = pd.read_pickle(history_file)
    history = history[history["period"].astype(str).eq(period)].copy()
    if history.empty:
        print(f"Warning: no AMS history rows found for prior period {period}.")
        return pd.DataFrame(columns=COMPARISON_VALUE_COLUMNS)

    renamed = history.rename(
        columns={
            "impressions": "Impressions",
            "clicks": "Clicks",
            "spend": "Spend",
            "14 day total sales": "Sales",
            "units sold": "Units sold",
            "count of campaigns": "Count of Campaigns",
            "publisher": "Publisher",
        }
    )
    if "Count of Campaigns" not in working.columns:
        working["Count of Campaigns"] = 1
    for column in ["Impressions", "Clicks", "Units sold", "Spend", "Sales", "Count of Campaigns"]:
        renamed[column] = pd.to_numeric(renamed[column], errors="coerce").fillna(0)
    return summarize_for_comparison(renamed)


def build_monthly_comparison(
    summary: pd.DataFrame,
    current_period: str,
    prior_summary: pd.DataFrame | None = None,
    prior_period: str | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, str]:
    prior_period = prior_period or previous_period(current_period)
    current = summarize_for_comparison(summary)
    prior = summarize_for_comparison(prior_summary) if prior_summary is not None else load_prior_month_summary(prior_period)
    group_keys = (
        pd.concat([current[["Publisher", "pgrp"]], prior[["Publisher", "pgrp"]]], ignore_index=True)
        .drop_duplicates()
        .sort_values(["Publisher", "pgrp"], kind="stable")
    )
    current = group_keys.merge(current, on=["Publisher", "pgrp"], how="left").fillna(0)
    prior = group_keys.merge(prior, on=["Publisher", "pgrp"], how="left").fillna(0)
    variance = current.copy()
    for column in COMPARISON_METRICS:
        variance[column] = current[column] - prior[column]
    return current, prior, variance, prior_period


def summarize_for_campaign_analysis(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=CAMPAIGN_ANALYSIS_COLUMNS)

    working = df.copy()
    for column in ["Publisher", "pgrp", "Campaign ID", "Campaign name", "ISBN"]:
        if column not in working.columns:
            working[column] = ""
    normalized_groups = working.apply(
        lambda row: normalize_publisher_group(row["Publisher"], row["pgrp"]),
        axis=1,
        result_type="expand",
    )
    working["Publisher"] = normalized_groups[0]
    working["Campaign ID"] = working["Campaign ID"].astype("string").fillna("").str.strip()
    working["Campaign name"] = working["Campaign name"].astype("string").fillna("").str.strip()
    if "Count of Campaigns" not in working.columns:
        working["Count of Campaigns"] = 1
    for column in ["Impressions", "Clicks", "Units sold", "Spend", "Sales", "Count of Campaigns"]:
        working[column] = pd.to_numeric(working[column], errors="coerce").fillna(0)

    grouped = (
        working.groupby(["Publisher", "Campaign ID", "Campaign name"], dropna=False)
        .agg(
            {
                "Impressions": "sum",
                "Clicks": "sum",
                "Units sold": "sum",
                "Spend": "sum",
                "Sales": "sum",
                "Count of Campaigns": "sum",
                "ISBN": lambda values: values[values.astype(str).ne("NO ISBN")].nunique(),
            }
        )
        .rename(columns={"ISBN": "ISBN Count"})
        .reset_index()
    )
    grouped["CTR"] = safe_divide(grouped["Clicks"], grouped["Impressions"])
    grouped["CVR"] = safe_divide(grouped["Units sold"], grouped["Clicks"])
    grouped["ACOS"] = safe_divide(grouped["Spend"], grouped["Sales"])
    grouped["ROAS"] = safe_divide(grouped["Sales"], grouped["Spend"])
    return grouped[CAMPAIGN_ANALYSIS_COLUMNS].sort_values(
        ["Publisher", "Spend"], ascending=[True, False], kind="stable"
    )


def load_prior_campaign_analysis_summary(period: str) -> pd.DataFrame:
    source_file = CAMPAIGN_FOLDER / f"{period_compact(period)}_MonthlyCampaigns.csv"
    if not source_file.exists():
        return pd.DataFrame(columns=CAMPAIGN_ANALYSIS_COLUMNS)
    df = read_campaign_csv(source_file)
    df = resolve_missing_asins(df, prompt=False)
    detail = build_campaign_detail(df)
    return summarize_for_campaign_analysis(detail)


def build_campaign_analysis(
    summary: pd.DataFrame,
    current_period: str,
    prior_summary: pd.DataFrame | None = None,
    prior_period: str | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, str]:
    prior_period = prior_period or previous_period(current_period)
    current = summarize_for_campaign_analysis(summary)
    prior = (
        summarize_for_campaign_analysis(prior_summary)
        if prior_summary is not None
        else load_prior_campaign_analysis_summary(prior_period)
    )
    group_keys = (
        pd.concat(
            [
                current[["Publisher", "Campaign ID", "Campaign name", "Spend"]],
                prior[["Publisher", "Campaign ID", "Campaign name", "Spend"]],
            ],
            ignore_index=True,
        )
        .drop_duplicates(subset=["Publisher", "Campaign ID", "Campaign name"], keep="first")
    )
    current_spend = current[["Publisher", "Campaign ID", "Campaign name", "Spend"]].rename(
        columns={"Spend": "_CurrentSpendSort"}
    )
    group_keys = group_keys.drop(columns="Spend").merge(
        current_spend,
        on=["Publisher", "Campaign ID", "Campaign name"],
        how="left",
    )
    group_keys["_CurrentSpendSort"] = group_keys["_CurrentSpendSort"].fillna(0)
    group_keys = group_keys.sort_values(
        ["Publisher", "_CurrentSpendSort"], ascending=[True, False], kind="stable"
    ).drop(columns="_CurrentSpendSort")
    current = group_keys.merge(current, on=["Publisher", "Campaign ID", "Campaign name"], how="left").fillna(0)
    prior = group_keys.merge(prior, on=["Publisher", "Campaign ID", "Campaign name"], how="left").fillna(0)
    variance = current.copy()
    for column in CAMPAIGN_ANALYSIS_METRICS:
        variance[column] = current[column] - prior[column]
    return current, prior, variance, prior_period


def comparison_display_rows(section_df: pd.DataFrame) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    if section_df.empty:
        return rows

    for publisher_index, (publisher, publisher_df) in enumerate(section_df.groupby("Publisher", sort=False, dropna=False)):
        if publisher_index:
            rows.append({"Label": "", "RowType": "spacer"})
        publisher_df = publisher_df.copy()
        total_row: dict[str, object] = {"Label": f"{publisher} Total", "RowType": "publisher"}
        for metric in COMPARISON_METRICS:
            if metric in {"CTR", "CVR", "ACOS", "ROAS"}:
                continue
            total_row[metric] = publisher_df[metric].sum()
        total_row["CTR"] = total_row["Clicks"] / total_row["Impressions"] if total_row["Impressions"] else 0
        total_row["CVR"] = total_row["Units sold"] / total_row["Clicks"] if total_row["Clicks"] else 0
        total_row["ACOS"] = total_row["Spend"] / total_row["Sales"] if total_row["Sales"] else 0
        total_row["ROAS"] = total_row["Sales"] / total_row["Spend"] if total_row["Spend"] else 0
        rows.append(total_row)

        detail_rows = publisher_df[publisher_df["pgrp"].astype(str).str.strip().ne("")]
        if detail_rows.empty:
            continue
        for _, row in detail_rows.iterrows():
            detail_row: dict[str, object] = {"Label": row["pgrp"], "RowType": "pgrp"}
            for metric in COMPARISON_METRICS:
                detail_row[metric] = row[metric]
            rows.append(detail_row)
    return rows


def write_comparison_section(
    worksheet,
    section_df: pd.DataFrame,
    *,
    start_row: int,
    start_col: int,
    title: str,
    formats: dict[str, object],
    section_style: str,
) -> None:
    total_row = start_row
    title_row = start_row + 2
    header_row = start_row + 3
    data_start_row = start_row + 4
    last_section_col = start_col + len(COMPARISON_METRICS)
    display_rows = comparison_display_rows(section_df)

    worksheet.write(total_row, start_col, "GRAND TOTALS", formats["section_title"])
    worksheet.write(title_row, start_col, title, formats["section_title_center"])
    for col in range(start_col + 1, last_section_col + 1):
        worksheet.write_blank(title_row, col, None, formats["section_title_center"])

    headers = ["Publisher / pgrp", *COMPARISON_METRICS]
    for offset, header in enumerate(headers):
        worksheet.write(header_row, start_col + offset, header, formats["header"])

    publisher_rows: list[int] = []
    for row_offset, row in enumerate(display_rows):
        excel_row = data_start_row + row_offset
        if row["RowType"] == "spacer":
            for offset in range(0, len(COMPARISON_METRICS) + 1):
                worksheet.write_blank(excel_row, start_col + offset, None)
            continue

        row_format = (
            formats[f"{section_style}_publisher_row"]
            if row["RowType"] == "publisher"
            else formats[f"{section_style}_detail_row"]
        )
        if row["RowType"] == "publisher":
            publisher_rows.append(excel_row)
        worksheet.write(excel_row, start_col, row["Label"], row_format)
        for metric_offset, metric in enumerate(COMPARISON_METRICS, start=1):
            cell_format = metric_format(metric, formats, row_type=row["RowType"], section_style=section_style)
            worksheet.write(excel_row, start_col + metric_offset, row[metric], cell_format)

    first_data_excel_row = data_start_row + 1
    last_data_excel_row = data_start_row + len(display_rows)
    for metric_offset, metric in enumerate(COMPARISON_METRICS, start=1):
        col = start_col + metric_offset
        col_letter = xl_col(col)
        cell_format = metric_format(metric, formats, row_type="publisher", section_style=section_style)
        if metric in {"CTR", "CVR", "ACOS", "ROAS"}:
            formula = comparison_ratio_formula(metric, total_row + 1, start_col)
            worksheet.write_formula(total_row, col, formula, cell_format)
        else:
            publisher_refs = ",".join(f"{col_letter}{row_idx + 1}" for row_idx in publisher_rows)
            total_formula = f"=SUM({publisher_refs})" if publisher_refs else "=0"
            worksheet.write_formula(
                total_row,
                col,
                total_formula,
                cell_format,
            )


def comparison_ratio_formula(metric: str, excel_row: int, start_col: int) -> str:
    positions = {metric_name: start_col + 1 + idx for idx, metric_name in enumerate(COMPARISON_METRICS)}
    if metric == "CTR":
        numerator, denominator = "Clicks", "Impressions"
    elif metric == "CVR":
        numerator, denominator = "Units sold", "Clicks"
    elif metric == "ACOS":
        numerator, denominator = "Spend", "Sales"
    elif metric == "ROAS":
        numerator, denominator = "Sales", "Spend"
    else:
        return ""
    return f"=IFERROR({xl_col(positions[numerator])}{excel_row}/{xl_col(positions[denominator])}{excel_row},0)"


def metric_format(metric: str, formats: dict[str, object], row_type: str = "pgrp", section_style: str = "current"):
    prefix = f"{section_style}_publisher_" if row_type == "publisher" else f"{section_style}_detail_"
    if metric in {"CTR", "CVR", "ACOS"}:
        return formats[f"{prefix}percent"]
    if metric == "ROAS":
        return formats[f"{prefix}number"]
    if metric in {"Spend", "Sales"}:
        return formats[f"{prefix}money"]
    return formats[f"{prefix}integer"]


def write_monthly_comparison_sheet(
    writer: pd.ExcelWriter,
    summary: pd.DataFrame,
    current_period: str,
    formats: dict[str, object],
    prior_summary: pd.DataFrame | None = None,
    prior_period: str | None = None,
) -> None:
    current, prior, variance, prior_period = build_monthly_comparison(
        summary,
        current_period,
        prior_summary=prior_summary,
        prior_period=prior_period,
    )
    worksheet = writer.book.add_worksheet("Monthly Comparisons")
    writer.sheets["Monthly Comparisons"] = worksheet

    write_comparison_section(
        worksheet,
        current,
        start_row=0,
        start_col=0,
        title=pd.Period(current_period, freq="M").strftime("%B %Y"),
        formats=formats,
        section_style="current",
    )
    write_comparison_section(
        worksheet,
        prior,
        start_row=0,
        start_col=12,
        title=pd.Period(prior_period, freq="M").strftime("%B %Y"),
        formats=formats,
        section_style="prior",
    )
    write_comparison_section(
        worksheet,
        variance,
        start_row=0,
        start_col=24,
        title="Variance",
        formats=formats,
        section_style="variance",
    )
    worksheet.freeze_panes(5, 1)
    for col in [0, 12, 24]:
        worksheet.set_column(col, col, 18)
        worksheet.set_column(col + 1, col + len(COMPARISON_METRICS), 10)


def campaign_analysis_display_rows(section_df: pd.DataFrame) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    if section_df.empty:
        return rows

    for publisher_index, (publisher, publisher_df) in enumerate(section_df.groupby("Publisher", sort=False, dropna=False)):
        if publisher_index:
            rows.append({"Campaign ID": "", "Campaign name": "", "RowType": "spacer"})
        publisher_df = publisher_df.copy()
        total_row: dict[str, object] = {
            "Campaign ID": "",
            "Campaign name": f"{publisher} Total",
            "RowType": "publisher",
        }
        for metric in CAMPAIGN_ANALYSIS_METRICS:
            if metric in {"CTR", "CVR", "ACOS", "ROAS"}:
                continue
            total_row[metric] = publisher_df[metric].sum()
        total_row["CTR"] = total_row["Clicks"] / total_row["Impressions"] if total_row["Impressions"] else 0
        total_row["CVR"] = total_row["Units sold"] / total_row["Clicks"] if total_row["Clicks"] else 0
        total_row["ACOS"] = total_row["Spend"] / total_row["Sales"] if total_row["Sales"] else 0
        total_row["ROAS"] = total_row["Sales"] / total_row["Spend"] if total_row["Spend"] else 0
        rows.append(total_row)

        for _, row in publisher_df.iterrows():
            detail_row: dict[str, object] = {
                "Campaign ID": row["Campaign ID"],
                "Campaign name": row["Campaign name"],
                "RowType": "campaign",
            }
            for metric in CAMPAIGN_ANALYSIS_METRICS:
                detail_row[metric] = row[metric]
            rows.append(detail_row)
    return rows


def write_campaign_analysis_section(
    worksheet,
    section_df: pd.DataFrame,
    *,
    start_row: int,
    start_col: int,
    title: str,
    formats: dict[str, object],
    section_style: str,
) -> None:
    total_row = start_row
    title_row = start_row + 2
    header_row = start_row + 3
    data_start_row = start_row + 4
    headers = ["Campaign ID", "Campaign name", *CAMPAIGN_ANALYSIS_METRICS]
    last_section_col = start_col + len(headers) - 1
    display_rows = campaign_analysis_display_rows(section_df)

    worksheet.write_blank(total_row, start_col, None, formats[f"{section_style}_publisher_row"])
    worksheet.write(total_row, start_col + 1, "GRAND TOTALS", formats[f"{section_style}_publisher_row"])
    worksheet.write(title_row, start_col, title, formats["section_title_center"])
    for col in range(start_col + 1, last_section_col + 1):
        worksheet.write_blank(title_row, col, None, formats["section_title_center"])

    for offset, header in enumerate(headers):
        worksheet.write(header_row, start_col + offset, header, formats["header"])

    publisher_rows: list[int] = []
    for row_offset, row in enumerate(display_rows):
        excel_row = data_start_row + row_offset
        if row["RowType"] == "spacer":
            for offset in range(len(headers)):
                worksheet.write_blank(excel_row, start_col + offset, None)
            continue

        row_type = "publisher" if row["RowType"] == "publisher" else "pgrp"
        row_format = (
            formats[f"{section_style}_publisher_row"]
            if row_type == "publisher"
            else formats[f"{section_style}_detail_row"]
        )
        if row_type == "publisher":
            publisher_rows.append(excel_row)
        for offset, dimension in enumerate(["Campaign ID", "Campaign name"]):
            worksheet.write(excel_row, start_col + offset, row[dimension], row_format)
        for metric_offset, metric in enumerate(CAMPAIGN_ANALYSIS_METRICS, start=2):
            cell_format = metric_format(metric, formats, row_type=row_type, section_style=section_style)
            worksheet.write(excel_row, start_col + metric_offset, row[metric], cell_format)

    for metric_offset, metric in enumerate(CAMPAIGN_ANALYSIS_METRICS, start=2):
        col = start_col + metric_offset
        col_letter = xl_col(col)
        cell_format = metric_format(metric, formats, row_type="publisher", section_style=section_style)
        if metric in {"CTR", "CVR", "ACOS", "ROAS"}:
            formula = campaign_analysis_ratio_formula(metric, total_row + 1, start_col)
            worksheet.write_formula(total_row, col, formula, cell_format)
        else:
            publisher_refs = ",".join(f"{col_letter}{row_idx + 1}" for row_idx in publisher_rows)
            total_formula = f"=SUM({publisher_refs})" if publisher_refs else "=0"
            worksheet.write_formula(total_row, col, total_formula, cell_format)


def campaign_analysis_ratio_formula(metric: str, excel_row: int, start_col: int) -> str:
    positions = {
        metric_name: start_col + 2 + idx
        for idx, metric_name in enumerate(CAMPAIGN_ANALYSIS_METRICS)
    }
    if metric == "CTR":
        numerator, denominator = "Clicks", "Impressions"
    elif metric == "CVR":
        numerator, denominator = "Units sold", "Clicks"
    elif metric == "ACOS":
        numerator, denominator = "Spend", "Sales"
    elif metric == "ROAS":
        numerator, denominator = "Sales", "Spend"
    else:
        return ""
    return f"=IFERROR({xl_col(positions[numerator])}{excel_row}/{xl_col(positions[denominator])}{excel_row},0)"


def write_campaign_analysis_sheet(
    writer: pd.ExcelWriter,
    summary: pd.DataFrame,
    current_period: str,
    formats: dict[str, object],
    prior_summary: pd.DataFrame | None = None,
    prior_period: str | None = None,
) -> None:
    current = summarize_for_campaign_analysis(summary)
    worksheet = writer.book.add_worksheet("Campaign Analysis")
    writer.sheets["Campaign Analysis"] = worksheet
    write_campaign_analysis_section(
        worksheet,
        current,
        start_row=0,
        start_col=0,
        title=pd.Period(current_period, freq="M").strftime("%B %Y"),
        formats=formats,
        section_style="current",
    )
    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, 1, 42)
    worksheet.set_column(2, len(CAMPAIGN_ANALYSIS_METRICS) + 1, 10)
    worksheet.freeze_panes(4, 2)


def filter_summary_for_publisher_report(summary: pd.DataFrame, report_suffix: str) -> pd.DataFrame:
    column, value = PUBLISHER_REPORTS[report_suffix]
    normalized = summary.copy()
    normalized_groups = normalized.apply(
        lambda row: normalize_publisher_group(row["Publisher"], row["pgrp"]),
        axis=1,
        result_type="expand",
    )
    normalized["Publisher"] = normalized_groups[0]
    normalized["pgrp"] = normalized_groups[1]
    return normalized[normalized[column].astype("string").str.strip().eq(value)].copy()


def write_summary_excel(
    summary: pd.DataFrame,
    output_file: Path,
    current_period: str,
    prior_summary: pd.DataFrame | None = None,
    prior_period: str | None = None,
    include_campaign_analysis: bool = True,
    campaign_detail: pd.DataFrame | None = None,
    prior_campaign_detail: pd.DataFrame | None = None,
) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="xlsxwriter", datetime_format="m/d/yyyy") as writer:
        workbook = writer.book
        formats = build_workbook_formats(workbook)
        write_monthly_comparison_sheet(
            writer,
            summary,
            current_period,
            formats,
            prior_summary=prior_summary,
            prior_period=prior_period,
        )
        if include_campaign_analysis and campaign_detail is not None:
            write_campaign_analysis_sheet(
                writer,
                campaign_detail,
                current_period,
                formats,
                prior_summary=prior_campaign_detail,
                prior_period=prior_period,
            )

        summary.to_excel(writer, sheet_name="USE_main", index=False, startrow=3)
        worksheet = writer.sheets["USE_main"]

        date_format = workbook.add_format({"num_format": "m/d/yyyy"})

        for col_idx, column in enumerate(summary.columns):
            worksheet.write(3, col_idx, column, formats["header"])

        first_data_row = 5
        last_row = len(summary) + 4
        column_positions = {column: idx for idx, column in enumerate(summary.columns)}
        worksheet.write(0, 5, "Total", formats["section_title"])
        worksheet.write(1, 5, "Subtotal", formats["section_title"])
        for column in ["Impressions", "Clicks", "Units sold", "Count of Campaigns"]:
            col = column_positions[column]
            total_range = f"{xl_col(col)}{first_data_row}:{xl_col(col)}{last_row}"
            worksheet.write_formula(0, col, f"=SUM({total_range})", formats["integer"])
            worksheet.write_formula(1, col, f"=SUBTOTAL(109,{total_range})", formats["integer"])
        for column in ["Spend", "Sales"]:
            col = column_positions[column]
            total_range = f"{xl_col(col)}{first_data_row}:{xl_col(col)}{last_row}"
            worksheet.write_formula(0, col, f"=SUM({total_range})", formats["money"])
            worksheet.write_formula(1, col, f"=SUBTOTAL(109,{total_range})", formats["money"])

        formulas = {
            "CTR": ("Clicks", "Impressions", formats["percent"]),
            "CVR": ("Units sold", "Clicks", formats["percent"]),
            "ACOS": ("Spend", "Sales", formats["percent"]),
            "ROAS": ("Sales", "Spend", formats["number"]),
        }
        for target, (numerator, denominator, cell_format) in formulas.items():
            target_col = column_positions[target]
            numerator_col = column_positions[numerator]
            denominator_col = column_positions[denominator]
            for row_idx, excel_row in [(0, 1), (1, 2)]:
                worksheet.write_formula(
                    row_idx,
                    target_col,
                    f"=IFERROR({xl_col(numerator_col)}{excel_row}/{xl_col(denominator_col)}{excel_row},0)",
                    cell_format,
                )

        worksheet.freeze_panes(4, 0)
        worksheet.autofilter(3, 0, max(3, last_row - 1), len(summary.columns) - 1)
        worksheet.set_column("A:A", 13)
        worksheet.set_column("B:B", 15)
        worksheet.set_column("C:C", 34)
        worksheet.set_column("D:E", 14)
        worksheet.set_column("F:F", 12, date_format)
        worksheet.set_column("G:H", 12, formats["integer"])
        worksheet.set_column("I:I", 10, formats["percent"])
        worksheet.set_column("J:J", 12, formats["integer"])
        worksheet.set_column("K:K", 10, formats["percent"])
        worksheet.set_column("L:M", 12, formats["money"])
        worksheet.set_column("N:N", 10, formats["percent"])
        worksheet.set_column("O:O", 10, formats["number"])
        worksheet.set_column("P:P", 18, formats["integer"])
        worksheet.set_column("Q:Q", 28)
        worksheet.set_column("R:R", 12)


def build_workbook_formats(workbook) -> dict[str, object]:
    accounting_no_decimals = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
    formats = {
        "header": workbook.add_format({"bold": True, "bg_color": "#BFBFBF"}),
        "section_title": workbook.add_format({"bold": True}),
        "section_title_center": workbook.add_format({"bold": True, "align": "center_across"}),
        "integer": workbook.add_format({"num_format": "#,##0"}),
        "money": workbook.add_format({"num_format": accounting_no_decimals}),
        "percent": workbook.add_format({"num_format": "0.00%"}),
        "number": workbook.add_format({"num_format": "0.00"}),
    }
    for style, publisher_color, detail_color in [
        ("current", "#D9EAD3", "#EAF2F8"),
        ("prior", "#FCE4D6", "#FFF2CC"),
        ("variance", "#D9E1F2", "#E4DFEC"),
    ]:
        formats[f"{style}_publisher_row"] = workbook.add_format({"bold": True, "bg_color": publisher_color})
        formats[f"{style}_publisher_integer"] = workbook.add_format({"bold": True, "bg_color": publisher_color, "num_format": "#,##0"})
        formats[f"{style}_publisher_money"] = workbook.add_format({"bold": True, "bg_color": publisher_color, "num_format": accounting_no_decimals})
        formats[f"{style}_publisher_percent"] = workbook.add_format({"bold": True, "bg_color": publisher_color, "num_format": "0.00%"})
        formats[f"{style}_publisher_number"] = workbook.add_format({"bold": True, "bg_color": publisher_color, "num_format": "0.00"})
        formats[f"{style}_detail_row"] = workbook.add_format({"bg_color": detail_color})
        formats[f"{style}_detail_integer"] = workbook.add_format({"bg_color": detail_color, "num_format": "#,##0"})
        formats[f"{style}_detail_money"] = workbook.add_format({"bg_color": detail_color, "num_format": accounting_no_decimals})
        formats[f"{style}_detail_percent"] = workbook.add_format({"bg_color": detail_color, "num_format": "0.00%"})
        formats[f"{style}_detail_number"] = workbook.add_format({"bg_color": detail_color, "num_format": "0.00"})
    return formats


def read_history() -> pd.DataFrame:
    if not AMS_HISTORY_PARQUET.exists():
        return pd.DataFrame(columns=HISTORY_COLUMNS)
    history = pd.read_parquet(AMS_HISTORY_PARQUET)
    for column in HISTORY_COLUMNS:
        if column not in history.columns:
            history[column] = pd.NA
    return history[HISTORY_COLUMNS]


def history_month_summary(period: str) -> pd.DataFrame:
    history = read_history()
    if history.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    return history[history["period"].astype(str).eq(period)][OUTPUT_COLUMNS].copy()


def write_history(history: pd.DataFrame) -> None:
    AMS_HISTORY_PARQUET.parent.mkdir(parents=True, exist_ok=True)
    history[HISTORY_COLUMNS].to_parquet(AMS_HISTORY_PARQUET, index=False)


def history_to_legacy_aggregate(history: pd.DataFrame) -> pd.DataFrame:
    output = pd.DataFrame(
        {
            "ASIN": history["ASIN"].astype("string").fillna("").str.strip(),
            "impressions": history["Impressions"],
            "clicks": history["Clicks"],
            "spend": history["Spend"],
            "14 day total sales": history["Sales"],
            "units sold": history["Units sold"],
            "count of campaigns": history["Count of Campaigns"],
            "ISBN": pd.to_numeric(
                history["ISBN"].replace({"NO ISBN": pd.NA, "": pd.NA}),
                errors="coerce",
            ),
            "CTR": history["CTR"],
            "CRV": history["CVR"],
            "ACOS": history["ACOS"],
            "ROAS": history["ROAS"],
            "period": history["period"].astype("string").fillna("").str.strip(),
            "source_file": history["source_file"],
            "title": history["Title"],
            "publisher": history["Publisher"],
            "pgrp": history["pgrp"],
            "PT": history["Format"],
            "OSD": pd.to_datetime(history["osd"], errors="coerce"),
        }
    )
    return output[LEGACY_AGGREGATE_COLUMNS]


def refresh_legacy_aggregate_outputs(history: pd.DataFrame) -> pd.DataFrame:
    history_periods = set(history["period"].dropna().astype(str).unique())
    aggregate = history_to_legacy_aggregate(history)

    if AMS_AGGREGATE_EXCEL.exists() and history_periods:
        existing = pd.read_excel(AMS_AGGREGATE_EXCEL)
        for column in LEGACY_AGGREGATE_COLUMNS:
            if column not in existing.columns:
                existing[column] = pd.NA
        existing = existing[LEGACY_AGGREGATE_COLUMNS]
        existing_periods = existing["period"].dropna().astype(str)
        existing = existing[~existing_periods.isin(history_periods)].copy()
        aggregate = pd.concat([existing, aggregate], ignore_index=True)

    if not aggregate.empty and "period" in aggregate.columns:
        aggregate = aggregate.sort_values(["period", "ASIN"], kind="stable").reset_index(drop=True)

    if AMS_AGGREGATE_EXCEL.exists():
        archive_dir = AMS_AGGREGATE_EXCEL.parent / "archive"
        archive_dir.mkdir(parents=True, exist_ok=True)
        archive_file = archive_dir / f"{AMS_AGGREGATE_EXCEL.stem}_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx"
        shutil.copy2(AMS_AGGREGATE_EXCEL, archive_file)

    AMS_AGGREGATE_EXCEL.parent.mkdir(parents=True, exist_ok=True)
    aggregate.to_excel(AMS_AGGREGATE_EXCEL, index=False)
    aggregate.to_pickle(AMS_AGGREGATE_PICKLE)
    return aggregate


def upsert_history_month(summary: pd.DataFrame, period: str, source_file: Path) -> pd.DataFrame:
    history = read_history()
    history = history[~history["period"].astype(str).eq(period)].copy()
    month_rows = summary.copy()
    month_rows.insert(0, "source_file", str(source_file))
    month_rows.insert(0, "period", period)
    updated = pd.concat([history, month_rows[HISTORY_COLUMNS]], ignore_index=True)
    updated = updated.sort_values(["period", "campaign", "ASIN"], kind="stable").reset_index(drop=True)
    write_history(updated)
    return updated


def xl_col(col_idx: int) -> str:
    letters = ""
    col_idx += 1
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def default_output_file(source_file: Path) -> Path:
    return source_file.with_name(f"{source_file.stem}{DEFAULT_OUTPUT_SUFFIX}")


def run(
    source_file: Path,
    output_file: Path | None = None,
    prompt: bool = True,
    missing_asin_overrides: dict[str, list[str]] | None = None,
    prior_input: Path | None = None,
    save_history: bool = True,
    write_publisher_reports: bool = True,
) -> Path:
    current_period = parse_period_from_filename(source_file)
    mapping = load_asin_isbn_map()
    metadata = load_item_metadata()

    df = read_campaign_csv(source_file)
    df = resolve_missing_asins(df, prompt=prompt, overrides=missing_asin_overrides)
    print_advertiser_summary(df)
    if prompt:
        input("Press Enter to continue and create the ISBN summary workbook...")

    summary = build_summary(df, mapping=mapping, metadata=metadata)
    campaign_detail = build_campaign_detail(df, mapping=mapping, metadata=metadata)
    prior_summary = None
    prior_campaign_detail = None
    prior_period = None
    if prior_input is not None:
        prior_period = parse_period_from_filename(prior_input)
        prior_df = read_campaign_csv(prior_input)
        prior_df = resolve_missing_asins(
            prior_df,
            prompt=prompt,
            overrides=missing_asin_overrides,
        )
        prior_summary = build_summary(prior_df, mapping=mapping, metadata=metadata)
        prior_campaign_detail = build_campaign_detail(prior_df, mapping=mapping, metadata=metadata)
        print(f"Using prior month CSV for comparison: {prior_input}")
    else:
        prior_period = previous_period(current_period)
        prior_source_file = CAMPAIGN_FOLDER / f"{period_compact(prior_period)}_MonthlyCampaigns.csv"
        if prior_source_file.exists():
            prior_df = read_campaign_csv(prior_source_file)
            prior_df = resolve_missing_asins(
                prior_df,
                prompt=False,
                overrides=missing_asin_overrides,
            )
            prior_campaign_detail = build_campaign_detail(prior_df, mapping=mapping, metadata=metadata)

    if save_history:
        history = upsert_history_month(summary, current_period, source_file)
        aggregate = refresh_legacy_aggregate_outputs(history)
        print(f"Saved AMS history parquet: {AMS_HISTORY_PARQUET}")
        print(f"  History rows: {len(history):,}")
        print(f"Refreshed AMS aggregate workbook: {AMS_AGGREGATE_EXCEL}")
        print(f"  Aggregate rows: {len(aggregate):,}")

    output = output_file or final_report_file(current_period, "ALL")
    write_summary_excel(
        summary,
        output,
        current_period,
        prior_summary=prior_summary,
        prior_period=prior_period,
        include_campaign_analysis=True,
        campaign_detail=campaign_detail,
        prior_campaign_detail=prior_campaign_detail,
    )
    if write_publisher_reports:
        base_prior_summary = (
            prior_summary
            if prior_summary is not None
            else history_month_summary(prior_period or previous_period(current_period))
        )
        for report_suffix in PUBLISHER_REPORTS:
            report_summary = filter_summary_for_publisher_report(summary, report_suffix)
            report_prior_summary = filter_summary_for_publisher_report(
                base_prior_summary, report_suffix
            )
            report_output = final_report_file(current_period, report_suffix)
            write_summary_excel(
                report_summary,
                report_output,
                current_period,
                prior_summary=report_prior_summary,
                prior_period=prior_period,
                include_campaign_analysis=False,
            )
            print(f"Saved AMS {report_suffix} report: {report_output}")

    no_isbn_count = int(summary["ISBN"].eq("NO ISBN").sum())
    print(f"Saved AMS monthly campaign summary: {output}")
    print(f"  Rows:    {len(summary):,}")
    print(f"  NO ISBN: {no_isbn_count:,}")
    return output


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build Amazon AMS monthly campaign summary.")
    parser.add_argument("--input", type=Path, help="Monthly campaign CSV file.")
    parser.add_argument("--output", type=Path, help="Output workbook path.")
    parser.add_argument("--prior-input", type=Path, help="Optional prior month campaign CSV for Monthly Comparisons.")
    parser.add_argument("--no-history", action="store_true", help="Do not save this month to the AMS history parquet.")
    parser.add_argument("--no-pwp", action="store_true", help="Do not write publisher-specific workbooks.")
    parser.add_argument("--no-prompt", action="store_true", help="Do not pause after advertiser summary.")
    parser.add_argument(
        "--missing-asin",
        action="append",
        help="Resolve a blank ASIN row with Campaign name=ASIN1,ASIN2. Can be repeated.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    source_file = args.input or choose_campaign_file()
    run(
        source_file,
        output_file=args.output,
        prompt=not args.no_prompt,
        missing_asin_overrides=parse_missing_asin_overrides(args.missing_asin),
        prior_input=args.prior_input,
        save_history=not args.no_history,
        write_publisher_reports=not args.no_pwp,
    )


if __name__ == "__main__":
    main()
