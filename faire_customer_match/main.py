from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from shared.db.connection import get_connection
from shared.db.query_runner import fetch_data_from_db


DEFAULT_INPUT = REPO_ROOT / "Faire_to_HBG_Number_Search.xlsx"
DEFAULT_CACHE = REPO_ROOT / "faire_customer_match" / "active_ebs_customers.csv"
DEFAULT_OUTPUT = REPO_ROOT / "Faire_to_HBG_Number_Search_matched.xlsx"
MATCH_COLUMNS = ["Account Name", "Address 1", "Address 2", "City", "State", "Zip Code", "Country"]
CUSTOMER_KEY_COLUMNS = [
    "_name_norm",
    "_address1_norm",
    "_address_norm",
    "_city_norm",
    "_state_norm",
    "_zip_norm",
    "_country_norm",
    "_addr_num",
    "_street_key",
]


CUSTOMER_QUERY = """
SELECT DISTINCT
    c.PARTYSITENUMBER,
    c.PARTYSITENAME,
    c.ADDRESS1,
    c.ADDRESS2,
    c.ADDRESS3,
    c.ADDRESS4,
    c.CITY,
    c.STATE,
    c.POSTAL_CODE AS ZipCode,
    c.COUNTRY_CODE
FROM ebs.customer c
WHERE
    c.ACCOUNT_STATUS = 'A'
    AND c.PARTYSITENUMBER IS NOT NULL;
""".strip()


def clean_scalar(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def normalize_text(value: object) -> str:
    text = clean_scalar(value).upper()
    text = text.replace("&", " AND ")
    text = re.sub(r"^[`'\"@]+", "", text)
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_zip(value: object) -> str:
    text = clean_scalar(value).upper()
    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    if len(digits) < 5:
        return digits.zfill(5)
    return digits[:5]


def normalize_country(value: object) -> str:
    text = normalize_text(value)
    if text in {"UNITED STATES", "UNITED STATES OF AMERICA", "USA", "U S A"}:
        return "US"
    return text


def address_number(value: object) -> str:
    match = re.match(r"\s*(\d+)", clean_scalar(value))
    return match.group(1) if match else ""


def street_key(value: object) -> str:
    text = normalize_text(value)
    pieces = [p for p in text.split() if p not in {"ST", "STREET", "RD", "ROAD", "AVE", "AVENUE"}]
    return " ".join(pieces[:3])


def similarity(left: str, right: str) -> float:
    if not left and not right:
        return 1.0
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    return SequenceMatcher(None, left, right).ratio()


def token_similarity(left: str, right: str) -> float:
    if not left and not right:
        return 1.0
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    if left in right or right in left:
        length_ratio = min(len(left), len(right)) / max(len(left), len(right))
        return max(0.82, length_ratio)

    left_tokens = set(left.split())
    right_tokens = set(right.split())
    if not left_tokens or not right_tokens:
        return 0.0

    overlap = len(left_tokens & right_tokens)
    if not overlap:
        return 0.0

    dice = 2 * overlap / (len(left_tokens) + len(right_tokens))
    length_ratio = min(len(left), len(right)) / max(len(left), len(right))
    return 0.85 * dice + 0.15 * length_ratio


def add_faire_keys(df: pd.DataFrame) -> pd.DataFrame:
    output = df.copy()
    output["_row_id"] = range(1, len(output) + 1)
    output["_name_norm"] = output["Account Name"].map(normalize_text)
    output["_address1_norm"] = output["Address 1"].map(normalize_text)
    output["_address2_norm"] = output["Address 2"].map(normalize_text)
    output["_address_norm"] = (
        output["Address 1"].map(clean_scalar) + " " + output["Address 2"].map(clean_scalar)
    ).map(normalize_text)
    output["_city_norm"] = output["City"].map(normalize_text)
    output["_state_norm"] = output["State"].map(normalize_text)
    output["_zip_norm"] = output["Zip Code"].map(normalize_zip)
    output["_country_norm"] = output["Country"].map(normalize_country)
    output["_addr_num"] = output["Address 1"].map(address_number)
    output["_street_key"] = output["Address 1"].map(street_key)
    return output


def add_customer_keys(df: pd.DataFrame) -> pd.DataFrame:
    output = df.copy()
    for column in [
        "PARTYSITENUMBER",
        "PARTYSITENAME",
        "ADDRESS1",
        "ADDRESS2",
        "ADDRESS3",
        "ADDRESS4",
        "CITY",
        "STATE",
        "ZipCode",
        "COUNTRY_CODE",
    ]:
        if column not in output.columns:
            output[column] = ""
    output["_name_norm"] = output["PARTYSITENAME"].map(normalize_text)
    output["_address1_norm"] = output["ADDRESS1"].map(normalize_text)
    output["_address_norm"] = (
        output["ADDRESS1"].map(clean_scalar)
        + " "
        + output["ADDRESS2"].map(clean_scalar)
        + " "
        + output["ADDRESS3"].map(clean_scalar)
        + " "
        + output["ADDRESS4"].map(clean_scalar)
    ).map(normalize_text)
    output["_city_norm"] = output["CITY"].map(normalize_text)
    output["_state_norm"] = output["STATE"].map(normalize_text)
    output["_zip_norm"] = output["ZipCode"].map(normalize_zip)
    output["_country_norm"] = output["COUNTRY_CODE"].map(normalize_country)
    output["_addr_num"] = output["ADDRESS1"].map(address_number)
    output["_street_key"] = output["ADDRESS1"].map(street_key)
    return output


def ensure_customer_keys(df: pd.DataFrame) -> pd.DataFrame:
    if all(column in df.columns for column in CUSTOMER_KEY_COLUMNS):
        return df.fillna("")
    return add_customer_keys(df)


def blocking_keys(row: pd.Series) -> list[tuple[str, str]]:
    keys = []
    if row["_zip_norm"] and row["_state_norm"] and row["_addr_num"] and row["_street_key"]:
        keys.append(("zip_state_street", "|".join([row["_zip_norm"], row["_state_norm"], row["_addr_num"], row["_street_key"]])))
    if row["_zip_norm"] and row["_state_norm"] and row["_city_norm"]:
        keys.append(("zip_state_city", "|".join([row["_zip_norm"], row["_state_norm"], row["_city_norm"]])))
    if row["_state_norm"] and row["_addr_num"] and row["_street_key"]:
        keys.append(("state_street", "|".join([row["_state_norm"], row["_addr_num"], row["_street_key"]])))
    return keys


def build_index(customers: pd.DataFrame) -> dict[tuple[str, str], list[int]]:
    index: dict[tuple[str, str], list[int]] = defaultdict(list)
    for idx, row in customers.iterrows():
        for key in blocking_keys(row):
            index[key].append(idx)
    return index


def candidate_indexes(row: pd.Series, index: dict[tuple[str, str], list[int]]) -> list[int]:
    groups = {key[0]: index.get(key, []) for key in blocking_keys(row)}
    street_zip = groups.get("zip_state_street", [])
    city_zip = groups.get("zip_state_city", [])
    street_state = groups.get("state_street", [])

    if street_zip:
        return street_zip

    if city_zip and street_state:
        city_set = set(city_zip)
        intersection = [idx for idx in street_state if idx in city_set]
        if intersection:
            return intersection

    if city_zip and len(city_zip) <= 175:
        return city_zip

    if street_state and len(street_state) <= 175:
        return street_state

    return city_zip[:175] or street_state[:175]


def score_candidate(faire: pd.Series, customer: pd.Series) -> tuple[float, str]:
    name_score = similarity(faire["_name_norm"], customer["_name_norm"])
    address_score = max(
        token_similarity(faire["_address_norm"], customer["_address_norm"]),
        token_similarity(faire["_address1_norm"], customer["_address1_norm"]),
    )
    city_score = token_similarity(faire["_city_norm"], customer["_city_norm"])
    state_match = faire["_state_norm"] == customer["_state_norm"] and bool(faire["_state_norm"])
    zip_match = faire["_zip_norm"] == customer["_zip_norm"] and bool(faire["_zip_norm"])
    country_match = not faire["_country_norm"] or not customer["_country_norm"] or faire["_country_norm"] == customer["_country_norm"]

    weighted = (
        45 * address_score
        + 30 * name_score
        + 10 * city_score
        + (10 if zip_match else 0)
        + (4 if state_match else 0)
        + (1 if country_match else 0)
    )

    if zip_match and state_match and address_score >= 0.95:
        reason = "excellent_address"
    elif zip_match and state_match and address_score >= 0.85 and name_score >= 0.50:
        reason = "good_address"
    elif zip_match and state_match and name_score >= 0.86:
        reason = "good_name_same_zip"
    elif state_match and address_score >= 0.90 and name_score >= 0.50:
        reason = "good_address_no_zip"
    else:
        reason = "review"

    return round(weighted, 2), reason


def confidence(score: float, reason: str, best: pd.Series | None, runner_up_score: float | None) -> str:
    if best is None:
        return "No candidate"
    margin = score - (runner_up_score or 0)
    if reason.startswith("excellent") and score >= 88 and margin >= 3:
        return "High"
    if reason.startswith("good") and score >= 78 and margin >= 5:
        return "Medium"
    return "Review"


def pick_best_match(faire: pd.Series, candidates: pd.DataFrame) -> dict[str, object]:
    if candidates.empty:
        return {
            "Matched PARTYSITENUMBER": "",
            "Match Confidence": "No candidate",
            "Match Score": 0,
            "Match Reason": "no_block_candidate",
        }

    scored = []
    for _, customer in candidates.iterrows():
        score, reason = score_candidate(faire, customer)
        scored.append((score, reason, customer))
    scored.sort(key=lambda item: item[0], reverse=True)

    best_score, reason, best = scored[0]
    runner_up = scored[1][0] if len(scored) > 1 else None
    return {
        "Matched PARTYSITENUMBER": best["PARTYSITENUMBER"],
        "Matched PARTYSITENAME": best["PARTYSITENAME"],
        "Matched Address 1": best["ADDRESS1"],
        "Matched Address 2": best["ADDRESS2"],
        "Matched City": best["CITY"],
        "Matched State": best["STATE"],
        "Matched ZipCode": best["ZipCode"],
        "Matched Country": best["COUNTRY_CODE"],
        "Match Confidence": confidence(best_score, reason, best, runner_up),
        "Match Score": best_score,
        "Match Reason": reason,
        "Runner Up Score": runner_up or "",
    }


def refresh_cache(cache_path: Path) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    engine = get_connection()
    customers = add_customer_keys(fetch_data_from_db(engine, CUSTOMER_QUERY))
    customers.to_csv(cache_path, index=False)
    print(f"Wrote {len(customers):,} active customers to {cache_path}")


def read_faire(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)
    missing = [column for column in MATCH_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"{path} is missing required column(s): {', '.join(missing)}")
    return df


def read_customers(cache_path: Path) -> pd.DataFrame:
    if not cache_path.exists():
        raise FileNotFoundError(
            f"Customer cache not found: {cache_path}. Run `refresh-cache` first."
        )
    return pd.read_csv(cache_path, dtype=str).fillna("")


def match_customers(input_path: Path, cache_path: Path, output_path: Path) -> None:
    faire = add_faire_keys(read_faire(input_path))
    customers = ensure_customer_keys(read_customers(cache_path))
    index = build_index(customers)

    matches = []
    for _, faire_row in faire.iterrows():
        candidate_ids = candidate_indexes(faire_row, index)
        candidates = customers.loc[candidate_ids] if candidate_ids else customers.iloc[0:0]
        matches.append(pick_best_match(faire_row, candidates))

    output = pd.concat([faire.drop(columns=[c for c in faire.columns if c.startswith("_")]), pd.DataFrame(matches)], axis=1)
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        output.to_excel(writer, sheet_name="Matched", index=False)
        summary = output["Match Confidence"].value_counts(dropna=False).rename_axis("Match Confidence").reset_index(name="Rows")
        summary.to_excel(writer, sheet_name="Summary", index=False)
    print(f"Wrote {len(output):,} matched Faire rows to {output_path}")
    print(output["Match Confidence"].value_counts(dropna=False).to_string())


def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Match Faire customer rows to active EBS customer accounts.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    refresh = subparsers.add_parser("refresh-cache", help="Download active EBS customers to a local CSV cache.")
    refresh.add_argument("--cache", type=Path, default=DEFAULT_CACHE)

    match = subparsers.add_parser("match", help="Match a Faire workbook against the local EBS customer cache.")
    match.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    match.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    match.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)

    both = subparsers.add_parser("run", help="Refresh the EBS cache, then match the Faire workbook.")
    both.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    both.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    both.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser.parse_args(argv)


def main(argv: Iterable[str] | None = None) -> None:
    args = parse_args(argv)
    if args.command == "refresh-cache":
        refresh_cache(args.cache)
    elif args.command == "match":
        match_customers(args.input, args.cache, args.output)
    elif args.command == "run":
        refresh_cache(args.cache)
        match_customers(args.input, args.cache, args.output)


if __name__ == "__main__":
    main()
