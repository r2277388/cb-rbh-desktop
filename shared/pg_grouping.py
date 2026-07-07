from __future__ import annotations

import pandas as pd


CHRONICLE_RPG_GROUPS = ("PTC", "RID", "GAM")
CHRONICLE_BOOK_GIFT_GROUPS = (
    "FWN",
    "ART",
    "ENT",
    "LIF",
    "CCB",
    "CPB",
    "CPA",
    "BAR-LIF",
    "BAR-ENT",
)


def pg_grouping_value(publisher: object, publishing_group: object, product_type: object) -> str:
    pub = "" if pd.isna(publisher) else str(publisher).strip()
    pgrp = "" if pd.isna(publishing_group) else str(publishing_group).strip()
    pt = "" if pd.isna(product_type) else str(product_type).strip()

    if pub in {"Quadrille", "Quadrille Publishing Limited"}:
        return "Quadrille"
    if pub and pub != "Chronicle":
        return pub
    if pub == "Chronicle" and pgrp == "CHL" and pt in {"BK", "FT"}:
        return "CBKids"
    if pub == "Chronicle" and pgrp in CHRONICLE_RPG_GROUPS and pt in {"BK", "FT"}:
        return "CBRPG"
    if pub == "Chronicle" and pgrp in CHRONICLE_BOOK_GIFT_GROUPS and pt == "BK":
        return "CBBook"
    if pub == "Chronicle" and pgrp in CHRONICLE_BOOK_GIFT_GROUPS and pt == "FT":
        return "CBGift"
    return "-"


def apply_pg_grouping(
    df: pd.DataFrame,
    publisher_col: str = "Pub",
    publishing_group_col: str = "pgrp",
    product_type_col: str = "pt",
    output_col: str = "PG_Grouping",
) -> pd.DataFrame:
    required = {publisher_col, publishing_group_col, product_type_col}
    if not required.issubset(df.columns):
        return df

    output = df.copy()
    output[output_col] = [
        pg_grouping_value(pub, pgrp, pt)
        for pub, pgrp, pt in zip(
            output[publisher_col],
            output[publishing_group_col],
            output[product_type_col],
        )
    ]
    return output


def build_pg_grouping_sql_case(
    publisher_col: str,
    publishing_group_col: str,
    product_type_col: str,
    alias: str = "PG_Grouping",
    quote: str = "'",
    indent: str = "",
) -> str:
    def q(value: str) -> str:
        return f"{quote}{value}{quote}"

    rpg_groups = ", ".join(q(value) for value in CHRONICLE_RPG_GROUPS)
    book_gift_groups = ", ".join(q(value) for value in CHRONICLE_BOOK_GIFT_GROUPS)
    lines = [
        "CASE",
        f"    WHEN {publisher_col} = {q('Quadrille Publishing Limited')} THEN {q('Quadrille')}",
        f"    WHEN {publisher_col} <> {q('Chronicle')} THEN {publisher_col}",
        f"    WHEN {publisher_col} = {q('Chronicle')} AND {publishing_group_col} = {q('CHL')} AND {product_type_col} IN ({q('BK')}, {q('FT')}) THEN {q('CBKids')}",
        f"    WHEN {publisher_col} = {q('Chronicle')} AND {publishing_group_col} IN ({rpg_groups}) AND {product_type_col} IN ({q('BK')}, {q('FT')}) THEN {q('CBRPG')}",
        f"    WHEN {publisher_col} = {q('Chronicle')} AND {publishing_group_col} IN ({book_gift_groups}) AND {product_type_col} = {q('BK')} THEN {q('CBBook')}",
        f"    WHEN {publisher_col} = {q('Chronicle')} AND {publishing_group_col} IN ({book_gift_groups}) AND {product_type_col} = {q('FT')} THEN {q('CBGift')}",
        f"    ELSE {q('-')}",
        f"END AS {alias}",
    ]
    return "\n".join(f"{indent}{line}" for line in lines)
