from __future__ import annotations

import numpy as np
import pandas as pd


LAG_PERIODS = (1, 2, 3, 6, 12)
ROLL_WINDOWS = (3, 6, 12)


def add_time_features(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["month"] = out["period"].dt.month
    out["quarter"] = out["period"].dt.quarter
    out["year"] = out["period"].dt.year
    return out


def add_lag_features(df: pd.DataFrame, target_col: str = "qty") -> pd.DataFrame:
    out = df.copy()
    grp = out.groupby(["isbn", "ssr_id"], group_keys=False)

    for lag in LAG_PERIODS:
        out[f"lag_{lag}"] = grp[target_col].shift(lag)

    for w in ROLL_WINDOWS:
        out[f"roll_mean_{w}"] = (
            grp[target_col].shift(1).rolling(w, min_periods=1).mean().reset_index(level=[0, 1], drop=True)
        )

    out["order_lag_1"] = grp["order_qty"].shift(1)
    return out


def make_supervised(df: pd.DataFrame) -> pd.DataFrame:
    out = df.sort_values(["isbn", "ssr_id", "period"]).reset_index(drop=True)
    out = add_time_features(out)
    out = add_lag_features(out)

    feature_cols = [
        "pgrp",
        "pt",
        "month",
        "quarter",
        "year",
        "order_qty",
        "order_lag_1",
        "lag_1",
        "lag_2",
        "lag_3",
        "lag_6",
        "lag_12",
        "roll_mean_3",
        "roll_mean_6",
        "roll_mean_12",
    ]

    out = out.dropna(subset=["lag_12"]).copy()
    out[feature_cols] = out[feature_cols].replace([np.inf, -np.inf], np.nan).fillna(0)
    return out


def build_future_rows(history: pd.DataFrame, horizon: int, freq: str = "M") -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for (isbn, ssr_id), g in history.groupby(["isbn", "ssr_id"], sort=False):
        g = g.sort_values("period")
        last = g.iloc[-1]
        for h in range(1, horizon + 1):
            rows.append(
                {
                    "period": last["period"] + pd.tseries.frequencies.to_offset(freq) * h,
                    "isbn": isbn,
                    "ssr_id": ssr_id,
                    "pgrp": last["pgrp"],
                    "pt": last["pt"],
                    "qty": np.nan,
                    "val": np.nan,
                    "order_qty": 0.0,
                }
            )
    return pd.DataFrame(rows)
