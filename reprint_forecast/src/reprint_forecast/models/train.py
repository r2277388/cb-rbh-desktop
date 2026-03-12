from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np
import pandas as pd
from sklearn.compose import ColumnTransformer
from sklearn.ensemble import RandomForestRegressor
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder

from reprint_forecast.models.evaluate import evaluate_predictions

try:
    from xgboost import XGBRegressor
    HAS_XGBOOST = True
except Exception:
    HAS_XGBOOST = False


@dataclass
class ModelResult:
    name: str
    metrics: dict[str, float]
    model: Any


class SeasonalNaiveModel:
    def __init__(self, seasonal_lag: int = 12) -> None:
        self.seasonal_lag = seasonal_lag

    def fit(self, X: pd.DataFrame, y: pd.Series) -> "SeasonalNaiveModel":
        return self

    def predict(self, X: pd.DataFrame) -> np.ndarray:
        col = f"lag_{self.seasonal_lag}"
        if col in X.columns:
            return X[col].to_numpy(dtype=float)
        return X["lag_1"].to_numpy(dtype=float)


def build_feature_matrix(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.Series]:
    y = df["qty"].astype(float)
    X = df[
        [
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
    ].copy()
    return X, y


def build_tabular_pipeline(model: Any) -> Pipeline:
    categorical = ["pgrp", "pt"]
    numeric = [
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

    pre = ColumnTransformer(
        transformers=[
            ("cat", OneHotEncoder(handle_unknown="ignore"), categorical),
            ("num", "passthrough", numeric),
        ]
    )

    return Pipeline([("pre", pre), ("model", model)])


def fit_and_score_candidates(
    train_df: pd.DataFrame,
    test_df: pd.DataFrame,
    seed: int,
    candidates: list[str],
) -> tuple[list[ModelResult], ModelResult]:
    X_train, y_train = build_feature_matrix(train_df)
    X_test, y_test = build_feature_matrix(test_df)

    results: list[ModelResult] = []

    if "seasonal_naive" in candidates:
        m = SeasonalNaiveModel(seasonal_lag=12)
        m.fit(X_train, y_train)
        pred = np.clip(m.predict(X_test), 0, None)
        results.append(ModelResult("seasonal_naive", evaluate_predictions(y_test, pd.Series(pred)), m))

    if "random_forest" in candidates:
        rf = RandomForestRegressor(
            n_estimators=500,
            random_state=seed,
            n_jobs=-1,
            min_samples_leaf=2,
        )
        model = build_tabular_pipeline(rf)
        model.fit(X_train, y_train)
        pred = np.clip(model.predict(X_test), 0, None)
        results.append(ModelResult("random_forest", evaluate_predictions(y_test, pd.Series(pred)), model))

    if "xgboost" in candidates and HAS_XGBOOST:
        xgb = XGBRegressor(
            n_estimators=500,
            learning_rate=0.05,
            max_depth=8,
            subsample=0.9,
            colsample_bytree=0.9,
            objective="reg:squarederror",
            random_state=seed,
            n_jobs=4,
        )
        model = build_tabular_pipeline(xgb)
        model.fit(X_train, y_train)
        pred = np.clip(model.predict(X_test), 0, None)
        results.append(ModelResult("xgboost", evaluate_predictions(y_test, pd.Series(pred)), model))

    if not results:
        raise RuntimeError("No models were trained. Check candidate list and dependencies.")

    best = sorted(results, key=lambda r: (r.metrics["wmape"], r.metrics["rmse"]))[0]
    return results, best
