from __future__ import annotations

import numpy as np
import pandas as pd


def wmape(y_true: pd.Series, y_pred: pd.Series) -> float:
    denom = np.abs(y_true).sum()
    if denom == 0:
        return float("nan")
    return float(np.abs(y_true - y_pred).sum() / denom)


def evaluate_predictions(actual: pd.Series, pred: pd.Series) -> dict[str, float]:
    err = actual - pred
    mae = float(np.mean(np.abs(err)))
    rmse = float(np.sqrt(np.mean(np.square(err))))
    mape = float(np.mean(np.abs(err) / np.maximum(np.abs(actual), 1e-6)))
    return {
        "mae": mae,
        "rmse": rmse,
        "mape": mape,
        "wmape": wmape(actual, pred),
    }
