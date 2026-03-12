from __future__ import annotations

from pathlib import Path

import joblib
import pandas as pd

from reprint_forecast.config import AppConfig
from reprint_forecast.data.cache import load_or_none, save_df
from reprint_forecast.data.prep import (
    filter_backlist_min_history,
    join_sales_orders,
    prep_orders,
    prep_sales,
)
from reprint_forecast.data.query import query_df
from reprint_forecast.features.build_features import build_future_rows, make_supervised
from reprint_forecast.models.train import build_feature_matrix, fit_and_score_candidates


def _cache_paths(cfg: AppConfig) -> dict[str, Path]:
    cache_dir = cfg.root / cfg.raw["paths"]["cache_dir"]
    return {
        "sales": cache_dir / "sales.pkl",
        "orders": cache_dir / "orders.pkl",
        "model_input": cache_dir / "model_input.pkl",
    }


def _load_raw_from_cache(cfg: AppConfig) -> tuple[pd.DataFrame, pd.DataFrame]:
    paths = _cache_paths(cfg)
    sales = load_or_none(paths["sales"])
    orders = load_or_none(paths["orders"])
    if sales is None or orders is None:
        raise RuntimeError("Missing cached raw data. Run `python -m reprint_forecast.cli refresh-data` first.")
    return sales, orders


def refresh_raw_data(cfg: AppConfig, force: bool = False) -> None:
    params = cfg.raw["query"]
    paths = _cache_paths(cfg)

    for key, sql_cfg_key in (("sales", "sales_sql"), ("orders", "orders_sql")):
        cached = load_or_none(paths[key])
        if cached is not None and not force:
            print(f"Using cached {key}: {paths[key]}")
            continue

        sql_path = cfg.root / cfg.raw["paths"][sql_cfg_key]
        try:
            df = query_df(sql_path, params)
            save_df(df, paths[key])
            print(f"Refreshed {key} cache: {paths[key]} ({len(df):,} rows)")
        except Exception as ex:
            if cached is not None:
                print(f"Query failed for {key}, falling back to cache. Error: {ex}")
                continue
            raise


def _prepare_model_input(cfg: AppConfig, refresh_first: bool) -> pd.DataFrame:
    if refresh_first:
        refresh_raw_data(cfg, force=True)

    sales_raw, orders_raw = _load_raw_from_cache(cfg)

    sales = prep_sales(sales_raw, frequency=cfg.raw["pipeline"]["frequency"])
    orders = prep_orders(orders_raw)
    merged = join_sales_orders(
        sales,
        orders,
        include_orders=bool(cfg.raw["pipeline"]["include_orders_as_feature"]),
        frequency=cfg.raw["pipeline"]["frequency"],
    )
    filtered = filter_backlist_min_history(merged, int(cfg.raw["pipeline"]["min_history_periods"]))

    save_df(filtered, _cache_paths(cfg)["model_input"])
    return filtered


def _fit_best_model(cfg: AppConfig, supervised: pd.DataFrame):
    test_year = int(cfg.raw["pipeline"]["test_year"])
    candidates = list(cfg.raw["models"]["candidates"])

    train_df = supervised[supervised["period"].dt.year < test_year].copy()
    test_df = supervised[supervised["period"].dt.year == test_year].copy()

    if train_df.empty or test_df.empty:
        raise RuntimeError(f"Train/test split for test_year={test_year} is empty. Adjust data window or config.")

    results, best = fit_and_score_candidates(
        train_df=train_df,
        test_df=test_df,
        seed=cfg.seed,
        candidates=candidates,
    )

    metrics = pd.DataFrame(
        [{"model": r.name, **r.metrics} for r in results]
    ).sort_values(["wmape", "rmse"], ascending=True)

    # Refit best on all available supervised data before forecasting.
    X_all, y_all = build_feature_matrix(supervised)
    if hasattr(best.model, "fit"):
        best.model.fit(X_all, y_all)

    return best, metrics


def _recursive_forecast(history: pd.DataFrame, model, horizon: int, freq: str) -> pd.DataFrame:
    work = history.copy().sort_values(["isbn", "ssr_id", "period"]).reset_index(drop=True)
    last_actual = work["period"].max()

    future_seed = build_future_rows(work, horizon=horizon, freq=freq)
    for dt in sorted(future_seed["period"].unique()):
        next_rows = future_seed[future_seed["period"] == dt]
        work = pd.concat([work, next_rows], ignore_index=True)

        sup = make_supervised(work)
        pred_rows = sup[sup["period"] == dt].copy()
        if pred_rows.empty:
            continue

        X_pred, _ = build_feature_matrix(pred_rows.fillna({"qty": 0}))
        yhat = model.predict(X_pred)
        pred_rows["qty_pred"] = yhat.clip(min=0)

        key = pred_rows[["period", "isbn", "ssr_id", "qty_pred"]]
        work = work.merge(key, on=["period", "isbn", "ssr_id"], how="left")
        work["qty"] = work["qty"].fillna(work["qty_pred"])
        work = work.drop(columns=["qty_pred"])

    forecast = work[work["period"] > last_actual].copy()
    forecast = forecast[["period", "isbn", "ssr_id", "pgrp", "pt", "qty", "order_qty"]]
    forecast = forecast.rename(columns={"qty": "forecast_qty"})
    return forecast.sort_values(["period", "isbn", "ssr_id"]) 


def _write_inventory_coverage(cfg: AppConfig, forecast: pd.DataFrame) -> None:
    inventory_path = cfg.root / cfg.raw["paths"]["inventory_snapshot"]
    if not inventory_path.exists():
        return

    inv = pd.read_csv(inventory_path)
    inv.columns = [c.lower() for c in inv.columns]
    if not {"isbn", "on_hand_qty"}.issubset(inv.columns):
        return

    horizon = int(cfg.raw["pipeline"]["forecast_horizon_periods"])
    by_isbn = forecast.groupby("isbn", as_index=False).agg(forecast_qty_6m=("forecast_qty", "sum"))
    out = inv.merge(by_isbn, on="isbn", how="left").fillna({"forecast_qty_6m": 0})

    out["avg_period_demand"] = out["forecast_qty_6m"] / max(horizon, 1)
    out["months_cover"] = out["on_hand_qty"] / out["avg_period_demand"].replace(0, pd.NA)
    out["has_6_month_supply"] = out["months_cover"].fillna(float("inf")) >= 6

    out_path = cfg.output_dir / "inventory_coverage.csv"
    out.to_csv(out_path, index=False)
    print(f"Wrote inventory coverage: {out_path}")


def run_full_pipeline(cfg: AppConfig, refresh_first: bool = False) -> None:
    cfg.output_dir.mkdir(parents=True, exist_ok=True)

    model_input = _prepare_model_input(cfg, refresh_first=refresh_first)
    supervised = make_supervised(model_input)

    best, metrics = _fit_best_model(cfg, supervised)

    metrics_path = cfg.output_dir / "metrics_2025.csv"
    metrics.to_csv(metrics_path, index=False)
    print(f"Best model: {best.name}")
    print(f"Wrote metrics: {metrics_path}")

    model_path = cfg.output_dir / "best_model.joblib"
    joblib.dump(best.model, model_path)

    horizon = int(cfg.raw["pipeline"]["forecast_horizon_periods"])
    freq = cfg.raw["pipeline"]["frequency"]
    forecast = _recursive_forecast(model_input, best.model, horizon=horizon, freq=freq)

    forecast_path = cfg.output_dir / "forecast_6m.csv"
    forecast.to_csv(forecast_path, index=False)
    print(f"Wrote forecast: {forecast_path}")

    _write_inventory_coverage(cfg, forecast)
