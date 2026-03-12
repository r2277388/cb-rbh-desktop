# Reprint Forecast

Deterministic 6-month demand forecasting for backlist titles using historical sales + Hachette orders.

## Goals
- Weekly-refreshable data pipeline with cached query results (`.pkl`) to avoid expensive reruns.
- Model comparison on 2025 holdout (including XGBoost).
- 6-month forward forecast by ISBN / customer group.
- Inventory coverage check when inventory snapshot is provided.

## Project Layout
- `sql/`: source SQL queries.
- `src/reprint_forecast/data/`: query, cache, and prep logic.
- `src/reprint_forecast/features/`: feature engineering.
- `src/reprint_forecast/models/`: model training and evaluation.
- `src/reprint_forecast/pipelines/`: end-to-end orchestration.
- `artifacts/`: outputs (metrics, forecasts, chosen model).
- `data/cache/`: pickled source data and prepared frames.

## Setup
```powershell
cd c:\Users\rbh\code\reprint_forecast
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install -e .
```

## Configuration
1. Copy `.env.example` to `.env` and set DB credentials.
2. Adjust `config.yaml` for frequency, top ISBN count, and forecast horizon.

## Shared DB Connection (Centralized)
This project will automatically try to use the centralized connection helper:
- `c:\Users\rbh\code\shared\db\connection.py`

If `shared` is available, SQL execution uses `shared.db.connection.get_connection()`.
If not, it falls back to local `.env` values (`DB_SERVER`, `DB_DATABASE`, etc.).

## Usage
Refresh or load cached raw data:
```powershell
python -m reprint_forecast.cli refresh-data --config config.yaml
```

Run model comparison on 2025 and produce 6-month forecast:
```powershell
python -m reprint_forecast.cli run --config config.yaml
```

## Weekly Runbook
1. Update SQL templates in:
   - `sql/sales.sql`
   - `sql/hachette_orders.sql`
2. Refresh cached source data (force rerun of SQL):
```powershell
python -m reprint_forecast.cli refresh-data --config config.yaml --force
```
3. Train + compare models on 2025 holdout + write 6-month forecast:
```powershell
python -m reprint_forecast.cli run --config config.yaml
```
4. Review outputs in `artifacts/`:
   - `metrics_2025.csv` (model comparison)
   - `best_model.joblib`
   - `forecast_6m.csv`
   - `inventory_coverage.csv` (if inventory snapshot exists)

## Cached PKLs
Source query cache is written to `data/cache/`:
- `sales.pkl`
- `orders.pkl`
- `model_input.pkl` (post-prep modeling frame)

`refresh-data` only reruns SQL when cache is missing or `--force` is provided.

Current default is monthly (`frequency: MS`) because the provided sales SQL returns monthly data (`period` = yyyymm). If you switch to daily sales SQL, update frequency to weekly and adjust feature windows.

## Determinism
- Fixed random seed in config.
- Stable temporal split (`test_year: 2025`).
- Sorted inputs before modeling.

## Inventory Coverage
If `data/raw/inventory_snapshot.csv` exists with columns:
- `isbn`
- `on_hand_qty`

pipeline will generate `artifacts/inventory_coverage.csv` with weeks/months of cover based on forecasted demand.
