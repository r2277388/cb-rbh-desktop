"""Core forecasting utilities for supply chain sales quantities."""

from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from xgboost import XGBRegressor
import pandas as pd
import numpy as np
import argparse

def load_and_clean_data(pkl_path="sales_data.pkl"):
    """Load a pickled DataFrame and apply standard cleanup."""

    df = pd.read_pickle(pkl_path)
    df.columns = df.columns.str.strip().str.lower()
    df.rename(
        columns={"period": "Period", "isbn": "ISBN", "pgrp": "PGRP", "qty": "Quantity"},
        inplace=True,
    )
    df["Period"] = pd.to_datetime(df["Period"].astype(str), format="%Y%m")
    df["month"] = df["Period"].dt.month
    df["year"] = df["Period"].dt.year
    df = df.sort_values(["ISBN", "Period"])
    return df

def create_lag_features(df, lags=range(1, 13)):
    """Create lagged quantity features for each ISBN."""

    for lag in lags:
        df[f"lag_{lag}"] = df.groupby("ISBN")["Quantity"].shift(lag)
    return df

def filter_isbns(df, min_history=16):
    """Keep ISBNs with at least ``min_history`` observations."""

    valid_isbns = df.groupby("ISBN").size()[lambda x: x >= min_history].index
    return df[df["ISBN"].isin(valid_isbns)].dropna()

def encode_pgrp(df):
    """One-hot encode the ``PGRP`` column."""

    encoder = OneHotEncoder(drop="first", sparse_output=False)
    pgrp_encoded = encoder.fit_transform(df[["PGRP"]])
    pgrp_encoded_df = pd.DataFrame(
        pgrp_encoded, columns=encoder.get_feature_names_out(["PGRP"])
    )
    df = pd.concat([df.reset_index(drop=True), pgrp_encoded_df], axis=1)
    return df, encoder, pgrp_encoded_df

def scale_quantity(df):
    """Standardize the ``Quantity`` column and return the scaler."""

    scaler = StandardScaler()
    df["Quantity_scaled"] = scaler.fit_transform(df[["Quantity"]])
    return df, scaler

def train_test_split(df, lags, pgrp_encoded_df):
    """Split into training and test sets using the last row of each ISBN."""

    test_df = df.groupby("ISBN").tail(1)
    train_df = df.drop(test_df.index)
    lag_cols = [f"lag_{l}" for l in lags]
    feature_cols = lag_cols + ["month", "year"] + list(pgrp_encoded_df.columns)
    X_train = train_df[feature_cols]
    y_train = train_df["Quantity_scaled"]
    X_test = test_df[feature_cols]
    y_test = test_df["Quantity_scaled"]
    return X_train, y_train, X_test, y_test, test_df, feature_cols

def train_xgb(X_train, y_train, **xgb_params):
    """Train an XGBoost regressor."""

    model = XGBRegressor(**xgb_params)
    model.fit(X_train, y_train)
    return model

def evaluate(y_actual, y_pred):
    """Compute regression metrics."""

    mae = mean_absolute_error(y_actual, y_pred)
    mse = mean_squared_error(y_actual, y_pred)
    rmse = np.sqrt(mse)
    r2 = r2_score(y_actual, y_pred)
    return mae, mse, rmse, r2

def get_feature_importance(model, feature_cols):
    """Return feature importances sorted descending."""

    return (
        pd.DataFrame({"Feature": feature_cols, "Importance": model.feature_importances_})
        .sort_values(by="Importance", ascending=False)
    )

def forecast_future(df, model, scaler, encoder, lags, months_ahead=12):
    """Generate forecasts for each ISBN for ``months_ahead`` months."""

    future_results = []
    isbn_groups = df.groupby('ISBN')
    for isbn, group in isbn_groups:
        group = group.sort_values('Period')
        last_row = group.iloc[-1].copy()
        lags_history = list(group['Quantity'].values[-max(lags):])
        pgrp = last_row['PGRP']
        pgrp_encoded = encoder.transform(pd.DataFrame([[pgrp]], columns=['PGRP']))
        last_period = last_row['Period']
        for i in range(1, months_ahead + 1):
            next_period = last_period + pd.DateOffset(months=1)
            month = next_period.month
            year = next_period.year
            lag_features = lags_history[-max(lags):][::-1][:max(lags)]
            lag_features = lag_features[::-1]
            if len(lag_features) < max(lags):
                lag_features = [0] * (max(lags) - len(lag_features)) + lag_features
            features = lag_features + [month, year] + list(pgrp_encoded[0])
            features = np.array(features).reshape(1, -1)
            pred_scaled = model.predict(features)
            pred = scaler.inverse_transform(pred_scaled.reshape(-1, 1)).ravel()[0]
            future_results.append({
                'ISBN': isbn,
                'Period': next_period,
                'Predicted': pred
            })
            lags_history.append(pred)
            last_period = next_period
    return pd.DataFrame(future_results)

def main():
    """Run the forecasting pipeline from the command line."""

    parser = argparse.ArgumentParser(description="Supply chain forecasting")
    parser.add_argument("--pkl-path", default="sales_data.pkl", help="Pickle file with sales data")
    parser.add_argument("--lags", type=int, default=12, help="Number of lag months")
    parser.add_argument("--min-history", type=int, default=16, help="Minimum observations per ISBN")
    parser.add_argument("--months-ahead", type=int, default=12, help="Months to forecast")
    parser.add_argument("--n-estimators", type=int, default=100, help="XGBoost trees")
    parser.add_argument("--learning-rate", type=float, default=0.1, help="XGBoost learning rate")
    parser.add_argument("--random-state", type=int, default=42, help="Random seed")
    args = parser.parse_args()

    lags = list(range(1, args.lags + 1))

    df = load_and_clean_data(args.pkl_path)
    df = create_lag_features(df, lags)
    df = filter_isbns(df, min_history=args.min_history)
    df, encoder, pgrp_encoded_df = encode_pgrp(df)
    df, scaler = scale_quantity(df)
    X_train, y_train, X_test, y_test, test_df, feature_cols = train_test_split(df, lags, pgrp_encoded_df)
    model = train_xgb(
        X_train,
        y_train,
        n_estimators=args.n_estimators,
        learning_rate=args.learning_rate,
        random_state=args.random_state,
    )
    y_pred_scaled = model.predict(X_test)
    y_pred = scaler.inverse_transform(y_pred_scaled.reshape(-1, 1)).ravel()
    y_actual = scaler.inverse_transform(y_test.values.reshape(-1, 1)).ravel()
    mae, mse, rmse, r2 = evaluate(y_actual, y_pred)
    feature_importance = get_feature_importance(model, feature_cols)
    results = test_df[["ISBN", "Period"]].copy()
    results["Actual"] = y_actual
    results["Predicted"] = y_pred

    print()
    print(df.ISBN.nunique(), "unique ISBNs found")
    print()
    print(f"Mean Absolute Error (MAE): {mae:,.0f}")
    print(f"Mean Squared Error (MSE): {mse:,.0f}")
    print(f"Root Mean Squared Error (RMSE): {rmse:,.0f}")
    print(f"R² Score: {r2:.4f}")
    print(results.head())
    print(feature_importance)

    # Forecast and save
    future_forecast = forecast_future(
        df, model, scaler, encoder, lags, months_ahead=args.months_ahead
    )
    print(future_forecast.head(20))
    future_forecast.to_pickle("future_forecast.pkl")
    print("Forecast saved to future_forecast.pkl")

if __name__ == "__main__":
    main()
