import pandas as pd
import numpy as np
from loader_saldet import upload_saldet
from loader_osd import upload_osd
from loader_inventory import load_inventory
from loader_hachetteorders import load_hachette_orders
import xgboost as xgb
from sklearn.model_selection import TimeSeriesSplit
from sklearn.metrics import mean_squared_error
import joblib

FEATURE_COLS = [
    'Year',
    'Month',
    'Day',
    'DayOfWeek',
    'WeekOfYear',
    'Lag1',
    'WeeksSinceOSD',
    'Available To Sell',
    'Frozen',
    'Reprint Freeze',
    'Quantity',
    'OrderOpen',
]

def load_and_preprocess_sales_data():
    df_sales = upload_saldet()
    df_sales['WeekStartDate'] = pd.to_datetime(df_sales['WeekStartDate'])
    df_sales = df_sales.groupby(['ISBN', 'WeekStartDate'])['qty'].sum().reset_index()
    return df_sales

def merge_metadata(df_sales, df_osd, df_inventory, df_orders):
    df = df_sales.merge(df_osd, on='ISBN', how='left')
    df = df.merge(
        df_inventory[['ISBN', 'Available To Sell', 'Frozen', 'Reprint Freeze']],
        on='ISBN',
        how='left',
    )
    df = df.merge(df_orders[['ISBN', 'Order Status', 'Quantity']], on='ISBN', how='left')

    df['OSD'] = pd.to_datetime(df['OSD'])
    df['WeeksSinceOSD'] = ((df['WeekStartDate'] - df['OSD']).dt.days / 7).fillna(0)
    df['Available To Sell'] = df['Available To Sell'].fillna(0)
    df['Quantity'] = df['Quantity'].fillna(0)
    df['Frozen'] = df['Frozen'].fillna(False).astype(int)
    df['Reprint Freeze'] = df['Reprint Freeze'].fillna(False).astype(int)
    df['OrderOpen'] = df['Order Status'].fillna('').str.contains('Open', case=False).astype(int)
    df.drop(columns=['Order Status'], inplace=True)
    return df

def create_features(df):
    df = df.sort_values(['ISBN', 'WeekStartDate'])
    df['Lag1'] = df.groupby('ISBN')['qty'].shift(1).fillna(0)
    df['Year'] = df['WeekStartDate'].dt.year
    df['Month'] = df['WeekStartDate'].dt.month
    df['Day'] = df['WeekStartDate'].dt.day
    df['DayOfWeek'] = df['WeekStartDate'].dt.dayofweek
    df['WeekOfYear'] = df['WeekStartDate'].dt.isocalendar().week
    return df

def get_top_isbns(df_sales, df_osd, n=10):
    # Filter OSD data for titles older than 1 year
    one_year_ago = pd.to_datetime('today') - pd.DateOffset(years=1)
    df_osd_filtered = df_osd[df_osd['OSD'] < one_year_ago]

    # Merge sales data with OSD data
    df_merged = pd.merge(df_sales, df_osd_filtered, on='ISBN',how = 'inner')

    # Get top ISBNs based on sales in the last 3 months
    three_months_ago = pd.to_datetime('today') - pd.DateOffset(months=3)
    recent_sales = df_merged[df_merged['WeekStartDate'] >= three_months_ago]
    top_isbns = recent_sales.groupby('ISBN')['qty'].sum().nlargest(n).index
    return top_isbns

def train_xgboost_model(df, isbn):
    df_isbn = df[df['ISBN'] == isbn]
    df_isbn = create_features(df_isbn)

    available_features = [c for c in FEATURE_COLS if c in df_isbn.columns]
    X = df_isbn[available_features]
    y = df_isbn['qty']

    tscv = TimeSeriesSplit(n_splits=5)
    train_index, test_index = list(tscv.split(X))[-1]
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y.iloc[train_index], y.iloc[test_index]

    model = xgb.XGBRegressor(objective='reg:squarederror', n_estimators=1000)
    model.fit(X_train, y_train, eval_set=[(X_test, y_test)], early_stopping_rounds=50, verbose=False)

    y_pred = model.predict(X_test)
    rmse = np.sqrt(mean_squared_error(y_test, y_pred))
    print(f'RMSE for {isbn}: {rmse}')

    model.fit(X, y)
    return model

def make_forecasts(model, df, isbn, periods=12):
    last_date = df[df['ISBN'] == isbn]['WeekStartDate'].max()
    future_dates = [last_date + pd.Timedelta(weeks=i) for i in range(1, periods + 1)]

    future_df = pd.DataFrame({'WeekStartDate': future_dates})
    future_df['ISBN'] = isbn

    osd_date = df[df['ISBN'] == isbn]['OSD'].iloc[-1] if 'OSD' in df.columns else pd.NaT
    future_df['OSD'] = osd_date
    future_df['WeeksSinceOSD'] = ((future_df['WeekStartDate'] - future_df['OSD']).dt.days / 7).fillna(0)

    for col in ['Available To Sell', 'Frozen', 'Reprint Freeze', 'Quantity', 'OrderOpen']:
        if col in df.columns:
            future_df[col] = df[df['ISBN'] == isbn][col].iloc[-1]

    future_df = create_features(future_df)
    available_features = [c for c in FEATURE_COLS if c in future_df.columns]
    X_future = future_df[available_features]
    future_df['Forecast'] = model.predict(X_future)

    return future_df

def forecast_multiple_isbns(df, isbns, periods=12):
    forecasts = []
    for isbn in isbns:
        if isbn not in df['ISBN'].values:
            print(f'ISBN {isbn} not found in data.')
            continue
        model = train_xgboost_model(df, isbn)
        future_df = make_forecasts(model, df, isbn, periods=periods)
        forecasts.append(future_df)

    if not forecasts:
        return [], pd.DataFrame()

    all_forecasts = pd.concat(forecasts, ignore_index=True)
    aggregated = all_forecasts.groupby('WeekStartDate')['Forecast'].sum().reset_index()
    aggregated.rename(columns={'Forecast': 'AggregatedForecast'}, inplace=True)
    return forecasts, aggregated

def main():
    df_sales = load_and_preprocess_sales_data()
    df_osd = upload_osd()
    df_inv = load_inventory()
    df_orders = load_hachette_orders()

    df_all = merge_metadata(df_sales, df_osd, df_inv, df_orders)

    top_isbns = get_top_isbns(df_sales, df_osd, n=5)
    forecasts, aggregated = forecast_multiple_isbns(df_all, top_isbns)

    forecast_dict = {isbn: f_df for isbn, f_df in zip(top_isbns, forecasts)}
    forecast_dict['aggregated'] = aggregated
    joblib.dump(forecast_dict, 'forecast_results_bl.pkl')

    print(aggregated)

if __name__ == '__main__':
    main()
