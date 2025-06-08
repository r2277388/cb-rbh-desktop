import pandas as pd
import numpy as np
from loader_saldet import upload_saldet
from loader_osd import upload_osd
import xgboost as xgb
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
import matplotlib.pyplot as plt
import joblib

def load_and_preprocess_sales_data():
    df_sales = upload_saldet()
    df_sales['WeekStartDate'] = pd.to_datetime(df_sales['WeekStartDate'])
    df_sales = df_sales.groupby(['ISBN', 'WeekStartDate'])['qty'].sum().reset_index()
    return df_sales

def create_features(df):
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
    
    X = df_isbn[['Year', 'Month', 'Day', 'DayOfWeek', 'WeekOfYear']]
    y = df_isbn['qty']
    
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, shuffle=False)
    
    model = xgb.XGBRegressor(objective='reg:squarederror', n_estimators=1000)
    model.fit(X_train, y_train, eval_set=[(X_test, y_test)], early_stopping_rounds=50, verbose=False)
    
    y_pred = model.predict(X_test)
    rmse = np.sqrt(mean_squared_error(y_test, y_pred))
    print(f'RMSE for {isbn}: {rmse}')
    
    return model

def make_forecasts(model, df, isbn, periods=12):
    last_date = df[df['ISBN'] == isbn]['WeekStartDate'].max()
    future_dates = [last_date + pd.Timedelta(weeks=i) for i in range(1, periods + 1)]
    
    future_df = pd.DataFrame({'WeekStartDate': future_dates})
    future_df = create_features(future_df)
    
    X_future = future_df[['Year', 'Month', 'Day', 'DayOfWeek', 'WeekOfYear']]
    future_df['Forecast'] = model.predict(X_future)
    
    return future_df

def main():
    df_sales = load_and_preprocess_sales_data()
    df_osd = upload_osd()
    
    # Test with a specific ISBN
    test_isbn = '9781452179612'
    
    # Ensure the test ISBN is in the sales data and OSD data
    if test_isbn in df_sales['ISBN'].values and test_isbn in df_osd['ISBN'].values:
        model = train_xgboost_model(df_sales, test_isbn)
        future_df = make_forecasts(model, df_sales, test_isbn)
        
        # Add ISBN to the forecast dataframe
        future_df['ISBN'] = test_isbn
        
        # Print the forecast dataframe
        print(future_df)
        
        # Save the forecast dictionary as a pkl file with the bl flag
        forecast_dict = {test_isbn: future_df}
        joblib.dump(forecast_dict, 'forecast_results_bl.pkl')
        
        # Plot the results
        plt.figure(figsize=(10, 6))
        plt.plot(df_sales[df_sales['ISBN'] == test_isbn]['WeekStartDate'], df_sales[df_sales['ISBN'] == test_isbn]['qty'], label='Historical Sales')
        plt.plot(future_df['WeekStartDate'], future_df['Forecast'], label='Forecast')
        plt.legend()
        plt.title(f'Sales Forecast for ISBN {test_isbn}')
        plt.xlabel('Date')
        plt.ylabel('qty')
        plt.show()
    else:
        print(f'ISBN {test_isbn} not found in sales or OSD data.')

if __name__ == '__main__':
    main()