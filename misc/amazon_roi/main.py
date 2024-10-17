import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression, Ridge, Lasso, ElasticNet
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.neighbors import KNeighborsRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.preprocessing import StandardScaler
import xgboost as xgb
from data_processing import data_prep

# Prepare the data
df_merged = data_prep()

# Drop the original pub_date if you don't need it anymore
df_merged.drop(columns=['Month','pub_date','Period'], inplace=True)
    
# One-hot encode categorical columns
# Note: pub & pgrp were removed from the SQL query
df_merged = pd.get_dummies(df_merged, columns=['SeasonOnly','PT'], drop_first=True)

# Define the target variable (y)
y = df_merged['Ordered Revenue']

# Identify numeric columns (excluding the target)
numeric_columns = df_merged.select_dtypes(include=[np.number]).columns.drop(['Ordered Revenue'])

# Scale only the numeric columns
scaler = StandardScaler()
df_merged[numeric_columns] = scaler.fit_transform(df_merged[numeric_columns])

# Define the feature matrix (X) - Exclude the target and identifiers
X = df_merged.drop(columns=['Ordered Revenue', 'ISBN'])

def evaluate_model(model, X_train, X_test, y_train, y_test, model_name, results):
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    mse = mean_squared_error(y_test, y_pred)
    r2 = r2_score(y_test, y_pred)
    print(f"\n--- {model_name} ---")
    print(f'{model_name} - Mean Squared Error: {mse:,.0f}')  # Format MSE with commas and no decimals
    print(f'{model_name} - R-squared: {r2:.2f}')
    results[model_name] = {'MSE': mse, 'R-squared': r2}  # Store the MSE and R-squared value
    if hasattr(model, 'coef_'):
        coefficients = pd.DataFrame(model.coef_, X.columns, columns=['Coefficient'])
        print(f"{model_name} - Coefficients:")
        print(coefficients)
    if hasattr(model, 'feature_importances_'):
        feature_importance = pd.DataFrame(model.feature_importances_, index=X.columns, columns=['Importance']).sort_values('Importance', ascending=False)
        print(f"{model_name} - Feature Importance:")
        print(feature_importance)

def main():
    # Split the data
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=42)
    
    # Dictionary to store MSE and R-squared values
    results = {}

    # Linear Regression
    evaluate_model(LinearRegression(), X_train, X_test, y_train, y_test, "Linear Regression", results)
    
    # Ridge Regression
    evaluate_model(Ridge(alpha=1.0), X_train, X_test, y_train, y_test, "Ridge Regression", results)
    
    # Lasso Regression
    evaluate_model(Lasso(alpha=0.1), X_train, X_test, y_train, y_test, "Lasso Regression", results)
    
    # Elastic Net
    evaluate_model(ElasticNet(alpha=0.1, l1_ratio=0.5), X_train, X_test, y_train, y_test, "Elastic Net", results)
    
    # Random Forest
    evaluate_model(RandomForestRegressor(n_estimators=100, random_state=42), X_train, X_test, y_train, y_test, "Random Forest", results)
    
    # Gradient Boosting Machine
    evaluate_model(GradientBoostingRegressor(n_estimators=100, learning_rate=0.1, random_state=42), X_train, X_test, y_train, y_test, "Gradient Boosting Machine", results)

    # XGBoost
    evaluate_model(xgb.XGBRegressor(n_estimators=100, learning_rate=0.1, random_state=42), X_train, X_test, y_train, y_test, "XGBoost", results)
    
    # K-Nearest Neighbors Regression (KNN)
    evaluate_model(KNeighborsRegressor(n_neighbors=5), X_train, X_test, y_train, y_test, "K-Nearest Neighbors Regression", results)
    
    # Print all MSE and R-squared values together
    print("\n--- MSE and R-squared Comparison ---")
    for model_name, metrics in results.items():
        print(f'{model_name}: MSE = {metrics["MSE"]:,.0f}, R-squared = {metrics["R-squared"]:.2f}')

if __name__ == '__main__':
    main()