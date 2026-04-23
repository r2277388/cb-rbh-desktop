

1) load_dataframe_from_pickle():
	return df_raw
2) clean_raw(df_raw):
	returns a cleaned up df with only a date and y (revenue) based on filters like pgrp, channel, ssr, or flbl

3) filter_outliers(df, threshold=7): 
	returns df but without outliers, default is 7 std

4) create_features(df):
	return df: like dayofweek, quarter, month etc

5) add_lags(df):
	returns df: 1, 2, and 3 year lags

6) target_features(df):
	returns two lists:
		target list (which is 'y')
		features list

(7) create_future_with_features(df):
	returns: returns df with future dates for next year, it also adds the features and lags as well as removes holidays
		this also seems to remove even last year data so we are only looking at current and next year info.

(8) df_acts_fcts(df, future_df, year_starting = 2023)
	returns a df with the data axis, a y column from the actuals, a 'pred' column from the future_df
		and combines so we have one df with a y and pred columns.


(9) reg_function(df): (note, reg = regressor)
	returns df by create_futures, add_lags, filter_outliers, 
		creating target and features lists
		applying xgbregressor

(9) generate_future_predictions(df):
	returns df_predictions
		applies clean_raw, create_features, add_lags, reg_function, df_acts_fcts

(10) multiple_forecasts(df):
		returns df with several forecasts.
			looks through each pgrp and produces a file with predictions

(11) multiple_forecast_combined(df)
	returns df but one for FL and one for BL

(12) 


	