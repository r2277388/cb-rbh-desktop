import pandas as pd
import random

def adjust_to_weekday(date):
    today = pd.Timestamp('today').normalize()  # Normalize to remove time component
    later_date = max(date, today)
    
    # Adjust to the next weekday if it falls on a weekend
    if later_date.weekday() == 5:  # Saturday
        return later_date + pd.DateOffset(days=2)
    elif later_date.weekday() == 6:  # Sunday
        return later_date + pd.DateOffset(days=1)
    return later_date

def get_later_date(date):
    today = pd.Timestamp('today').normalize()  # Normalize to remove time component
    later_date = max(date, today)
    return adjust_to_weekday(later_date)

def today():
    return pd.Timestamp('today').normalize()

def next_three_days():
    today = pd.Timestamp('today').normalize()
    tomorrow = today + pd.DateOffset(days=1)
    next_day = today + pd.DateOffset(days=2)
    following_day = today + pd.DateOffset(days=3)
    
    # List of possible dates
    possible_dates = [tomorrow, next_day, following_day]
    
    # Select a date based on probability
    selected_date = random.choices(possible_dates, weights=[0.33, 0.33, 0.34], k=1)[0]
    
    # Adjust the selected date to ensure it does not fall on a weekend
    return adjust_to_weekday(selected_date)

# Function to summarize the "val" field by EstimateDate
def summarize_by_estimate_date(df):
    todays_date = pd.Timestamp('today').normalize()
    next_5_days = todays_date + pd.DateOffset(days=5)
    end_of_month = todays_date + pd.offsets.MonthEnd(0)
    end_of_3_months = todays_date + pd.DateOffset(months=3)

    # Filter data for the next 5 days, rest of the month, and next 3 months
    next_5_days_df = df[(df['EstimateDate'] >= todays_date) & (df['EstimateDate'] <= next_5_days)]
    rest_of_month_df = df[(df['EstimateDate'] > next_5_days) & (df['EstimateDate'] <= end_of_month)]
    next_3_months_df = df[(df['EstimateDate'] > end_of_month) & (df['EstimateDate'] <= end_of_3_months)]

    # Summarize the "val" field
    daily_summary = next_5_days_df.groupby('EstimateDate')['val'].sum().reset_index()
    daily_summary['EstimateDate'] = daily_summary['EstimateDate'].dt.normalize()
    daily_summary['val'] = daily_summary['val'].apply(lambda x: f"{x:,}")

    summary = {
        'Next 5 Days Total': f"{next_5_days_df['val'].sum():,}",
        'Rest of the Month': f"{rest_of_month_df['val'].sum():,}",
        'Next 3 Months': f"{next_3_months_df['val'].sum():,}"
    }

    return daily_summary, summary
