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