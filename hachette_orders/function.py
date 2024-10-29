def adjust_to_weekday(date):
    if date.weekday() == 5:  # Saturday
        return date + pd.DateOffset(days=2)
    elif date.weekday() == 6:  # Sunday
        return date + pd.DateOffset(days=1)
    return date