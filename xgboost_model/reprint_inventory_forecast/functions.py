import pandas as pd

def to_saturday(date):
    date = pd.to_datetime(date)
    return date + pd.offsets.Week(weekday=5)