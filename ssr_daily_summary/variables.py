from datetime import datetime as dt, timedelta
from dateutil.relativedelta import relativedelta


def get_default_prior_weekday(today=None):
    """Return the most recent weekday before today."""
    today = today or dt.today()
    prior_day = today.date() - timedelta(days=1)

    while prior_day.weekday() >= 5:
        prior_day -= timedelta(days=1)

    return prior_day


def format_display_date(date_value):
    return date_value.strftime("%A, %Y-%m-%d")


def prompt_for_prior_day():
    default_day = get_default_prior_weekday()
    default_display = format_display_date(default_day)

    while True:
        prior_day_str = input(
            f"Please enter the prior day (yyyy-mm-dd) "
            f"or press Enter for the default [{default_display}]: "
        ).strip()

        if not prior_day_str:
            return default_day

        try:
            prior_day = dt.strptime(prior_day_str, "%Y-%m-%d").date()
        except ValueError:
            print(
                f"{prior_day_str} is not a valid calendar date. "
                "Please enter a real date in yyyy-mm-dd format."
            )
            continue

        if prior_day.weekday() >= 5:
            confirmation = input(
                f"{format_display_date(prior_day)} is a weekend. "
                "SSR reports normally use warehouse weekdays. "
                "Run anyway? [y/N]: "
            ).strip().lower()
            if confirmation not in {"y", "yes"}:
                print("Please choose a Monday-Friday report date.")
                continue

        return prior_day


def get_variables(use_current_date=False):
    """Get the variables needed for the queries. Uses current date if `use_current_date` is True."""
    if use_current_date:
        # Use today's date
        prior_day_dt = dt.today()
    else:
        prior_day_dt = prompt_for_prior_day()

    # Correct date formatting for the variables
    prior_day = prior_day_dt.strftime('%Y-%m-%d')  # For date comparisons
    prior_day_ly_dt = prior_day_dt - relativedelta(years=1)  # Prior year as a datetime object
    prior_day_ly = prior_day_ly_dt.strftime('%Y-%m-%d')  # Format to string for return
    
    tp = prior_day_dt.strftime('%Y%m')  # Format for 'YYYYMM'
    typ1 = prior_day_dt.strftime('%Y') + '01'  # First month of the current year in 'YYYYMM'
    
    tply = prior_day_ly_dt.strftime('%Y%m')  # Prior year in 'YYYYMM'
    lyp1 = prior_day_ly_dt.strftime('%Y') + '01'  # First month of the prior year in 'YYYYMM'
    
    # Return the correct variables
    return prior_day, prior_day_ly, tp, typ1, tply, lyp1

def main():
    # Correct unpacking of the returned values
    variables = get_variables()
    if variables:
        prior_day, prior_day_ly, tp, typ1, tply, lyp1 = variables
        print()
        print(f"prior_day: {prior_day}")
        print(f"prior_day_ly: {prior_day_ly}")
        print()
        print(f"tp: {tp}")
        print(f"typ1: {typ1}")
        print()
        print(f"tply: {tply}")
        print(f"lyp1: {lyp1}")
        print()
        
if __name__ == '__main__':
    main()  # Call the main function
