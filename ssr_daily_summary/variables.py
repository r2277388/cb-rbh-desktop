from datetime import datetime as dt
from dateutil.relativedelta import relativedelta

def get_variables(use_current_date=False):
    """Get the variables needed for the queries. Uses current date if `use_current_date` is True."""
    if use_current_date:
        # Use today's date
        prior_day_dt = dt.today()
    else:
        # Prompt for date input
        try:
            prior_day_str = input('Please enter the prior day (yyyy-mm-dd): ')
            prior_day_dt = dt.strptime(prior_day_str, '%Y-%m-%d')
        except ValueError:
            print("Invalid date format. Please enter the date in yyyy-mm-dd format.")
            return None

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