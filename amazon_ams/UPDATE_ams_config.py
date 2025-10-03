import pandas as pd

## Update the end date to represent the new month being added to the reports
month_list = pd.date_range(start='2023-01-01', end='2025-09-01', freq='MS').strftime('%Y-%m').tolist()

# make sure the tab_dict is updated with the new months and files
# This dictionary maps month strings to their respective Excel file paths and sheet names.
tab_dict = {
        ## 2023 AMS Data
        '2023-01': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 01 - January - Performance by ASIN_ALL.xlsx"},
        '2023-02': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 02 - February - Performance by ASIN_ALL.xlsx"},
        '2023-03': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 03 - March - Performance by ASIN_ALL.xlsx"},
        '2023-04': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 04 - April - Performance by ASIN_ALL.xlsx"},
        '2023-05': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 05 - May - Performance by ASIN_ALL.xlsx"},
        '2023-06': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 06 - June - Performance by ASIN_ALL.xlsx"},
        '2023-07': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 07 - July - Performance by ASIN_ALL.xlsx"},
        '2023-08': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 08 - August - Performance by ASIN_ALL.xlsx"},
        '2023-09': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 09 - September - Performance by ASIN_ALL.xlsx"},
        '2023-10': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 10 - October - Performance by ASIN_ALL.xlsx"},
        '2023-11': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 11 - November - Performance by ASIN_ALL.xlsx"},
        '2023-12': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2023\2023- 12 - December - Performance by ASIN_ALL.xlsx"},
        
        ## 2024 AMS Data
        '2024-01': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 01 - January - Performance by ASIN_ALL.xlsx"},
        '2024-02': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 02 - February - Performance by ASIN_ALL.xlsx"},
        '2024-03': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 03 - March - Performance by ASIN_ALL.xlsx"},
        '2024-04': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 04 - April - Performance by ASIN_ALL.xlsx"},
        '2024-05': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 05 - May - Performance by ASIN_ALL.xlsx"},
        '2024-06': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 06 - June - Performance by ASIN_ALL.xlsx"},
        '2024-07': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 07 - July - Performance by ASIN_ALL.xlsx"},
        '2024-08': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 08 - August - Performance by ASIN_ALL.xlsx"},
        '2024-09': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 09 - September - Performance by ASIN_ALL.xlsx"},
        '2024-10': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 10 - October - Performance by ASIN_ALL.xlsx"},
        '2024-11': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 11 - November - Performance by ASIN_ALL.xlsx"},
        '2024-12': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2024\2024- 12 - December - Performance by ASIN_ALL.xlsx"},
        ## 2025 AMS Data
        '2025-01': {'tab':'USE_main','skiprows':1,'file':fr"G:\SALES\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 01 - January- Performance by ASIN_ALL.xlsx"},
    	'2025-02': {'tab':'USE_main','skiprows':1,'file':fr"G:\SALES\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 02 - February - Performance by ASIN_ALL.xlsx"},
    	'2025-03': {'tab':'USE_main','skiprows':1,'file':fr"G:\SALES\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 03 - March - Performance by ASIN_ALL.xlsx"},
        '2025-04': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 04 - April - Performance by ASIN_ALL.xlsx"},
        '2025-05': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 05 - May - Performance by ASIN_ALL.xlsx"},
        '2025-06': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 06 - June - Performance by ASIN_ALL.xlsx"},
        '2025-07': {'tab':'USE_main','skiprows':1,'file':fr"G:\SALES\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 07 - July - Performance by ASIN_ALL.xlsx"},
        '2025-08': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 08 - August - Performance by ASIN_ALL.xlsx"},
        '2025-09': {'tab':'USE_main','skiprows':1,'file':fr"G:\Sales\Amazon\AMAZON ADVERTISING\MONTHLY REPORTS\MONTHLY REPORTS - PERFORMANCE BY ASIN\2025\2025 - 09 - September - Performance by ASIN_ALL.xlsx"}
    }