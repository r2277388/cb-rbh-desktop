from __future__ import annotations

import pandas as pd


CRITERIA_ROWS = [
    {
        "Tab": "Raw_BTR",
        "Purpose": "Preserves the current Amazon Vendor Central download.",
        "Criteria / Logic": (
            "Contains every row and field from the newer source workbook without "
            "ASIN filtering or status comparison. ISBN uses the shared Amazon "
            "resolver: manual overrides, then the latest Catalog's ISBN-13, EAN, "
            "and Model Number, followed by a direct ISBN or ISBN-10 conversion when "
            "the ASIN itself is an ISBN. Title and Publisher come from ebs.Item."
        ),
    },
    {
        "Tab": "Cleaned",
        "Purpose": "Provides one meaningful, most-recent row per ASIN.",
        "Criteria / Logic": (
            "For each ASIN, keeps the row with the latest Submission date. If that "
            "row has Status = 'Rejected: Active offer', keeps the most recent earlier "
            "row with Status = 'Active' or 'Accepted' when one exists. 'Accepted' is "
            "included because current Amazon exports use that value rather than "
            "'Active' in the Status field."
        ),
    },
    {
        "Tab": "Status Changes",
        "Purpose": "Shows meaningful differences between the previous and current files.",
        "Criteria / Logic": (
            "Compares the Cleaned rows by ASIN. Includes ASINs whose Status changed "
            "and ASINs appearing for the first time. ASINs missing from the newer "
            "six-month download are ignored. Each result includes the current complete "
            "source row plus Change Type, Previous Status, and Previous Status "
            "description. Results are grouped by current Status, with the smallest "
            "status group first and headers repeated for each group. Status and "
            "Status description appear first, followed by comparison fields and "
            "the mapped ISBN, Title, and Publisher."
        ),
    },
    {
        "Tab": "Active ASINs",
        "Purpose": "Identifies accepted offers currently within their sell-through window.",
        "Criteria / Logic": (
            "Uses the Cleaned tab and keeps rows where Offer state = 'Active', "
            "Status = 'Accepted', and the report date falls on or between the "
            "Sell-through start date and Sell-through end date. Days Left equals "
            "Sell-through end date minus the report date; an offer ending on the "
            "report date has 0 days left."
        ),
    },
    {
        "Tab": "Criteria",
        "Purpose": "Documents how the report is constructed.",
        "Criteria / Logic": (
            "The report date and output filename are derived from the timestamp in "
            "the newer Amazon export filename."
        ),
    },
]


def write_criteria_sheet(writer: pd.ExcelWriter) -> None:
    sheet_name = "Criteria"
    criteria = pd.DataFrame(CRITERIA_ROWS)
    criteria.to_excel(writer, sheet_name=sheet_name, index=False)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E78",
            "border": 1,
            "text_wrap": True,
            "valign": "top",
        }
    )
    body_format = workbook.add_format(
        {"text_wrap": True, "valign": "top", "border": 1}
    )
    for column_number, column_name in enumerate(criteria.columns):
        worksheet.write(0, column_number, column_name, header_format)
    worksheet.set_column(0, 0, 18, body_format)
    worksheet.set_column(1, 1, 38, body_format)
    worksheet.set_column(2, 2, 110, body_format)
    worksheet.set_row(0, 30)
    for row_number in range(1, len(criteria) + 1):
        worksheet.set_row(row_number, 72)
    worksheet.freeze_panes(1, 0)
