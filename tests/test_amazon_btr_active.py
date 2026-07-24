import pandas as pd

from amazon_btr.workbook import build_active_asins


def test_active_asins_are_accepted_active_and_within_sell_through_window():
    cleaned = pd.DataFrame(
        [
            {
                "ASIN": "ACTIVE",
                "Offer state": "Active",
                "Status": "Accepted",
                "Sell-through start date": "2026-07-01",
                "Sell-through end date": "2026-07-31",
            },
            {
                "ASIN": "ENDED",
                "Offer state": "Active",
                "Status": "Accepted",
                "Sell-through start date": "2026-06-01",
                "Sell-through end date": "2026-07-22",
            },
            {
                "ASIN": "COMPLETED",
                "Offer state": "Completed",
                "Status": "Accepted",
                "Sell-through start date": "2026-07-01",
                "Sell-through end date": "2026-08-01",
            },
        ]
    )

    active = build_active_asins(cleaned, pd.Timestamp("2026-07-23"))

    assert active["ASIN"].tolist() == ["ACTIVE"]
    assert active["Days Left"].tolist() == [8]
