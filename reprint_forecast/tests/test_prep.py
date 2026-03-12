from __future__ import annotations

import unittest

import pandas as pd

from reprint_forecast.data.prep import (
    filter_backlist_min_history,
    join_sales_orders,
    prep_orders,
    prep_sales,
)


class PrepTests(unittest.TestCase):
    def test_prep_sales_parses_period_and_fills_missing_months(self) -> None:
        raw = pd.DataFrame(
            {
                "period": ["202401", "202403"],
                "isbn": ["111", "111"],
                "pgrp": ["chl", "chl"],
                "pt": ["bk", "bk"],
                "ssr_id": [1, 1],
                "qty": [10, 30],
                "val": [100, 300],
            }
        )
        out = prep_sales(raw, frequency="MS")

        self.assertEqual(len(out), 3)
        self.assertEqual(out["period"].dt.strftime("%Y-%m").tolist(), ["2024-01", "2024-02", "2024-03"])
        feb = out[out["period"] == pd.Timestamp("2024-02-01")].iloc[0]
        self.assertEqual(feb["qty"], 0)

    def test_join_sales_orders_aligns_entered_date_to_period(self) -> None:
        sales_raw = pd.DataFrame(
            {
                "period": ["202401", "202402"],
                "isbn": ["111", "111"],
                "pgrp": ["CHL", "CHL"],
                "pt": ["BK", "BK"],
                "ssr_id": [1, 1],
                "qty": [10, 20],
                "val": [100, 200],
            }
        )
        orders_raw = pd.DataFrame(
            {
                "ssr_id": [1],
                "isbn": ["111"],
                "pgrp": ["CHL"],
                "entered_date": ["2024-02-15"],
                "release_date": ["2024-02-20"],
                "order_cancel_date": ["2024-02-28"],
                "order_type_code": ["REGULAR"],
                "qty": [7],
            }
        )

        sales = prep_sales(sales_raw, frequency="MS")
        orders = prep_orders(orders_raw)
        out = join_sales_orders(sales, orders, include_orders=True, frequency="MS")

        feb = out[out["period"] == pd.Timestamp("2024-02-01")].iloc[0]
        jan = out[out["period"] == pd.Timestamp("2024-01-01")].iloc[0]
        self.assertEqual(feb["order_qty"], 7)
        self.assertEqual(jan["order_qty"], 0)

    def test_filter_backlist_min_history_keeps_only_long_history_isbns(self) -> None:
        periods = pd.date_range("2024-01-01", periods=12, freq="MS")
        long_rows = pd.DataFrame(
            {
                "period": periods,
                "isbn": ["111"] * len(periods),
                "ssr_id": [1] * len(periods),
                "pgrp": ["CHL"] * len(periods),
                "pt": ["BK"] * len(periods),
                "qty": [1] * len(periods),
                "val": [1] * len(periods),
                "order_qty": [0] * len(periods),
            }
        )
        short_rows = pd.DataFrame(
            {
                "period": periods[:6],
                "isbn": ["222"] * 6,
                "ssr_id": [1] * 6,
                "pgrp": ["ENT"] * 6,
                "pt": ["BK"] * 6,
                "qty": [1] * 6,
                "val": [1] * 6,
                "order_qty": [0] * 6,
            }
        )
        out = filter_backlist_min_history(pd.concat([long_rows, short_rows], ignore_index=True), 12)
        self.assertEqual(set(out["isbn"].unique()), {"111"})


if __name__ == "__main__":
    unittest.main()

