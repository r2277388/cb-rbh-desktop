import unittest

import pandas as pd

from shared.bookscan_calendar import bookscan_week, first_bookscan_sunday


class BookScanCalendarTests(unittest.TestCase):
    def test_2026_new_year_stub_and_first_weeks(self):
        cases = {
            "2026-01-01": (2025, 52, "2025-12-28", "2026-01-03"),
            "2026-01-03": (2025, 52, "2025-12-28", "2026-01-03"),
            "2026-01-04": (2026, 1, "2026-01-04", "2026-01-10"),
            "2026-01-10": (2026, 1, "2026-01-04", "2026-01-10"),
            "2026-01-11": (2026, 2, "2026-01-11", "2026-01-17"),
            "2026-01-17": (2026, 2, "2026-01-11", "2026-01-17"),
        }

        for date_text, expected in cases.items():
            with self.subTest(date=date_text):
                result = bookscan_week(date_text)
                self.assertEqual((result.year, result.week), expected[:2])
                self.assertEqual(result.week_start.strftime("%Y-%m-%d"), expected[2])
                self.assertEqual(result.week_end.strftime("%Y-%m-%d"), expected[3])

    def test_week_one_is_first_sunday_through_saturday_for_many_years(self):
        for year in range(2010, 2041):
            with self.subTest(year=year):
                first_sunday = first_bookscan_sunday(year)
                self.assertEqual(first_sunday.weekday(), 6)

                week_one_start = bookscan_week(first_sunday)
                week_one_end = bookscan_week(first_sunday + pd.Timedelta(days=6))

                self.assertEqual(week_one_start.year, year)
                self.assertEqual(week_one_start.week, 1)
                self.assertEqual(week_one_start.week_start, first_sunday)
                self.assertEqual(week_one_start.week_end, first_sunday + pd.Timedelta(days=6))
                self.assertEqual(week_one_end.year, year)
                self.assertEqual(week_one_end.week, 1)

    def test_dates_before_first_sunday_belong_to_prior_bookscan_year(self):
        for year in range(2010, 2041):
            first_sunday = first_bookscan_sunday(year)
            jan1 = pd.Timestamp(year, 1, 1)
            stub_dates = pd.date_range(jan1, first_sunday - pd.Timedelta(days=1), freq="D")

            for date in stub_dates:
                with self.subTest(date=date.strftime("%Y-%m-%d")):
                    result = bookscan_week(date)
                    self.assertEqual(result.year, year - 1)
                    self.assertLess(result.week_start, first_sunday)
                    self.assertLessEqual(date, result.week_end)

    def test_all_weeks_run_sunday_through_saturday(self):
        for date in pd.date_range("2010-01-01", "2040-12-31", freq="D"):
            with self.subTest(date=date.strftime("%Y-%m-%d")):
                result = bookscan_week(date)
                self.assertEqual(result.week_start.weekday(), 6)
                self.assertEqual(result.week_end.weekday(), 5)
                self.assertEqual((result.week_end - result.week_start).days, 6)
                self.assertGreaterEqual(date.normalize(), result.week_start)
                self.assertLessEqual(date.normalize(), result.week_end)

    def test_year_week_label_is_zero_padded(self):
        self.assertEqual(bookscan_week("2026-01-04").year_week, "2026-01")
        self.assertEqual(bookscan_week("2026-03-15").year_week, "2026-11")


if __name__ == "__main__":
    unittest.main()
