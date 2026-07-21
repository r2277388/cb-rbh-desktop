import pandas as pd

from x_gap_report.main import normalize_isbn, weekly_metrics


def test_normalize_isbn_preserves_leading_zeroes():
    assert normalize_isbn("736313543698") == "0736313543698"
    assert normalize_isbn("978-1-7972-2659-0") == "9781797226590"


def test_weekly_metrics_uses_same_bookscan_week_for_lytd():
    frame = pd.DataFrame({
        "ISBN": ["9781797226590"] * 5,
        "Week": pd.to_datetime(["2025-07-05", "2025-07-12", "2025-07-19", "2026-07-04", "2026-07-11"]),
        "qty": [1, 2, 100, 10, 20],
    })
    result, latest = weekly_metrics(frame, date_col="Week", value_col="qty")
    assert latest == pd.Timestamp("2026-07-11")
    assert result.loc["9781797226590", "Weekly"] == 20
    assert result.loc["9781797226590", "YTD"] == 30
    assert result.loc["9781797226590", "LYTD"] == 3
    assert result.loc["9781797226590", "Last Year"] == 103


def test_weekly_metrics_can_align_lytd_to_a_combined_section_date():
    frame = pd.DataFrame({
        "ISBN": ["9781797226590"] * 4,
        "Week": pd.to_datetime(["2025-07-05", "2025-07-12", "2026-07-04", "2026-07-11"]),
        "qty": [10, 20, 100, 200],
    })
    source = frame[frame["Week"].ne(pd.Timestamp("2026-07-11"))]
    result, latest = weekly_metrics(
        source,
        date_col="Week",
        value_col="qty",
        as_of=pd.Timestamp("2026-07-11"),
    )
    assert latest == pd.Timestamp("2026-07-04")
    assert result.loc["9781797226590", "LYTD"] == 30
