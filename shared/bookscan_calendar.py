from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class BookScanWeek:
    year: int
    week: int
    year_week: str
    week_start: pd.Timestamp
    week_end: pd.Timestamp


def first_bookscan_sunday(year: int) -> pd.Timestamp:
    jan1 = pd.Timestamp(year, 1, 1)
    return jan1 + pd.Timedelta(days=(6 - jan1.weekday()) % 7)


def bookscan_week(date: object) -> BookScanWeek:
    week_end = pd.Timestamp(date).normalize()
    year = int(week_end.year)
    first_sunday = first_bookscan_sunday(year)

    if week_end < first_sunday:
        return bookscan_week(pd.Timestamp(year - 1, 12, 31))

    week = ((week_end - first_sunday).days // 7) + 1
    week_start = first_sunday + pd.Timedelta(weeks=week - 1)
    return BookScanWeek(
        year=year,
        week=int(week),
        year_week=f"{year}-{week:02d}",
        week_start=week_start,
        week_end=week_start + pd.Timedelta(days=6),
    )


def bookscan_parts(dates: pd.Series) -> pd.DataFrame:
    values = dates.map(bookscan_week)
    return pd.DataFrame(
        {
            "BookScanYear": values.map(lambda value: value.year),
            "BookScanWeek": values.map(lambda value: value.week),
            "BookScanWeekStart": values.map(lambda value: value.week_start),
            "BookScanWeekEnd": values.map(lambda value: value.week_end),
        },
        index=dates.index,
    )
