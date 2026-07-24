from pathlib import Path

import pandas as pd

from amazon_btr.main import build_status_changes, clean_btr_data, find_latest_two_files


def _row(asin: str, date: str, status: str, description: str = "") -> dict:
    return {
        "ASIN": asin,
        "Submission date": date,
        "Status": status,
        "Status description": description,
        "Offer name": f"Offer {asin}",
    }


def test_cleaned_uses_latest_row_per_asin():
    raw = pd.DataFrame(
        [
            _row("A1", "2026-01-01", "Evaluating"),
            _row("A1", "2026-02-01", "Accepted"),
            _row("A2", "2026-03-01", "Not accepted"),
        ]
    )
    cleaned = clean_btr_data(raw).set_index("ASIN")
    assert len(cleaned) == 2
    assert cleaned.loc["A1", "Status"] == "Accepted"


def test_rejected_active_offer_falls_back_to_latest_accepted_or_active():
    raw = pd.DataFrame(
        [
            _row("A1", "2026-01-01", "Accepted"),
            _row("A1", "2026-02-01", "Rejected: Active offer"),
            _row("A2", "2026-01-01", "Active"),
            _row("A2", "2026-02-01", "Rejected: Active offer"),
        ]
    )
    cleaned = clean_btr_data(raw).set_index("ASIN")
    assert cleaned.loc["A1", "Status"] == "Accepted"
    assert cleaned.loc["A2", "Status"] == "Active"


def test_status_changes_include_changed_and_new_but_not_removed_or_unchanged():
    previous = clean_btr_data(
        pd.DataFrame(
            [
                _row("CHANGED", "2026-01-01", "Evaluating", "Old"),
                _row("SAME", "2026-01-01", "Accepted", "Same"),
                _row("REMOVED", "2026-01-01", "Accepted", "Gone"),
            ]
        )
    )
    current = clean_btr_data(
        pd.DataFrame(
            [
                _row("CHANGED", "2026-02-01", "Accepted", "Approved"),
                _row("SAME", "2026-02-01", "Accepted", "Same"),
                _row("NEW", "2026-02-01", "Evaluating", "Under review"),
            ]
        )
    )
    changes = build_status_changes(previous, current).set_index("ASIN")
    assert set(changes.index) == {"CHANGED", "NEW"}
    assert changes.loc["CHANGED", "Previous Status"] == "Evaluating"
    assert changes.loc["CHANGED", "Status description"] == "Approved"
    assert changes.loc["NEW", "Change Type"] == "New ASIN"


def test_latest_files_are_selected_by_export_timestamp(tmp_path: Path):
    for name in (
        "2026-05-31T20_18_55.012Z.xlsx",
        "2026-07-22T20_18_55.012Z.xlsx",
        "2026-07-23T21_33_55.550Z.xlsx",
    ):
        (tmp_path / name).touch()
    previous, current = find_latest_two_files(tmp_path)
    assert previous.name.startswith("2026-07-22")
    assert current.name.startswith("2026-07-23")
