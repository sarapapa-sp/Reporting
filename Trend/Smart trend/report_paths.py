"""Shared naming convention for daily P&L reports."""
from datetime import date
from pathlib import Path


REPORT_DIR = Path(__file__).resolve().parent / "trend_params"
WINDOWS = {"Short": 5, "Long": 60, "Broad": 100, "Month": 7, "Quarter": 10}


def report_path(cob_date: date, window=None) -> Path:
    suffix = "_" + window.lower() if window else ""
    return REPORT_DIR / f"{cob_date:%Y}" / f"{cob_date:%m}" / f"trend_{cob_date.isoformat()}{suffix}.json"
