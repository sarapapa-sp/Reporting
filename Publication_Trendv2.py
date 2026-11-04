"""Trend metrics for monthly Report_data_Mon.xlsx workbooks.

Folder example: /data/reports/2026/Report_data_Nov.xlsx
Install: pip install pandas numpy openpyxl
Run: python report_trend_analyzer.py /data/reports --as-of 2026-11-04
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


class ReportTrendAnalyzer:
    """Calculate per-(report_name, indicator) metrics for version='Daily'."""

    REQUIRED_COLUMNS = (
        "cob_date", "bus_day", "month", "version", "report_name",
        "indicator", "daily", "mtd", "ytd",
    )

    def __init__(
        self,
        root_folder: str | Path,
        version: str = "Daily",
        short_window: int = 5,
        long_window: int = 20,
        anomaly_z_threshold: float = 2.0,
    ) -> None:
        if short_window < 2 or long_window < short_window:
            raise ValueError("Require 2 <= short_window <= long_window")
        self.root_folder = Path(root_folder)
        self.version = version
        self.short_window = short_window
        self.long_window = long_window
        self.anomaly_z_threshold = anomaly_z_threshold

    # 1) Filename and path generation
    def generate_filename(self, cob_date: str | pd.Timestamp) -> str:
        dt = pd.Timestamp(cob_date)
        return f"Report_data_{dt.strftime('%b')}.xlsx"

    def generate_file_path(self, cob_date: str | pd.Timestamp) -> Path:
        dt = pd.Timestamp(cob_date)
        return self.root_folder / str(dt.year) / self.generate_filename(dt)

    # 2) Existence check
    def file_exists(self, cob_date: str | pd.Timestamp) -> bool:
        return self.generate_file_path(cob_date).is_file()

    # 3) Identify history files, including across calendar-year boundaries
    def get_history_paths(
        self, as_of_date: str | pd.Timestamp, months_to_load: int = 3
    ) -> list[Path]:
        if months_to_load < 1:
            raise ValueError("months_to_load must be at least 1")
        end_month = pd.Timestamp(as_of_date).to_period("M")
        months = [end_month - offset for offset in range(months_to_load - 1, -1, -1)]
        return [self.generate_file_path(month.start_time) for month in months]

    # 4) Read only required columns and Daily-version rows
    def load_daily_data(
        self, as_of_date: str | pd.Timestamp, months_to_load: int = 3
    ) -> pd.DataFrame:
        as_of = pd.Timestamp(as_of_date).normalize()
        frames: list[pd.DataFrame] = []
        current_file = self.generate_file_path(as_of)

        if not current_file.is_file():
            raise FileNotFoundError(f"Current-month workbook not found: {current_file}")

        for path in self.get_history_paths(as_of, months_to_load):
            if not path.is_file():
                # Older history may be missing; current month is required.
                continue

            data = pd.read_excel(
                path,
                engine="openpyxl",
                usecols=lambda name: str(name).strip().lower() in self.REQUIRED_COLUMNS,
            )
            data.columns = data.columns.astype(str).str.strip().str.lower()
            missing = set(self.REQUIRED_COLUMNS) - set(data.columns)
            if missing:
                raise ValueError(f"Missing columns in {path}: {sorted(missing)}")

            data = data.loc[:, self.REQUIRED_COLUMNS].copy()
            selected = data["version"].astype("string").str.strip().str.casefold() == self.version.casefold()
            data = data.loc[selected.fillna(False)].copy()
            if data.empty:
                continue

            data["cob_date"] = pd.to_datetime(data["cob_date"], errors="coerce").dt.normalize()
            for col in ("daily", "mtd", "ytd", "bus_day", "month"):
                data[col] = pd.to_numeric(data[col], errors="coerce")
            for col in ("report_name", "indicator"):
                data[col] = data[col].astype("string").str.strip()

            # Drop rows after the selected analysis date, but reject malformed rows.
            invalid = data[list(self.REQUIRED_COLUMNS)].isna().any(axis=1)
            invalid |= data["report_name"].eq("") | data["indicator"].eq("")
            if invalid.any():
                raise ValueError(f"Found {int(invalid.sum())} invalid Daily rows in {path}")
            wrong_month = data["month"].ne(data["cob_date"].dt.month)
            if wrong_month.any():
                raise ValueError(f"Month field does not match cob_date in {path}")
            data = data.loc[data["cob_date"] <= as_of]
            if not data.empty:
                frames.append(data)

        if not frames:
            raise ValueError(f"No '{self.version}' data on/before {as_of.date()} in readable files")

        combined = pd.concat(frames, ignore_index=True)
        keys = ["cob_date", "report_name", "indicator"]
        duplicates = combined.duplicated(keys, keep=False)
        if duplicates.any():
            sample = combined.loc[duplicates, keys].head(5).to_dict("records")
            raise ValueError(f"Duplicate Daily rows for (date, report, indicator): {sample}")
        return combined.sort_values(["report_name", "indicator", "cob_date"]).reset_index(drop=True)

    @staticmethod
    def _difference(current: float, previous: float | None) -> float | None:
        return None if previous is None else float(current - previous)

    @staticmethod
    def _percent_change(current: float, previous: float | None) -> float | None:
        if previous is None or previous == 0:
            return None  # Undefined when baseline is zero
        return float(100 * (current - previous) / abs(previous))

    # 5) Calculate a metrics record for every distinct report and indicator
    def calculate_trend_parameters(
        self, daily_data: pd.DataFrame, as_of_date: str | pd.Timestamp
    ) -> pd.DataFrame:
        as_of = pd.Timestamp(as_of_date).normalize()
        rows: list[dict[str, Any]] = []
        for (report, indicator), group in daily_data.groupby(
            ["report_name", "indicator"], sort=True
        ):
            group = group.sort_values("cob_date").reset_index(drop=True)
            latest = group.iloc[-1]
            dates = group["cob_date"]
            values = group["daily"].to_numpy(dtype=float)
            n = len(values)
            current = float(values[-1])
            previous = float(values[-2]) if n >= 2 else None

            short = values[-self.short_window:]
            long = values[-self.long_window:]
            short_avg = float(short.mean())
            long_avg = float(long.mean())
            slope = float(np.polyfit(np.arange(len(short)), short, 1)[0]) if len(short) >= 3 else None
            direction = (
                "Insufficient history" if slope is None else
                "Rising" if slope > 1e-10 else
                "Falling" if slope < -1e-10 else "Flat"
            )

            # Anomaly baseline deliberately EXCLUDES today's point.
            historical = values[:-1][-self.long_window:]
            zscore: float | None = None
            anomaly: bool | None = None
            if len(historical) >= 5:
                mean = float(historical.mean())
                std = float(historical.std(ddof=1))
                if std > 0:
                    zscore = float((current - mean) / std)
                    anomaly = bool(abs(zscore) >= self.anomaly_z_threshold)
                else:
                    anomaly = bool(current != mean)

            # MTD can only be compared within the SAME calendar month.
            # YTD can only be compared within the SAME calendar year.
            latest_dt = pd.Timestamp(latest["cob_date"])
            earlier = group.loc[dates < latest_dt]
            previous_mtd_rows = earlier.loc[
                (earlier["cob_date"].dt.year == latest_dt.year)
                & (earlier["cob_date"].dt.month == latest_dt.month)
            ]
            previous_ytd_rows = earlier.loc[
                earlier["cob_date"].dt.year == latest_dt.year
            ]
            prev_mtd = float(previous_mtd_rows.iloc[-1]["mtd"]) if not previous_mtd_rows.empty else None
            prev_ytd = float(previous_ytd_rows.iloc[-1]["ytd"]) if not previous_ytd_rows.empty else None
            mtd = float(latest["mtd"])
            ytd = float(latest["ytd"])

            # Like-for-like comparisons using the SAME business-day number.
            # No substitute is made if the exact matching bus_day is missing.
            latest_bus_day = latest["bus_day"]
            previous_month = latest_dt.to_period("M") - 1
            prior_month_match = earlier.loc[
                (earlier["cob_date"].dt.to_period("M") == previous_month)
                & (earlier["bus_day"] == latest_bus_day)
            ]
            last_year_match = earlier.loc[
                (earlier["cob_date"].dt.year == latest_dt.year - 1)
                & (earlier["cob_date"].dt.month == latest_dt.month)
                & (earlier["bus_day"] == latest_bus_day)
            ]
            prior_month_daily = (
                float(prior_month_match.iloc[-1]["daily"])
                if not prior_month_match.empty else None
            )
            prior_month_mtd = (
                float(prior_month_match.iloc[-1]["mtd"])
                if not prior_month_match.empty else None
            )
            prior_year_ytd = (
                float(last_year_match.iloc[-1]["ytd"])
                if not last_year_match.empty else None
            )

            rows.append({
                "report_name": str(report),
                "indicator": str(indicator),
                "version": self.version,
                "latest_cob_date": latest_dt.strftime("%Y-%m-%d"),
                "days_since_latest": int((as_of - latest_dt).days),
                "bus_day": int(latest["bus_day"]),
                "observations_loaded": n,
                "daily_current": current,
                "daily_previous": previous,
                "daily_change": self._difference(current, previous),
                "daily_change_pct": self._percent_change(current, previous),
                "daily_avg_short": short_avg,
                "daily_avg_long": long_avg,
                "daily_std_short": float(short.std(ddof=1)) if len(short) >= 2 else None,
                "daily_min_short": float(short.min()),
                "daily_max_short": float(short.max()),
                "daily_change_N_business_days": (
                    float(current - values[-(self.short_window + 1)])
                    if n > self.short_window else None
                ),
                "daily_slope_short": slope,
                "daily_trend_direction": direction,
                "daily_anomaly_zscore": zscore,
                "daily_is_anomaly": anomaly,
                "daily_previous_month_same_bus_day": prior_month_daily,
                "daily_mom_change": self._difference(current, prior_month_daily),
                "daily_mom_change_pct": self._percent_change(current, prior_month_daily),
                "mtd_current": mtd,
                "mtd_previous_same_month": prev_mtd,
                "mtd_change": self._difference(mtd, prev_mtd),
                "mtd_change_pct": self._percent_change(mtd, prev_mtd),
                "mtd_previous_month_same_bus_day": prior_month_mtd,
                "mtd_mom_change": self._difference(mtd, prior_month_mtd),
                "mtd_mom_change_pct": self._percent_change(mtd, prior_month_mtd),
                "ytd_current": ytd,
                "ytd_previous_same_year": prev_ytd,
                "ytd_change": self._difference(ytd, prev_ytd),
                "ytd_change_pct": self._percent_change(ytd, prev_ytd),
                "ytd_previous_year_same_bus_day": prior_year_ytd,
                "ytd_yoy_change": self._difference(ytd, prior_year_ytd),
                "ytd_yoy_change_pct": self._percent_change(ytd, prior_year_ytd),
            })
        return pd.DataFrame(rows)

    # 6) Public entry point, ready to call from a UI or an AI orchestrator
    def generate_trend_parameters(
        self, as_of_date: str | pd.Timestamp, months_to_load: int = 3
    ) -> pd.DataFrame:
        history = self.load_daily_data(as_of_date, months_to_load)
        return self.calculate_trend_parameters(history, as_of_date)

    def parameters_for_llm(
        self, as_of_date: str | pd.Timestamp, months_to_load: int = 3
    ) -> list[dict[str, Any]]:
        results = self.generate_trend_parameters(as_of_date, months_to_load)
        # Convert numpy/pandas values to regular JSON-compatible Python values.
        return json.loads(results.to_json(orient="records"))


def main() -> None:
    parser = argparse.ArgumentParser(description="Compute trend metrics by report and indicator")
    parser.add_argument("root_folder", help="Parent directory containing year subfolders")
    parser.add_argument("--as-of", required=True, help="COB date YYYY-MM-DD")
    parser.add_argument("--months", type=int, default=3, help="Number of monthly files to search")
    parser.add_argument("--out", help="Optional output CSV path")
    args = parser.parse_args()

    analyzer = ReportTrendAnalyzer(args.root_folder)
    trends = analyzer.generate_trend_parameters(args.as_of, args.months)
    if args.out:
        trends.to_csv(args.out, index=False)
        print(f"Saved {len(trends)} report/indicator combinations to {args.out}")
    else:
        print(trends.to_string(index=False))


if __name__ == "__main__":
    main()

'''
Parameter	Description
daily_current	Latest daily value
daily_previous	Previous available daily value
daily_change	Current daily − previous daily
daily_change_pct	Percentage change from previous daily
daily_avg_short	Average of latest 5 observations
daily_avg_long	Average of latest 20 observations
daily_std_short	Standard deviation of latest 5 observations
daily_min_short	Minimum value in latest 5 observations
daily_max_short	Maximum value in latest 5 observations
daily_change_N_business_days	Difference from 5 observations earlier
daily_slope_short	Linear regression slope over latest 5 observations
daily_trend_direction	Rising, Falling, Flat, or Insufficient history
daily_anomaly_zscore	How unusual the latest value is relative to previous observations
daily_is_anomaly	Whether the latest value breaches the anomaly threshold


'''


