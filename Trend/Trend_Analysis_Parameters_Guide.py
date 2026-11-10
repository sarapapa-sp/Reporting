"""Trend metrics and historical P&L for monthly Report_data_Mon.xlsx files.

Layout: root_folder/YYYY/Report_data_Mon.xlsx
Install: pip install pandas numpy openpyxl
Use run_trend_analysis.py for command-line execution.
"""
from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


class ReportTrendAnalyzer:
    """Calculate report/indicator trend metrics for a selected data version."""

    EXCEPTION_COBS = frozenset({"2026-01-01"})
    TREND_PERIODS = ("n_business_days", "current_quarter", "current_year", "all")

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
        log_folder: str | Path | None = None,
        as_of_date: str | pd.Timestamp | None = None,
    ) -> None:
        if short_window < 2 or long_window < short_window:
            raise ValueError("Require 2 <= short_window <= long_window")
        self.root_folder = Path(root_folder)
        self.version = version
        self.short_window = short_window
        self.long_window = long_window
        self.anomaly_z_threshold = anomaly_z_threshold
        self.log_folder = Path(log_folder) if log_folder is not None else self.root_folder / "logs"
        self.logger = logging.getLogger(f"{__name__}.{id(self)}")
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False
        self.log_path: Path | None = None
        if as_of_date is not None:
            self.setup_logging(as_of_date)

    def setup_logging(self, cob_date: str | pd.Timestamp) -> Path:
        """Create one log file per analyzer instance / execution."""
        dt = pd.Timestamp(cob_date).normalize()
        if pd.isna(dt):
            raise ValueError("Invalid COB date")
        for handler in list(self.logger.handlers):
            handler.close()
            self.logger.removeHandler(handler)
        self.log_folder.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        self.log_path = self.log_folder / f"log_trend_{dt:%Y%m%d}_{timestamp}.log"
        handler = logging.FileHandler(self.log_path, encoding="utf-8")
        handler.setFormatter(logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
        ))
        self.logger.addHandler(handler)
        self.logger.info("Trend analyzer initialized | COB=%s | version=%s | root=%s", dt.date(), self.version, self.root_folder)
        self.logger.info("short_window=%s | long_window=%s | anomaly_z_threshold=%s", self.short_window, self.long_window, self.anomaly_z_threshold)
        return self.log_path

    def _ensure_logging(self, as_of_date: str | pd.Timestamp) -> None:
        if self.log_path is None:
            self.setup_logging(as_of_date)

    def close_logging(self) -> None:
        """Flush/close log handlers explicitly after the run."""
        for handler in list(self.logger.handlers):
            handler.flush()
            handler.close()
            self.logger.removeHandler(handler)

    def generate_filename(self, cob_date: str | pd.Timestamp) -> str:
        dt = pd.Timestamp(cob_date)
        return f"Report_data_{dt.strftime('%b')}.xlsx"

    def generate_file_path(self, cob_date: str | pd.Timestamp) -> Path:
        dt = pd.Timestamp(cob_date)
        return self.root_folder / str(dt.year) / self.generate_filename(dt)

    def file_exists(self, cob_date: str | pd.Timestamp) -> bool:
        return self.generate_file_path(cob_date).is_file()

    def get_history_paths(self, as_of_date: str | pd.Timestamp, months_to_load: int = 3) -> list[Path]:
        if months_to_load < 1:
            raise ValueError("months_to_load must be at least 1")
        end_month = pd.Timestamp(as_of_date).to_period("M")
        months = [end_month - offset for offset in range(months_to_load - 1, -1, -1)]
        paths = [self.generate_file_path(month.start_time) for month in months]
        self.logger.info("Requested history months=%s | workbooks=%s", months_to_load, [p.name for p in paths])
        return paths

    def load_daily_data(self, as_of_date: str | pd.Timestamp, months_to_load: int = 3) -> pd.DataFrame:
        as_of = pd.Timestamp(as_of_date).normalize()
        self._ensure_logging(as_of)
        self.logger.info("Loading history | requested COB=%s | months=%s | version=%s", as_of.date(), months_to_load, self.version)
        try:
            frames: list[pd.DataFrame] = []
            current_file = self.generate_file_path(as_of)
            if not current_file.is_file():
                raise FileNotFoundError(f"Current-month workbook not found: {current_file.name}")
            for path in self.get_history_paths(as_of, months_to_load):
                if not path.is_file():
                    self.logger.warning("Historical workbook missing; skipping: %s", path.name)
                    continue
                self.logger.info("Reading workbook: %s", path.name)
                data = pd.read_excel(
                    path, engine="openpyxl",
                    usecols=lambda name: str(name).strip().lower() in self.REQUIRED_COLUMNS,
                )
                data.columns = data.columns.astype(str).str.strip().str.lower()
                missing = set(self.REQUIRED_COLUMNS) - set(data.columns)
                if missing:
                    raise ValueError(f"Missing columns in {path.name}: {sorted(missing)}")
                data = data.loc[:, self.REQUIRED_COLUMNS].copy()
                selected = data["version"].astype("string").str.strip().str.casefold() == self.version.casefold()
                data = data.loc[selected.fillna(False)].copy()
                data["cob_date"] = pd.to_datetime(data["cob_date"], errors="coerce").dt.normalize()
                matched_dates = sorted(data["cob_date"].dropna().dt.strftime("%Y-%m-%d").unique().tolist())
                self.logger.info(
                    "Version filter | workbook=%s | version=%s | matched records=%d | unique reports=%d | unique COB dates=%d | COBs=%s",
                    path.name, self.version, len(data), data["report_name"].dropna().nunique(),
                    len(matched_dates), matched_dates,
                )
                if data.empty:
                    continue
                for col in ("daily", "mtd", "ytd", "bus_day", "month"):
                    data[col] = pd.to_numeric(data[col], errors="coerce")
                for col in ("report_name", "indicator"):
                    data[col] = data[col].astype("string").str.strip()
                invalid = data[list(self.REQUIRED_COLUMNS)].isna().any(axis=1)
                invalid |= data["report_name"].eq("") | data["indicator"].eq("")
                if invalid.any():
                    raise ValueError(f"Found {int(invalid.sum())} invalid {self.version} rows in {path.name}")
                wrong_month = data["month"].ne(data["cob_date"].dt.month)
                if wrong_month.any():
                    raise ValueError(f"Month field does not match cob_date in {path.name}")
                data = data.loc[data["cob_date"] <= as_of]
                self.logger.info("Usable rows %s: %d | COBs: %s", path.name, len(data), sorted(data["cob_date"].dt.strftime("%Y-%m-%d").unique().tolist()))
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
            combined = combined.sort_values(["report_name", "indicator", "cob_date"]).reset_index(drop=True)
            self.logger.info("History loaded | records=%d | COBs=%s | combinations=%d", len(combined), sorted(combined["cob_date"].dt.strftime("%Y-%m-%d").unique().tolist()), combined.groupby(["report_name", "indicator"]).ngroups)
            self.validate_cob_continuity(combined, as_of, months_to_load)
            return combined
        except Exception:
            self.logger.exception("Failed loading history for COB=%s", as_of.date())
            raise

    def validate_cob_continuity(
        self,
        history: pd.DataFrame,
        as_of_date: str | pd.Timestamp,
        months_to_load: int | None = None,
    ) -> None:
        """Raise if any expected weekday COB is absent, except EXCEPTION_COBS.

        Checks all dates in the requested month range (when supplied) and also
        checks each report/indicator from its first observed date through as_of.
        Uses Monday-Friday as the business calendar; add holidays to EXCEPTION_COBS.
        """
        as_of = pd.Timestamp(as_of_date).normalize()
        self._ensure_logging(as_of)
        exceptions = {pd.Timestamp(day).normalize() for day in self.EXCEPTION_COBS}
        if months_to_load is not None:
            if months_to_load < 1:
                raise ValueError("months_to_load must be at least 1")
            start = (as_of.to_period("M") - (months_to_load - 1)).start_time
        else:
            if history.empty:
                raise ValueError("Cannot validate COB continuity: empty history")
            start = pd.to_datetime(history["cob_date"]).min().normalize()
        expected = pd.DatetimeIndex(pd.bdate_range(start, as_of)).normalize()
        observed = set(pd.to_datetime(history["cob_date"]).dt.normalize())
        missing = [day for day in expected if day not in observed]
        exempted = [day for day in missing if day in exceptions]
        unexpected = [day for day in missing if day not in exceptions]
        fmt = lambda days: [day.strftime("%Y-%m-%d") for day in days]
        self.logger.info(
            "COB continuity | start=%s | end=%s | expected weekdays=%d | observed COBs=%s",
            start.date(), as_of.date(), len(expected), fmt(sorted(observed)),
        )
        self.logger.info("Exception COBs configured: %s", sorted(self.EXCEPTION_COBS))
        if exempted:
            self.logger.warning("Missing COBs allowed by exception: %s", fmt(exempted))
        if unexpected:
            self.logger.error("Missing required COBs: %s", fmt(unexpected))
            raise ValueError(f"Missing required COB dates: {', '.join(fmt(unexpected))}")

        # A date present for one report must not mask a gap in another report.
        # Permit report/indicator combinations to start later, but require
        # continuity from their first appearance to the requested COB.
        for (report, indicator), group in history.groupby(["report_name", "indicator"], sort=True):
            group_dates = set(pd.to_datetime(group["cob_date"]).dt.normalize())
            first_date = min(group_dates)
            group_expected = pd.bdate_range(first_date, as_of)
            group_missing = [day for day in group_expected if day not in group_dates]
            group_exempted = [day for day in group_missing if day in exceptions]
            group_unexpected = [day for day in group_missing if day not in exceptions]
            if group_exempted:
                self.logger.warning(
                    "Report/indicator=%s/%s | allowed exception COBs=%s",
                    report, indicator, fmt(group_exempted),
                )
            if group_unexpected:
                self.logger.error(
                    "Report/indicator=%s/%s | missing required COBs=%s",
                    report, indicator, fmt(group_unexpected),
                )
                raise ValueError(
                    f"Missing required COB dates for report={report}, indicator={indicator}: "
                    + ", ".join(fmt(group_unexpected))
                )
        self.logger.info("COB continuity validation passed | requested COB=%s", as_of.date())

    def select_previous_month_cob(
        self,
        current_cob: str | pd.Timestamp,
    ) -> pd.Timestamp:
        """Map a current COB to the prior month by centered business-day alignment.

        Uses full Monday-Friday calendars excluding EXCEPTION_COBS, not just dates
        already loaded this month. When current has more days, repeats the middle
        previous-month COB; when previous has more, skips middle previous-month
        COB(s). A stable mapping is therefore available even mid-month.
        """
        current = pd.Timestamp(current_cob).normalize()
        self._ensure_logging(current)
        if current.weekday() >= 5 or current.strftime("%Y-%m-%d") in self.EXCEPTION_COBS:
            raise ValueError(f"Not an expected business COB: {current.date()}")
        previous_period = current.to_period("M") - 1
        exceptions = {pd.Timestamp(x).normalize() for x in self.EXCEPTION_COBS}

        def monthly_days(period: pd.Period) -> list[pd.Timestamp]:
            return [pd.Timestamp(d).normalize() for d in pd.bdate_range(period.start_time, period.end_time)
                    if pd.Timestamp(d).normalize() not in exceptions]

        current_days = monthly_days(current.to_period("M"))
        previous_days = monthly_days(previous_period)
        current_pos = current_days.index(current)
        current_count, previous_count = len(current_days), len(previous_days)
        if current_count >= previous_count:
            # Insert repeated middle positions into the longer current calendar.
            # Example: 22 current vs 21 previous: positions 11 and 12 (1-based)
            # both match previous position 11.
            extras = current_count - previous_count
            insert_at = (previous_count + 1) // 2
            prior_pos = current_pos if current_pos < insert_at else max(insert_at, current_pos - extras)
        else:
            # Drop surplus COBs from the middle of the longer previous calendar.
            extras = previous_count - current_count
            skip_at = (previous_count - extras) // 2
            prior_pos = current_pos if current_pos < skip_at else current_pos + extras
        prior_pos = max(0, min(prior_pos, previous_count - 1))
        selected = previous_days[prior_pos]
        adjustment = (
            "repeat middle previous-month COB" if current_count > previous_count else
            "skip middle previous-month COB(s)" if current_count < previous_count else
            "one-to-one"
        )
        self.logger.info(
            "Previous-month COB selection | current=%s (position=%d/%d) | previous=%s "
            "(position=%d/%d) | adjustment=%s",
            current.strftime("%Y-%m-%d"), current_pos + 1, current_count,
            selected.strftime("%Y-%m-%d"), prior_pos + 1, previous_count, adjustment,
        )
        self.logger.info(
            "COB alignment calendars | current month=%s | current COBs=%s | previous month=%s | previous COBs=%s",
            current.to_period("M"), [x.strftime("%Y-%m-%d") for x in current_days],
            previous_period, [x.strftime("%Y-%m-%d") for x in previous_days],
        )
        return selected

    @staticmethod
    def _difference(current: float, previous: float | None) -> float | None:
        return None if previous is None else float(current - previous)

    @staticmethod
    def _percent_change(current: float, previous: float | None) -> float | None:
        if previous is None or previous == 0:
            return None
        return float(100 * (current - previous) / abs(previous))

    def calculate_trend_parameters(self, daily_data: pd.DataFrame, as_of_date: str | pd.Timestamp) -> pd.DataFrame:
        as_of = pd.Timestamp(as_of_date).normalize()
        self._ensure_logging(as_of)
        self.logger.info("Trend calculation started | requested COB=%s | source rows=%d", as_of.date(), len(daily_data))
        try:
            self.validate_cob_continuity(daily_data, as_of)

            rows: list[dict[str, Any]] = []
            for (report, indicator), group in daily_data.groupby(["report_name", "indicator"], sort=True):
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
                latest_dt = pd.Timestamp(latest["cob_date"])
                earlier = group.loc[dates < latest_dt]
                previous_mtd_rows = earlier.loc[
                    (earlier["cob_date"].dt.year == latest_dt.year)
                    & (earlier["cob_date"].dt.month == latest_dt.month)
                ]
                previous_ytd_rows = earlier.loc[earlier["cob_date"].dt.year == latest_dt.year]
                prev_mtd = float(previous_mtd_rows.iloc[-1]["mtd"]) if not previous_mtd_rows.empty else None
                prev_ytd = float(previous_ytd_rows.iloc[-1]["ytd"]) if not previous_ytd_rows.empty else None
                mtd = float(latest["mtd"])
                ytd = float(latest["ytd"])
                latest_bus_day = latest["bus_day"]
                selected_previous_cob = self.select_previous_month_cob(latest_dt)
                prior_month_match = earlier.loc[
                    earlier["cob_date"] == selected_previous_cob
                ]
                if prior_month_match.empty:
                    self.logger.warning(
                        "Selected previous-month COB is unavailable for report=%s | indicator=%s | COB=%s",
                        report, indicator, selected_previous_cob.strftime("%Y-%m-%d"),
                    )
                last_year_match = earlier.loc[
                    (earlier["cob_date"].dt.year == latest_dt.year - 1)
                    & (earlier["cob_date"].dt.month == latest_dt.month)
                    & (earlier["bus_day"] == latest_bus_day)
                ]
                prior_month_daily = float(prior_month_match.iloc[-1]["daily"]) if not prior_month_match.empty else None
                prior_month_mtd = float(prior_month_match.iloc[-1]["mtd"]) if not prior_month_match.empty else None
                prior_year_ytd = float(last_year_match.iloc[-1]["ytd"]) if not last_year_match.empty else None
                self.logger.info("Trend: report=%s | indicator=%s | latest COB=%s | observations=%d | direction=%s | anomaly=%s", report, indicator, latest_dt.date(), n, direction, anomaly)
                rows.append({
                    "report_name": str(report), "indicator": str(indicator), "version": self.version,
                    "latest_cob_date": latest_dt.strftime("%Y-%m-%d"),
                    "days_since_latest": int((as_of - latest_dt).days),
                    "bus_day": int(latest["bus_day"]), "observations_loaded": n,
                    "daily_current": current, "daily_previous": previous,
                    "daily_change": self._difference(current, previous),
                    "daily_change_pct": self._percent_change(current, previous),
                    "daily_avg_short": short_avg, "daily_avg_long": long_avg,
                    "daily_std_short": float(short.std(ddof=1)) if len(short) >= 2 else None,
                    "daily_min_short": float(short.min()), "daily_max_short": float(short.max()),
                    "daily_change_N_business_days": (
                        float(current - values[-(self.short_window + 1)])
                        if n > self.short_window else None
                    ),
                    "daily_slope_short": slope, "daily_trend_direction": direction,
                    "daily_anomaly_zscore": zscore, "daily_is_anomaly": anomaly,
                    "previous_month_selected_cob": selected_previous_cob.strftime("%Y-%m-%d"),
                    "daily_previous_month_same_bus_day": prior_month_daily,
                    "daily_mom_change": self._difference(current, prior_month_daily),
                    "daily_mom_change_pct": self._percent_change(current, prior_month_daily),
                    "mtd_current": mtd, "mtd_previous_same_month": prev_mtd,
                    "mtd_change": self._difference(mtd, prev_mtd),
                    "mtd_change_pct": self._percent_change(mtd, prev_mtd),
                    "mtd_previous_month_same_bus_day": prior_month_mtd,
                    "mtd_mom_change": self._difference(mtd, prior_month_mtd),
                    "mtd_mom_change_pct": self._percent_change(mtd, prior_month_mtd),
                    "ytd_current": ytd, "ytd_previous_same_year": prev_ytd,
                    "ytd_change": self._difference(ytd, prev_ytd),
                    "ytd_change_pct": self._percent_change(ytd, prev_ytd),
                    "ytd_previous_year_same_bus_day": prior_year_ytd,
                    "ytd_yoy_change": self._difference(ytd, prior_year_ytd),
                    "ytd_yoy_change_pct": self._percent_change(ytd, prior_year_ytd),
                })
            result = pd.DataFrame(rows)
            self.logger.info("Trend calculation complete | combinations=%d", len(result))
            return result
        except Exception:
            self.logger.exception("Trend calculation failed")
            raise

    def generate_trend_parameters(self, as_of_date: str | pd.Timestamp, months_to_load: int = 3) -> pd.DataFrame:
        history = self.load_daily_data(as_of_date, months_to_load)
        return self.calculate_trend_parameters(history, as_of_date)

    def parameters_for_llm(self, as_of_date: str | pd.Timestamp, months_to_load: int = 3) -> list[dict[str, Any]]:
        results = self.generate_trend_parameters(as_of_date, months_to_load)
        return json.loads(results.to_json(orient="records"))

    def _historical_pnl(
        self, as_of_date: str | pd.Timestamp, n: int, metric: str,
        months_to_load: int = 3, history: pd.DataFrame | None = None,
    ) -> pd.DataFrame:
        """Wide DataFrame: report_name, indicator, then n available COB dates (oldest to newest).

        Missing values for a report/indicator on a selected COB remain NaN.
        n counts the requested/latest available COB and n-1 earlier COBs.
        """
        if not isinstance(n, int) or isinstance(n, bool) or n < 1:
            raise ValueError("n must be a positive integer")
        if metric not in ("daily", "mtd", "ytd"):
            raise ValueError("metric must be daily, mtd, or ytd")
        as_of = pd.Timestamp(as_of_date).normalize()
        self._ensure_logging(as_of)
        try:
            if history is None:
                history = self.load_daily_data(as_of, months_to_load)
            data = history.loc[history["cob_date"] <= as_of].copy()
            self.validate_cob_continuity(data, as_of)
            dates = sorted(data["cob_date"].dropna().unique())[-n:]
            date_labels = [pd.Timestamp(d).strftime("%Y-%m-%d") for d in dates]
            self.logger.info("%s P&L | requested COB=%s | requested n=%s | selected COBs=%s", metric.upper(), as_of.date(), n, date_labels)
            if len(dates) < n:
                self.logger.warning("Only %d distinct COBs available of %d requested for %s", len(dates), n, metric)
            if not dates:
                raise ValueError(f"No COB dates available on/before {as_of.date()}")
            selected = data.loc[data["cob_date"].isin(dates), ["report_name", "indicator", "cob_date", metric]].copy()
            selected["cob_date"] = selected["cob_date"].dt.strftime("%Y-%m-%d")
            result = selected.pivot(index=["report_name", "indicator"], columns="cob_date", values=metric)
            result = result.reindex(columns=date_labels).reset_index()
            result.columns.name = None
            self.logger.info("%s P&L result | rows=%d | date columns=%d", metric.upper(), len(result), len(date_labels))
            return result
        except Exception:
            self.logger.exception("Failed generating %s P&L history", metric.upper())
            raise

    def get_previous_daily_pnl(self, as_of_date: str | pd.Timestamp, n: int, months_to_load: int = 3, history: pd.DataFrame | None = None) -> pd.DataFrame:
        return self._historical_pnl(as_of_date, n, "daily", months_to_load, history)

    def get_previous_mtd_pnl(self, as_of_date: str | pd.Timestamp, n: int, months_to_load: int = 3, history: pd.DataFrame | None = None) -> pd.DataFrame:
        return self._historical_pnl(as_of_date, n, "mtd", months_to_load, history)

    def get_previous_ytd_pnl(self, as_of_date: str | pd.Timestamp, n: int, months_to_load: int = 3, history: pd.DataFrame | None = None) -> pd.DataFrame:
        return self._historical_pnl(as_of_date, n, "ytd", months_to_load, history)


    def period_start(self, as_of_date: str | pd.Timestamp, trend_period: str) -> pd.Timestamp:
        """Return calendar start of a quarter or year (the COB is inclusive)."""
        as_of = pd.Timestamp(as_of_date).normalize()
        if trend_period == "current_quarter":
            return as_of.to_period("Q").start_time.normalize()
        if trend_period == "current_year":
            return as_of.to_period("Y").start_time.normalize()
        raise ValueError("trend_period must be current_quarter or current_year")

    def required_months_to_load(
        self, as_of_date: str | pd.Timestamp, trend_period: str,
        months_to_load: int = 3,
    ) -> int:
        """Expand the workbook lookback to cover the full selected calendar period."""
        if months_to_load < 1:
            raise ValueError("months_to_load must be at least 1")
        if trend_period not in self.TREND_PERIODS:
            raise ValueError(f"trend_period must be one of {self.TREND_PERIODS}")
        if trend_period == "n_business_days":
            return months_to_load
        start = self.period_start(as_of_date, "current_year" if trend_period == "all" else trend_period)
        as_of = pd.Timestamp(as_of_date).normalize()
        months_needed = (as_of.year - start.year) * 12 + as_of.month - start.month + 1
        # Calendar-period modes intentionally load from their period start;
        # loading earlier months would incorrectly require unrelated COBs.
        return months_needed

    def calculate_period_trend_parameters(
        self,
        history: pd.DataFrame,
        as_of_date: str | pd.Timestamp,
        trend_period: str = "current_quarter",
        n: int = 5,
    ) -> pd.DataFrame:
        """Return a report/indicator trend for N COBs, QTD, or YTD.

        Daily P&L is used for all aggregate statistics and regression.
        The separate end-of-period MTD/YTD snapshots are not summed.
        The latest COB is included, and exceptional COBs are excluded only
        from the expected calendar (not from actual observed records).
        """
        if trend_period not in ("n_business_days", "current_quarter", "current_year"):
            raise ValueError("Invalid period: use n_business_days, current_quarter or current_year")
        if not isinstance(n, int) or isinstance(n, bool) or n < 1:
            raise ValueError("n must be a positive integer")
        as_of = pd.Timestamp(as_of_date).normalize()
        self._ensure_logging(as_of)
        data = history.loc[history["cob_date"] <= as_of].copy()
        if data.empty:
            raise ValueError("No history available for trend calculation")
        # Also catch report/indicator gaps in the original loaded history.
        self.validate_cob_continuity(data, as_of)
        if trend_period == "n_business_days":
            dates = sorted(pd.to_datetime(data["cob_date"]).unique())[-n:]
            if len(dates) < n:
                self.logger.warning("N-business-day trend requested %d COBs but only %d available", n, len(dates))
            data = data.loc[data["cob_date"].isin(dates)].copy()
            start = pd.Timestamp(dates[0]).normalize()
        else:
            start = self.period_start(as_of, trend_period)
            data = data.loc[data["cob_date"] >= start].copy()
            # A selected period is only valid when the entire period exists;
            # the existing per-group validator allows late-introduced reports.
            self.validate_cob_continuity(data, as_of, months_to_load=None)
            expected = set(pd.bdate_range(start, as_of)) - {
                pd.Timestamp(x).normalize() for x in self.EXCEPTION_COBS
            }
            observed = set(pd.to_datetime(data["cob_date"]).dt.normalize())
            missing = sorted(expected - observed)
            if missing:
                labels = [x.strftime("%Y-%m-%d") for x in missing]
                self.logger.error("%s period missing COBs: %s", trend_period, labels)
                raise ValueError(f"Missing required {trend_period} COB dates: {labels}")
        dates_used = sorted(data["cob_date"].dt.strftime("%Y-%m-%d").unique().tolist())
        self.logger.info(
            "Period trend selection | period=%s | n=%s | start=%s | end=%s | COB count=%s | selected COBs=%s",
            trend_period, n, start.date(), as_of.date(), len(dates_used), dates_used,
        )
        if data.empty:
            raise ValueError(f"No available records for trend_period={trend_period}")
        result_rows = []
        for (report, indicator), group in data.groupby(["report_name", "indicator"], sort=True):
            group = group.sort_values("cob_date")
            vals = group["daily"].to_numpy(dtype=float)
            count = len(vals)
            first, last = float(vals[0]), float(vals[-1])
            slope = float(np.polyfit(np.arange(count), vals, 1)[0]) if count >= 3 else None
            direction = (
                "Insufficient history" if slope is None else
                "Rising" if slope > 1e-10 else
                "Falling" if slope < -1e-10 else "Flat"
            )
            std = float(vals.std(ddof=1)) if count >= 2 else None
            # Z-score compares latest day to earlier daily observations in this period.
            prior = vals[:-1]
            zscore = None
            is_anomaly = None
            if len(prior) >= 5:
                prior_std = float(prior.std(ddof=1))
                prior_mean = float(prior.mean())
                if prior_std > 0:
                    zscore = float((last - prior_mean) / prior_std)
                    is_anomaly = bool(abs(zscore) >= self.anomaly_z_threshold)
                else:
                    is_anomaly = bool(last != prior_mean)
            period_dates = group["cob_date"].dt.strftime("%Y-%m-%d").tolist()
            result_rows.append({
                "report_name": str(report), "indicator": str(indicator), "version": self.version,
                "trend_period": trend_period, "period_requested_start": start.strftime("%Y-%m-%d"),
                "period_first_cob": period_dates[0], "period_last_cob": period_dates[-1],
                "period_observations": count, "period_cobs": ", ".join(period_dates),
                "daily_first": first, "daily_latest": last,
                "daily_period_change": self._difference(last, first),
                "daily_period_change_pct": self._percent_change(last, first),
                "daily_period_sum": float(vals.sum()),
                "daily_period_average": float(vals.mean()),
                "daily_period_std": std, "daily_period_min": float(vals.min()),
                "daily_period_max": float(vals.max()),
                "daily_period_slope": slope, "daily_period_direction": direction,
                "daily_period_anomaly_zscore": zscore,
                "daily_period_is_anomaly": is_anomaly,
                "mtd_first": float(group.iloc[0]["mtd"]),
                "mtd_latest": float(group.iloc[-1]["mtd"]),
                "ytd_first": float(group.iloc[0]["ytd"]),
                "ytd_latest": float(group.iloc[-1]["ytd"]),
            })
            self.logger.info("Period trend | period=%s | report=%s | indicator=%s | COBs=%d | direction=%s",
                             trend_period, report, indicator, count, direction)
        results = pd.DataFrame(result_rows)
        self.logger.info("Period trends complete | period=%s | combinations=%d", trend_period, len(results))
        return results

    def get_period_pnl(
        self, as_of_date: str | pd.Timestamp, trend_period: str, metric: str = "daily",
        months_to_load: int = 3, history: pd.DataFrame | None = None,
    ) -> pd.DataFrame:
        """Wide daily/MTD/YTD history from the quarter or year start through COB."""
        if trend_period not in ("current_quarter", "current_year"):
            raise ValueError("trend_period must be current_quarter or current_year")
        if metric not in ("daily", "mtd", "ytd"):
            raise ValueError("metric must be daily, mtd, or ytd")
        as_of = pd.Timestamp(as_of_date).normalize()
        self._ensure_logging(as_of)
        if history is None:
            lookback = self.required_months_to_load(as_of, trend_period, months_to_load)
            history = self.load_daily_data(as_of, lookback)
        # Reuse the strict period validity check before building the pivot.
        self.calculate_period_trend_parameters(history, as_of, trend_period)
        start = self.period_start(as_of, trend_period)
        selected = history.loc[history["cob_date"].between(start, as_of),
                               ["report_name", "indicator", "cob_date", metric]].copy()
        date_columns = sorted(selected["cob_date"].dt.strftime("%Y-%m-%d").unique().tolist())
        selected["cob_date"] = selected["cob_date"].dt.strftime("%Y-%m-%d")
        output = selected.pivot(index=["report_name", "indicator"], columns="cob_date", values=metric)
        output = output.reindex(columns=date_columns).reset_index()
        output.columns.name = None
        self.logger.info("Period P&L pivot | period=%s | metric=%s | selected COBs=%s | rows=%d",
                         trend_period, metric, date_columns, len(output))
        return output
