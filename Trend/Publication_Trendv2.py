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
        self.logger.info("Requested history months=%s | workbooks=%s", months_to_load, [str(p) for p in paths])
        return paths

    def load_daily_data(self, as_of_date: str | pd.Timestamp, months_to_load: int = 3) -> pd.DataFrame:
        as_of = pd.Timestamp(as_of_date).normalize()
        self._ensure_logging(as_of)
        self.logger.info("Loading history | requested COB=%s | months=%s | version=%s", as_of.date(), months_to_load, self.version)
        try:
            frames: list[pd.DataFrame] = []
            current_file = self.generate_file_path(as_of)
            if not current_file.is_file():
                raise FileNotFoundError(f"Current-month workbook not found: {current_file}")
            for path in self.get_history_paths(as_of, months_to_load):
                if not path.is_file():
                    self.logger.warning("Historical workbook missing; skipping: %s", path)
                    continue
                self.logger.info("Reading workbook: %s", path)
                data = pd.read_excel(
                    path, engine="openpyxl",
                    usecols=lambda name: str(name).strip().lower() in self.REQUIRED_COLUMNS,
                )
                data.columns = data.columns.astype(str).str.strip().str.lower()
                missing = set(self.REQUIRED_COLUMNS) - set(data.columns)
                if missing:
                    raise ValueError(f"Missing columns in {path}: {sorted(missing)}")
                data = data.loc[:, self.REQUIRED_COLUMNS].copy()
                selected = data["version"].astype("string").str.strip().str.casefold() == self.version.casefold()
                data = data.loc[selected.fillna(False)].copy()
                self.logger.info("Version-matched records in %s: %d", path.name, len(data))
                if data.empty:
                    continue
                data["cob_date"] = pd.to_datetime(data["cob_date"], errors="coerce").dt.normalize()
                for col in ("daily", "mtd", "ytd", "bus_day", "month"):
                    data[col] = pd.to_numeric(data[col], errors="coerce")
                for col in ("report_name", "indicator"):
                    data[col] = data[col].astype("string").str.strip()
                invalid = data[list(self.REQUIRED_COLUMNS)].isna().any(axis=1)
                invalid |= data["report_name"].eq("") | data["indicator"].eq("")
                if invalid.any():
                    raise ValueError(f"Found {int(invalid.sum())} invalid Daily rows in {path}")
                wrong_month = data["month"].ne(data["cob_date"].dt.month)
                if wrong_month.any():
                    raise ValueError(f"Month field does not match cob_date in {path}")
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
            return combined
        except Exception:
            self.logger.exception("Failed loading history for COB=%s", as_of.date())
            raise

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
