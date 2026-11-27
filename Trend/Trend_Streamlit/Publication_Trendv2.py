"""Daily report/indicator trend analysis. Input: root/YYYY/Flat_Conso_Reports_Mon.xlsx."""

from __future__ import annotations

import logging
from pathlib import Path
from time import perf_counter

import numpy as np
import pandas as pd

PERIODS = ("short", "long", "broad", "month", "quarter", "fullyear")
MINIMUM = {
    "short": 5,
    "long": 60,
    "broad": 100,
    "month": 7,
    "quarter": 10,
    "fullyear": 20,
}
REQUIRED = ("cob_date", "version", "report_name", "indicator", "daily")
PAIR_COLUMNS = ["report_name", "indicator"]
COB_KEY = ["cob_date", *PAIR_COLUMNS]
CALENDAR_WINDOWS = {"month": "M", "quarter": "Q", "fullyear": "Y"}
TREND_METRICS = ("avg", "std", "zscore", "anomaly", "direction")
CHANGE_METRICS = ("min", "max", "change", "change_pct", "slope")
WINDOW_METADATA = ("start", "end", "observations", "status")


class ReportTrendAnalyzer:
    EXCEPTION_COBS = frozenset({"2026-01-01"})

    def __init__(
        self,
        root_folder,
        version="Daily",
        short_window=5,
        long_window=60,
        broad_window=100,
        anomaly_z_threshold=2.0,
        log_folder=None,
        as_of_date=None,
    ):
        self.root_folder = Path(root_folder)
        self.version = version
        self.windows = {
            "short": short_window,
            "long": long_window,
            "broad": broad_window,
        }
        if not (2 <= short_window <= long_window <= broad_window):
            raise ValueError("Invalid window sizes")
        self.anomaly_z_threshold = float(anomaly_z_threshold)
        if not np.isfinite(self.anomaly_z_threshold) or self.anomaly_z_threshold <= 0:
            raise ValueError("The anomaly threshold must be positive and finite.")
        self.log_folder = Path(log_folder) if log_folder else self.root_folder / "logs"
        self.logger = logging.Logger("trend")
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False
        self.logger.addHandler(logging.NullHandler())
        self.log_path = None
        if as_of_date is not None:
            self.setup_logging(as_of_date)

    def setup_logging(self, cob_date):
        """Write an audit log once per COB; never overwrite an existing audit."""
        self.close_logging()
        cob = pd.Timestamp(cob_date).normalize()
        self.log_folder.mkdir(parents=True, exist_ok=True)
        self.log_path = self.log_folder / f"log_trend_{cob:%Y%m%d}.log"
        # A prior execution already recorded this COB's audit. Do not append.
        # Exclusive creation avoids overwriting a log made by another process.
        try:
            handler = logging.FileHandler(self.log_path, mode="x", encoding="utf-8")
        except FileExistsError:
            return self.log_path
        handler.setFormatter(
            logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s")
        )
        self.logger.addHandler(handler)
        return self.log_path

    def close_logging(self):
        for handler in self.logger.handlers[:]:
            if isinstance(handler, logging.FileHandler):
                self.logger.removeHandler(handler)
                handler.close()

    def _log(self, cob):
        expected = self.log_folder / f"log_trend_{pd.Timestamp(cob):%Y%m%d}.log"
        if self.log_path != expected:
            self.setup_logging(cob)

    def generate_filename(self, cob_date):
        return f"Flat_Conso_Reports_{pd.Timestamp(cob_date):%b}.xlsx"

    def generate_file_path(self, cob_date):
        d = pd.Timestamp(cob_date)
        return self.root_folder / str(d.year) / self.generate_filename(d)

    def file_exists(self, cob_date):
        return self.generate_file_path(cob_date).is_file()

    def get_latest_cob_by_version(self):
        """Read each available monthly workbook and find max COB per version."""
        result = {"daily": None, "flash": None, "final": None}
        for path in sorted(self.root_folder.glob("*/Flat_Conso_Reports_*.xlsx")):
            try:
                df = pd.read_excel(
                    path,
                    usecols=lambda c: str(c).strip().lower() in ("cob_date", "version"),
                )
                df.columns = df.columns.str.strip().str.lower()
                if not {"cob_date", "version"}.issubset(df):
                    continue
                df["cob_date"] = pd.to_datetime(df["cob_date"], errors="coerce")
                for key, group in df.groupby(
                    df["version"].astype(str).str.strip().str.lower()
                ):
                    if key in result and group["cob_date"].notna().any():
                        value = group["cob_date"].max().normalize()
                        if result[key] is None or value > pd.Timestamp(result[key]):
                            result[key] = value.strftime("%Y-%m-%d")
            except Exception as exc:
                self.logger.warning("Unable to inspect workbook %s: %s", path.name, exc)
        return result

    def load_daily_data(self, as_of_date, months_to_load=3, *, heatmap_cobs=None):
        """Load backwards across years until each current report/indicator has broad history.

        Also covers the selected calendar year for full-year statistics. A missing
        historical workbook is logged; unavailable observations are never invented.
        """
        asof = pd.Timestamp(as_of_date).normalize()
        self._log(asof)
        current_path = self.generate_file_path(asof)
        if not self.file_exists(asof):
            raise FileNotFoundError(
                f"Missing selected-COB workbook: {current_path.name}"
            )
        period = asof.to_period("M")
        available = sorted(
            (
                p
                for p in self.root_folder.glob("*/Flat_Conso_Reports_*.xlsx")
                if p.is_file()
            )
        )
        earliest_year = min(
            (int(p.parent.name) for p in available if p.parent.name.isdigit()),
            default=asof.year,
        )
        earliest = pd.Period(f"{earliest_year}-01", freq="M")
        # Ensure full-year coverage, even if 100 observations are found earlier.
        min_months = max(int(months_to_load), asof.month)
        frames = []
        loaded = []
        current_keys = None
        required = self.windows["broad"]
        month_index = 0
        while period >= earliest:
            path = self.generate_file_path(period.start_time)
            loaded.append(path.name)
            if not path.is_file():
                self.logger.warning("Missing workbook: %s", path.name)
            else:
                df = self._read_month(path, asof)
                frames.append(df)
                if period == asof.to_period("M"):
                    current_keys = set(
                        map(
                            tuple,
                            df.loc[
                                df["cob_date"] == asof, ["report_name", "indicator"]
                            ].to_numpy(),
                        )
                    )
            month_index += 1
            if month_index >= min_months and current_keys and frames:
                accumulated = pd.concat(frames, ignore_index=True)
                if heatmap_cobs:
                    if self._has_heatmap_history(accumulated, heatmap_cobs, period):
                        break
                else:
                    historical = accumulated.loc[accumulated["cob_date"] < asof]
                    counts = historical.groupby(["report_name", "indicator"])[
                        "cob_date"
                    ].nunique()
                    if all(counts.get(key, 0) >= required for key in current_keys):
                        break
            period -= 1
        self.logger.info("Requested workbooks=%s", loaded)
        if not frames:
            raise ValueError("No history found")
        data = pd.concat(frames, ignore_index=True)
        if data.duplicated(COB_KEY).any():
            raise ValueError("Duplicate COB/report/indicator records")
        if current_keys:
            hist = (
                data[data["cob_date"] < asof]
                .groupby(["report_name", "indicator"])["cob_date"]
                .nunique()
            )
            lacking = {
                str(k): int(hist.get(k, 0))
                for k in current_keys
                if hist.get(k, 0) < required
            }
            if lacking:
                self.logger.warning(
                    "Broad history below %s COBs: %s", required, lacking
                )
        return data.sort_values(["report_name", "indicator", "cob_date"]).reset_index(
            drop=True
        )

    def _read_month(self, path, as_of):
        """Read one workbook and keep the requested version through the selected COB."""
        frame = pd.read_excel(path)
        frame.columns = frame.columns.astype(str).str.strip().str.lower()
        missing = set(REQUIRED) - set(frame.columns)
        if missing:
            raise ValueError(f"{path.name}: missing {sorted(missing)}")

        version = frame["version"].astype(str).str.strip().str.casefold()
        frame = frame.loc[version.eq(self.version.casefold())].copy()
        frame["cob_date"] = pd.to_datetime(
            frame["cob_date"], errors="coerce"
        ).dt.normalize()
        if frame["cob_date"].isna().any():
            raise ValueError(f"{path.name}: invalid cob_date for {self.version}")
        frame = frame.loc[frame["cob_date"].le(as_of)].copy()
        frame["daily"] = pd.to_numeric(frame["daily"], errors="coerce")
        if (
            frame[list(REQUIRED)].isna().any().any()
            or not np.isfinite(frame["daily"]).all()
        ):
            raise ValueError(
                f"{path.name}: missing or invalid values for {self.version}"
            )

        dates = sorted(frame["cob_date"].dt.strftime("%Y-%m-%d").unique())
        self.logger.info(
            "Workbook=%s | version=%s | matched_rows=%s | unique_reports=%s | unique_report_indicators=%s | unique_COBs=%s | COBs=%s",
            path.name,
            self.version,
            len(frame),
            frame["report_name"].nunique(),
            len(frame[PAIR_COLUMNS].drop_duplicates()),
            len(dates),
            dates,
        )
        return frame

    def _has_heatmap_history(self, history, n_cobs, loaded_month):
        """Check the baseline before the oldest visible observation of each pair."""
        displayed_dates = history["cob_date"].drop_duplicates().nlargest(n_cobs)
        if (
            len(displayed_dates) < n_cobs
            or loaded_month.year > displayed_dates.min().year
        ):
            return False

        visible = history.loc[history["cob_date"].isin(displayed_dates)]
        earliest_cobs = (
            visible.groupby(PAIR_COLUMNS)["cob_date"].min().rename("baseline_end")
        )
        history = history.join(earliest_cobs, on=PAIR_COLUMNS)
        baseline = history.loc[history["cob_date"] < history["baseline_end"]]
        counts = baseline.groupby(PAIR_COLUMNS)["cob_date"].nunique()
        return (
            counts.reindex(earliest_cobs.index, fill_value=0)
            .ge(self.windows["broad"])
            .all()
        )

    def _window(self, history, as_of, period):
        """Select observations strictly before the date being scored."""
        if period not in PERIODS:
            raise ValueError(f"period must be one of {PERIODS}")
        dates = history["cob_date"]
        if not dates.is_monotonic_increasing:
            history = history.sort_values("cob_date")
            dates = history["cob_date"]
        end = dates.searchsorted(as_of)
        if period in self.windows:
            start = max(0, end - self.windows[period])
        else:
            calendar_start = as_of.to_period(CALENDAR_WINDOWS[period]).start_time
            start = dates.searchsorted(calendar_start)
        return history.iloc[start:end]

    def _score_baseline(self, values, current):
        """Return mean, sample spread, z-score and anomaly for a valid baseline."""
        mean = float(np.mean(values))
        std = float(np.std(values, ddof=1))
        if std == 0:
            if current == mean:
                return mean, std, 0.0, False
            return mean, std, None, True
        score = (float(current) - mean) / std
        return mean, std, score, abs(score) >= self.anomaly_z_threshold

    def _stats(self, window, current, period):
        """Build the full set of export statistics for one trend window."""
        count = len(window)
        enough_history = count >= self.windows.get(period, MINIMUM[period])
        metrics = dict.fromkeys(TREND_METRICS + CHANGE_METRICS)
        if enough_history:
            values = window["daily"].to_numpy(dtype=float)
            mean, std, score, flagged = self._score_baseline(values, current)
            slope = float(np.polyfit(np.arange(count), values, 1)[0])
            direction = "Flat"
            if slope > 1e-10:
                direction = "Rising"
            elif slope < -1e-10:
                direction = "Falling"
            change = float(values[-1] - values[0])
            metrics.update(
                avg=mean,
                std=std,
                zscore=score,
                anomaly=flagged,
                direction=direction,
                min=float(values.min()),
                max=float(values.max()),
                change=change,
                change_pct=100 * change / abs(values[0]) if values[0] else None,
                slope=slope,
            )
        metrics.update(
            start=window["cob_date"].min().strftime("%Y-%m-%d") if count else None,
            end=window["cob_date"].max().strftime("%Y-%m-%d") if count else None,
            observations=count,
            status="ok" if enough_history else "insufficient_data",
        )
        return {f"{period}_{name}": value for name, value in metrics.items()}

    def calculate_trend_parameters(self, daily_data, as_of_date):
        asof = pd.Timestamp(as_of_date).normalize()
        self._log(asof)
        data = daily_data.copy()
        data["cob_date"] = pd.to_datetime(data["cob_date"]).dt.normalize()
        data = data[data["cob_date"] <= asof]
        rows = []
        for (report, indicator), grp in data.groupby(
            ["report_name", "indicator"], sort=True
        ):
            current = grp[grp["cob_date"] == asof]
            if current.empty:
                continue
            value = float(current.iloc[-1]["daily"])
            row = {
                "cob_date": asof,
                "report_name": report,
                "indicator": indicator,
                "daily": value,
                "version": self.version,
            }
            for period in PERIODS:
                window = self._window(grp, asof, period)
                row.update(self._stats(window, value, period))
            row["computed_cols"] = ", ".join(
                p for p in PERIODS if row[p + "_status"] == "ok"
            )
            rows.append(row)
        if not rows:
            raise ValueError(f"No selected COB rows for {asof.date()}")
        trendcols = [f"{p}_{k}" for p in PERIODS for k in TREND_METRICS]
        assoccols = [f"{p}_{k}" for p in PERIODS for k in CHANGE_METRICS]
        meta = ["daily", "version", "computed_cols"] + [
            f"{p}_{k}" for p in PERIODS for k in WINDOW_METADATA
        ]
        result = pd.DataFrame(rows)[
            ["cob_date", "report_name", "indicator"] + trendcols + assoccols + meta
        ]
        self.logger.info(
            "Trend output | report_indicator_pairs=%s | columns=%s",
            len(result),
            len(result.columns),
        )
        if len(self.logger.handlers) > 1:
            for period in PERIODS:
                good = result[f"{period}_status"].eq("ok")
                # Log the reference COBs once per period, rather than for every report.
                reference = self._window(data, asof, period)
                dates = sorted(
                    reference["cob_date"].dt.strftime("%Y-%m-%d").unique().tolist()
                )
                self.logger.info(
                    "Period=%s | required_observations=%s | successful_pairs=%s/%s | reference_COBs=%s | zscore_threshold=%s",
                    period,
                    self.windows.get(period, MINIMUM[period]),
                    int(good.sum()),
                    len(result),
                    dates,
                    self.anomaly_z_threshold,
                )
        return result

    def get_heatmap_statistics(self, history, dates, window):
        """Use the same statistics as the trend result for each visible COB."""
        columns = [
            "cob_date",
            "report_name",
            "indicator",
            "daily",
            "mean",
            "std",
            "zscore",
            "status",
        ]
        if window not in PERIODS:
            raise ValueError(f"window must be one of {PERIODS}")
        minimum = self.windows.get(window, MINIMUM[window])
        rows = []
        for (report, indicator), group in history.groupby(PAIR_COLUMNS, sort=True):
            group = group.sort_values("cob_date")
            current_rows = group.loc[group["cob_date"].isin(dates)]
            for current in current_rows.itertuples(index=False):
                values = self._window(group, current.cob_date, window)["daily"]
                mean = std = score = None
                status = "Insufficient history"
                if len(values) >= minimum:
                    mean, std, score, flagged = self._score_baseline(
                        values.to_numpy(dtype=float), current.daily
                    )
                    status = "Flagged" if flagged else "Within range"
                rows.append(
                    (
                        current.cob_date,
                        report,
                        indicator,
                        current.daily,
                        mean,
                        std,
                        score,
                        status,
                    )
                )
        return pd.DataFrame(rows, columns=columns)

    def generate_trend_parameters(self, as_of_date, months_to_load=3):
        return self.calculate_trend_parameters(
            self.load_daily_data(as_of_date, months_to_load), as_of_date
        )

    def _month_business_days(self, period):
        """Business calendar: weekdays excluding configured exception COBs.

        Supply organization holiday dates in EXCEPTION_COBS for exact matching.
        """
        dates = pd.bdate_range(period.start_time, period.end_time)
        return [
            d.normalize()
            for d in dates
            if d.strftime("%Y-%m-%d") not in self.EXCEPTION_COBS
        ]

    def select_previous_month_cob(self, current_cob):
        """Center-aligned business-day mapping from the original trend engine.

        Extra current-month days repeat the central previous-month day;
        extra previous-month days are skipped around the center.
        """
        current = pd.Timestamp(current_cob).normalize()
        current_days = self._month_business_days(current.to_period("M"))
        prior_days = self._month_business_days(current.to_period("M") - 1)
        if current not in current_days:
            raise ValueError(f"Not a business COB: {current.date()}")
        c = current_days.index(current)
        nc = len(current_days)
        np_ = len(prior_days)
        if nc >= np_:
            extra = nc - np_
            middle = (np_ + 1) // 2
            p = c if c < middle else max(middle, c - extra)
        else:
            extra = np_ - nc
            middle = (np_ - extra) // 2
            p = c if c < middle else c + extra
        p = max(0, min(p, np_ - 1))
        selected = prior_days[p]
        self.logger.info(
            "Previous-month COB mapping | current=%s | current bus_count=%s/%s | previous=%s | previous bus_count=%s/%s",
            current.date(),
            c + 1,
            nc,
            selected.date(),
            p + 1,
            np_,
        )
        self.logger.info(
            "Previous-month calendar COBs=%s",
            [x.strftime("%Y-%m-%d") for x in prior_days],
        )
        return selected

    def get_previous_month_comparison(self, history, cob_date, comparison_cob=5):
        """Return selected and prior current-month COB comparisons, per report/indicator."""
        asof = pd.Timestamp(cob_date).normalize()
        current = history[
            (history["cob_date"].dt.to_period("M") == asof.to_period("M"))
            & (history["cob_date"] <= asof)
        ]
        dates = sorted(current["cob_date"].unique(), reverse=True)[
            : int(comparison_cob)
        ]
        current = current[current["cob_date"].isin(dates)].copy()
        if current.empty:
            return pd.DataFrame()
        if "bus_count" not in current.columns:
            raise ValueError(
                "The selected month's workbook needs a bus_count column for comparison."
            )
        mapping = {pd.Timestamp(d): self.select_previous_month_cob(d) for d in dates}
        current["previous_month_cob"] = current["cob_date"].map(mapping)
        prior = history[["cob_date", "report_name", "indicator", "daily"]].rename(
            columns={"cob_date": "previous_month_cob", "daily": "previous_month_daily"}
        )
        current = current.merge(
            prior,
            on=["previous_month_cob", "report_name", "indicator"],
            how="left",
            validate="many_to_one",
        )
        previous_days = self._month_business_days(asof.to_period("M") - 1)
        day_numbers = {day: number for number, day in enumerate(previous_days, 1)}
        current["previous_month_bus_count"] = current["previous_month_cob"].map(
            day_numbers
        )
        current["previous_month_change"] = (
            current["daily"] - current["previous_month_daily"]
        )
        current["previous_month_change_pct"] = (
            100
            * current["previous_month_change"]
            / current["previous_month_daily"].abs().replace(0, float("nan"))
        )
        self.logger.info(
            "Previous-month comparison current COBs=%s | previous COBs=%s",
            sorted(current["cob_date"].dt.strftime("%Y-%m-%d").unique().tolist()),
            sorted(
                current["previous_month_cob"].dt.strftime("%Y-%m-%d").unique().tolist()
            ),
        )
        return (
            current[
                [
                    "cob_date",
                    "report_name",
                    "indicator",
                    "bus_count",
                    "daily",
                    "previous_month_cob",
                    "previous_month_bus_count",
                    "previous_month_daily",
                    "previous_month_change",
                    "previous_month_change_pct",
                ]
            ]
            .sort_values(
                ["cob_date", "report_name", "indicator"], ascending=[False, True, True]
            )
            .reset_index(drop=True)
        )

    def run_trend_analysis(
        self,
        cob_date,
        comparison_cob=5,
        months_to_load=3,
        trend_period="n_business_days",
        *,
        heatmap_cobs=None,
    ):
        self._log(pd.Timestamp(cob_date).normalize())
        started = perf_counter()
        self.logger.info(
            "START | cob_date=%s | version=%s | comparison_cob=%s | months_to_load=%s | trend_period=%s | windows=%s | minimum_observations=%s | zscore_threshold=%s | exception_COBs=%s | workbook_root=%s | log_file=%s",
            cob_date,
            self.version,
            comparison_cob,
            months_to_load,
            trend_period,
            self.windows,
            MINIMUM,
            self.anomaly_z_threshold,
            sorted(self.EXCEPTION_COBS),
            self.root_folder,
            self.log_path,
        )
        try:
            result = self._run_trend_analysis_impl(
                cob_date, comparison_cob, months_to_load, heatmap_cobs
            )
        except Exception:
            self.logger.exception(
                "FAILED | cob_date=%s | duration_seconds=%.3f",
                cob_date,
                perf_counter() - started,
            )
            raise
        else:
            self.logger.info(
                "SUCCESS | cob_date=%s | duration_seconds=%.3f | trend_rows=%s | history_rows=%s | comparison_rows=%s",
                cob_date,
                perf_counter() - started,
                len(result["trend"]),
                len(result["history"]),
                len(result["comparison"]),
            )
        finally:
            self.close_logging()
        return result

    def _run_trend_analysis_impl(
        self, cob_date, comparison_cob, months_to_load, heatmap_cobs=None
    ):
        history = self.load_daily_data(
            cob_date, months_to_load, heatmap_cobs=heatmap_cobs
        )
        trend = self.calculate_trend_parameters(history, cob_date)
        comparison = self.get_previous_month_comparison(
            history, cob_date, comparison_cob
        )
        # Keep both sides of the merge as normalized datetime64, never string vs datetime.
        trend["cob_date"] = pd.to_datetime(
            trend["cob_date"], errors="raise"
        ).dt.normalize()
        comparison["cob_date"] = pd.to_datetime(
            comparison["cob_date"], errors="raise"
        ).dt.normalize()
        today = comparison.loc[
            comparison["cob_date"] == pd.Timestamp(cob_date).normalize()
        ].copy()
        if not today.empty:
            cols = [
                "cob_date",
                "report_name",
                "indicator",
                "previous_month_cob",
                "previous_month_bus_count",
                "previous_month_daily",
                "previous_month_change",
                "previous_month_change_pct",
            ]
            trend = trend.merge(
                today[cols],
                on=["cob_date", "report_name", "indicator"],
                how="left",
                validate="one_to_one",
            )
            identity = ["cob_date", "report_name", "indicator"]
            trend = trend[identity + [c for c in trend.columns if c not in identity]]
        # Keep dates consistently typed within the calculation engine.
        return {"trend": trend, "history": history, "comparison": comparison}

    def get_daily_pnl_distribution(
        self, cob_date, occurrence, report_name, indicator, months_to_load=12
    ):
        asof = pd.Timestamp(cob_date).normalize()
        history = self.load_daily_data(asof, months_to_load)
        group = history[
            (history["report_name"] == report_name)
            & (history["indicator"] == indicator)
        ]
        current = group[group["cob_date"] == asof]
        if current.empty:
            raise ValueError("No selected COB for report and indicator")
        window = self._window(group, asof, occurrence)
        stats = self._stats(window, float(current.iloc[-1]["daily"]), occurrence)
        avg = stats[f"{occurrence}_avg"]
        std = stats[f"{occurrence}_std"]
        threshold = self.anomaly_z_threshold
        return {
            "cob_date": asof.strftime("%Y-%m-%d"),
            "report_name": report_name,
            "indicator": indicator,
            "occurrence": occurrence,
            "current_daily": float(current.iloc[-1]["daily"]),
            "mean": avg,
            "std": std,
            "zscore": stats[f"{occurrence}_zscore"],
            "zscore_threshold": threshold,
            "lower_endpoint": avg - threshold * std if avg is not None else None,
            "upper_endpoint": avg + threshold * std if avg is not None else None,
            "historical_values": window["daily"].tolist(),
            "status": stats[f"{occurrence}_status"],
            "period_start": stats[f"{occurrence}_start"],
            "period_end": stats[f"{occurrence}_end"],
        }

    def get_previous_daily_pnl(self, as_of_date, n, months_to_load=3, history=None):
        data = (
            history
            if history is not None
            else self.load_daily_data(as_of_date, months_to_load)
        )
        data = data[data["cob_date"] <= pd.Timestamp(as_of_date)]
        dates = sorted(data["cob_date"].unique(), reverse=True)[:n]
        selected = data[data["cob_date"].isin(dates)].copy()
        selected["cob_date"] = selected["cob_date"].dt.strftime("%Y-%m-%d")
        cols = [pd.Timestamp(d).strftime("%Y-%m-%d") for d in dates]
        return (
            selected.pivot(
                index=["report_name", "indicator"], columns="cob_date", values="daily"
            )
            .reindex(columns=cols)
            .reset_index()
        )
