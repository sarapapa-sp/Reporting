"""CLI/functional runner for Publication_Trendv2.ReportTrendAnalyzer."""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pandas as pd

from Publication_Trendv2 import ReportTrendAnalyzer


def run_trend_analysis(
    root_folder: str | Path,
    cob_date: str,
    n: int = 5,
    months_to_load: int = 3,
    version: str = "Daily",
    output_folder: str | Path | None = None,
    log_folder: str | Path | None = None,
    short_window: int = 5,
    long_window: int = 20,
    anomaly_z_threshold: float = 2.0,
    trend_period: str = "n_business_days",
) -> dict[str, Any]:
    """Read params and return traditional metrics, historical pivots, selected period trends.

    Quarter/year automatically load from period start; all loads from year start.
    Returns DataFrames and log_path, optionally saving result CSV files.
    """
    analyzer = ReportTrendAnalyzer(
        root_folder=root_folder, version=version, short_window=short_window,
        long_window=long_window, anomaly_z_threshold=anomaly_z_threshold,
        log_folder=log_folder, as_of_date=cob_date,
    )
    try:
        if trend_period not in analyzer.TREND_PERIODS:
            raise ValueError(f"trend_period must be one of {analyzer.TREND_PERIODS}")
        lookback = analyzer.required_months_to_load(cob_date, trend_period, months_to_load)
        analyzer.logger.info(
            "RUN START | cob_date=%s | n=%d | months_to_load=%d | trend_period=%s | actual_months=%d",
            cob_date, n, months_to_load, trend_period, lookback,
        )
        # Load and validate once; quarter/year automatically expand required history.
        history = analyzer.load_daily_data(cob_date, lookback)
        outputs: dict[str, Any] = {
            "trends": analyzer.calculate_trend_parameters(history, cob_date),
            "daily": analyzer.get_previous_daily_pnl(cob_date, n, months_to_load, history),
            "mtd": analyzer.get_previous_mtd_pnl(cob_date, n, months_to_load, history),
            "ytd": analyzer.get_previous_ytd_pnl(cob_date, n, months_to_load, history),
        }
        requested_periods = (
            ("n_business_days", "current_quarter", "current_year")
            if trend_period == "all" else (trend_period,)
        )
        for period in requested_periods:
            outputs[f"trends_{period}"] = analyzer.calculate_period_trend_parameters(
                history, cob_date, trend_period=period, n=n,
            )
            if period in ("current_quarter", "current_year"):
                # Daily P&L pivot for all COBs of the selected quarter/year.
                outputs[f"daily_{period}"] = analyzer.get_period_pnl(
                    cob_date, period, metric="daily", history=history,
                )
        if output_folder is not None:
            dest = Path(output_folder)
            dest.mkdir(parents=True, exist_ok=True)
            cob_label = pd.Timestamp(cob_date).strftime("%Y%m%d")
            for kind, frame in outputs.items():
                output_file = dest / f"{kind}_{cob_label}.csv"
                frame.to_csv(output_file, index=False)
                analyzer.logger.info("Output saved | %s | rows=%d", output_file, len(frame))
        analyzer.logger.info("RUN COMPLETE")
        outputs["log_path"] = analyzer.log_path
        return outputs
    except Exception:
        analyzer.logger.exception("RUN FAILED")
        raise
    finally:
        analyzer.close_logging()


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate report/indicator trend metrics and historical P&L")
    parser.add_argument("root_folder", help="Parent folder containing YYYY/Report_data_Mon.xlsx files")
    parser.add_argument("--as-of", required=True, dest="cob_date", help="COB date (YYYY-MM-DD)")
    parser.add_argument("--n", type=int, default=5, help="Number of distinct available COBs including as-of")
    parser.add_argument("--months", type=int, default=3, dest="months_to_load")
    parser.add_argument(
        "--trend-period", default="n_business_days",
        choices=ReportTrendAnalyzer.TREND_PERIODS,
        help="Compute N business days, current quarter, current year, or all period trends",
    )
    parser.add_argument("--version", default="Daily")
    parser.add_argument("--out-dir", dest="output_folder", help="Folder for four output CSV files")
    parser.add_argument("--log-dir", dest="log_folder", help="Folder for log files")
    parser.add_argument("--short-window", type=int, default=5)
    parser.add_argument("--long-window", type=int, default=20)
    parser.add_argument("--anomaly-z-threshold", type=float, default=2.0)
    args = vars(parser.parse_args())
    args["root_folder"] = args["root_folder"]
    results = run_trend_analysis(**args)
    for kind, result in results.items():
        if isinstance(result, pd.DataFrame):
            print(f"{kind}: {result.shape[0]} rows x {result.shape[1]} columns")
    print(f"Log: {results['log_path']}")


if __name__ == "__main__":
    main()
