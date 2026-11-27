"""CLI entrypoint; e.g. python run_trend.py --root /data --cob-date 2027-01-05."""

import argparse
from pathlib import Path

import pandas as pd
from export_utils import make_excel
from Publication_Trendv2 import ReportTrendAnalyzer


def main():
    parser = argparse.ArgumentParser(
        description="Run Publication Trend analysis without Streamlit"
    )
    parser.add_argument(
        "--root",
        required=True,
        help="Root directory containing YYYY/Flat_Conso_Reports_Mon.xlsx",
    )
    parser.add_argument(
        "--cob-date",
        help="YYYY-MM-DD. Defaults to the latest COB for the selected version.",
    )
    parser.add_argument(
        "--version", default="Daily", choices=["Daily", "Flash", "Final"]
    )
    parser.add_argument("--comparison-cob", type=int, default=5)
    parser.add_argument("--months-to-load", type=int, default=3)
    parser.add_argument(
        "--trend-period",
        default="n_business_days",
        choices=["n_business_days", "current_quarter", "current_year", "all"],
        help="Audit label. The workbook contains all six trend windows.",
    )
    parser.add_argument(
        "--output", help="Output XLSX path; defaults to trend_<version>_<date>.xlsx"
    )
    args = parser.parse_args()
    if args.comparison_cob < 1 or args.months_to_load < 1:
        parser.error("comparison-cob and months-to-load must be positive")
    analyzer = ReportTrendAnalyzer(args.root, version=args.version)
    cob = args.cob_date or analyzer.get_latest_cob_by_version()[args.version.lower()]
    if cob is None:
        parser.error("No COB found for selected version")
    cob = pd.Timestamp(cob).strftime("%Y-%m-%d")
    output = Path(args.output or f"trend_{args.version.lower()}_{cob}.xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)
    result = analyzer.run_trend_analysis(
        cob, args.comparison_cob, args.months_to_load, args.trend_period
    )
    output.write_bytes(make_excel(result))
    print(
        f"Created {output.resolve()} | {len(result['trend'])} trend rows | {len(result['comparison'])} comparison rows"
    )


if __name__ == "__main__":
    main()
