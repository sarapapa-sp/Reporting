"""Command-line entry point for a configured report."""

import argparse
from datetime import date
from pathlib import Path

import pandas as pd

from report_builder import ReportBuilder


def sample_global_markets(version):
    # Sample snapshots have the same schema but different PnL values.
    pnl = {
        "Daily": [10, 20, 30, 40],
        "V1": [100, 200, 300, 400],
        "V2": [105, 210, 315, 420],
    }
    return pd.DataFrame({
        "Region": ["Asia", "Japan", "Eur", "Amer"],
        "Subregion": ["Asia", "Japan", "Europe", "Americas"],
        "Desk": ["HKG", "Tokyo", "Internal", "New York"],
        "PnL": pnl[version],
        "Revenue": [500, 700, 800, 900],
        "Trade_Count": [1, 1, 1, 1],
    })


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", default="global_markets", help="Report definition folder name")
    parser.add_argument("--date", type=date.fromisoformat, required=True, help="YYYY-MM-DD")
    parser.add_argument("--version", choices=["Daily", "V1", "V2"], required=True)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--csv", type=Path, help="Input snapshot for this date and version")
    source.add_argument("--sample", action="store_true", help="Use synthetic Global Markets data")
    args = parser.parse_args()

    if args.sample:
        if args.report != "global_markets":
            parser.error("--sample supplies data for global_markets only")
        dataframe = sample_global_markets(args.version)
    else:
        dataframe = pd.read_csv(args.csv)

    index = ReportBuilder().build_report(args.report, args.date, args.version, dataframe)
    print(index)


if __name__ == "__main__":
    main()
