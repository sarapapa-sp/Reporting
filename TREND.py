
from pathlib import Path
import calendar
import pandas as pd
import numpy as np


class TrendEngine:
    """
    Generates TREND parameters for each Report name + Indicator.

    Expected input columns:
        COB date, Version, month_number, business_day,
        Report name, Indicator, daily, mtd, ytd

    Monthly file naming:
        Conso_Reports_Jan.xlsx
        Conso_Reports_Feb.xlsx
        ...
        Conso_Reports_Dec.xlsx

    Historical trend statistics use Daily observations only.
    """

    REQUIRED_COLUMNS = [
        "COB date",
        "Version",
        "month_number",
        "business_day",
        "Report name",
        "Indicator",
        "daily",
        "mtd",
        "ytd",
    ]

    def __init__(
        self,
        data_dir,
        history_window=20,
        recent_window=5,
        min_history=5,
        ddof=1,
    ):
        self.data_dir = Path(data_dir)
        self.history_window = history_window
        self.recent_window = recent_window
        self.min_history = min_history
        self.ddof = ddof

        self.data = pd.DataFrame()
        self.trend_results = pd.DataFrame()

    @staticmethod
    def _month_abbr(date):
        """Return the English three-letter month abbreviation."""
        return calendar.month_abbr[date.month]

    def _get_month_file(self, year, month):
        """
        Resolve the monthly file.

        Note: the specified naming convention does not contain a year.
        Therefore, this implementation assumes one year's files per
        data directory.
        """
        month_abbr = calendar.month_abbr[month]
        file_path = (
            self.data_dir / f"Conso_Reports_{month_abbr}.xlsx"
        )

        if not file_path.exists():
            return None

        return file_path

    def _months_to_read(self, selected_date, months_back=3):
        """Return current month and preceding months, chronologically."""
        periods = []

        for offset in range(months_back, -1, -1):
            period = (
                selected_date.to_period("M") - offset
            )
            periods.append((period.year, period.month))

        return periods

    def load_data(
        self,
        selected_date,
        months_back=3,
        sheet_name=0,
    ):
        """
        Load the selected month and preceding monthly files.

        months_back=3 reads up to three preceding months. This is
        normally sufficient for 20 previous COB observations, but
        missing data may require increasing it.
        """
        selected_date = pd.Timestamp(selected_date).normalize()

        frames = []
        missing_files = []

        for year, month in self._months_to_read(
            selected_date, months_back
        ):
            file_path = self._get_month_file(year, month)

            if file_path is None:
                missing_files.append(file_path)
                continue

            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
            )

            # Normalize column names.
            df.columns = [
                str(col).strip() for col in df.columns
            ]

            missing_columns = (
                set(self.REQUIRED_COLUMNS) - set(df.columns)
            )

            if missing_columns:
                raise ValueError(
                    f"{file_path.name} is missing columns: "
                    f"{sorted(missing_columns)}"
                )

            df["COB date"] = pd.to_datetime(
                df["COB date"], errors="coerce"
            ).dt.normalize()

            df["Version"] = (
                df["Version"].astype("string").str.strip()
            )
            df["Report name"] = (
                df["Report name"].astype("string").str.strip()
            )
            df["Indicator"] = (
                df["Indicator"].astype("string").str.strip()
            )

            for col in ["daily", "mtd", "ytd"]:
                df[col] = pd.to_numeric(
                    df[col], errors="coerce"
                )

            df = df.dropna(
                subset=[
                    "COB date",
                    "Version",
                    "Report name",
                    "Indicator",
                ]
            )

            # Do not use future observations.
            df = df[df["COB date"] <= selected_date]

            frames.append(df)

        if not frames:
            raise FileNotFoundError(
                f"No usable monthly files found in {self.data_dir}"
            )

        self.data = pd.concat(
            frames, ignore_index=True
        )

        self.data = self.data.drop_duplicates(
            subset=[
                "COB date",
                "Version",
                "Report name",
                "Indicator",
            ],
            keep="last",
        )

        self.data = self.data.sort_values(
            ["Report name", "Indicator", "COB date", "Version"]
        ).reset_index(drop=True)

        self.loaded_through = selected_date
        self.missing_files = [
            str(path) for path in missing_files
        ]

        return self.data.copy()

    @staticmethod
    def _safe_std(series, ddof=1):
        """Return standard deviation, or NaN if insufficient data."""
        series = pd.to_numeric(series, errors="coerce").dropna()

        if len(series) <= ddof:
            return np.nan

        return float(series.std(ddof=ddof))

    def _get_history(self, group, selected_date):
        """
        Get the preceding Daily observations only.
        The selected COB itself is excluded.
        """
        history = group[
            (group["Version"].str.lower() == "daily")
            & (group["COB date"] < selected_date)
        ].sort_values("COB date")

        return history.tail(self.history_window)

    def _get_previous_cob(self, group, current_row):
        """
        Find the previous COB observation.

        Prefer the same version as the current row. If that version
        is unavailable on the previous COB, fall back to Daily.
        """
        prior = group[
            group["COB date"] < current_row["COB date"]
        ].copy()

        if prior.empty:
            return None

        prior_date = prior["COB date"].max()
        prior_at_date = prior[
            prior["COB date"] == prior_date
        ]

        same_version = prior_at_date[
            prior_at_date["Version"].str.lower()
            == str(current_row["Version"]).lower()
        ]

        if not same_version.empty:
            return same_version.iloc[-1]

        daily = prior_at_date[
            prior_at_date["Version"].str.lower() == "daily"
        ]

        if not daily.empty:
            return daily.iloc[-1]

        return prior_at_date.iloc[-1]

    def _calculate_one(
        self,
        group,
        current_row,
        selected_date,
    ):
        """Calculate TREND metrics for one report/indicator."""
        history = self._get_history(
            group, selected_date
        )

        daily_history = history["daily"].dropna()
        current_daily = current_row["daily"]

        history_count = len(daily_history)
        historical_mean = daily_history.mean()
        historical_std = self._safe_std(
            daily_history, self.ddof
        )

        # Daily Z-score against preceding Daily observations.
        if (
            pd.notna(current_daily)
            and history_count >= self.min_history
            and pd.notna(historical_std)
            and historical_std > 0
        ):
            daily_zscore = (
                current_daily - historical_mean
            ) / historical_std
        else:
            daily_zscore = np.nan

        # Recent volatility relative to historical volatility.
        # The recent window excludes the selected COB.
        recent_values = daily_history.tail(
            self.recent_window
        )

        recent_std = self._safe_std(
            recent_values, self.ddof
        )

        if (
            pd.notna(recent_std)
            and pd.notna(historical_std)
            and historical_std > 0
        ):
            volatility_ratio = recent_std / historical_std
        else:
            volatility_ratio = np.nan

        daily_avg_5 = (
            recent_values.mean()
            if len(recent_values) > 0
            else np.nan
        )

        previous_row = self._get_previous_cob(
            group, current_row
        )

        if previous_row is None:
            mtd_change = np.nan
            ytd_change = np.nan
            mtd_daily_error = np.nan
            ytd_daily_error = np.nan
        else:
            mtd_change = (
                current_row["mtd"] - previous_row["mtd"]
                if pd.notna(current_row["mtd"])
                and pd.notna(previous_row["mtd"])
                else np.nan
            )

            ytd_change = (
                current_row["ytd"] - previous_row["ytd"]
                if pd.notna(current_row["ytd"])
                and pd.notna(previous_row["ytd"])
                else np.nan
            )

            # Reconciliation checks:
            # cumulative change should equal current daily PNL.
            mtd_daily_error = (
                mtd_change - current_daily
                if pd.notna(mtd_change)
                and pd.notna(current_daily)
                else np.nan
            )

            ytd_daily_error = (
                ytd_change - current_daily
                if pd.notna(ytd_change)
                and pd.notna(current_daily)
                else np.nan
            )

        return {
            "COB date": current_row["COB date"],
            "Version": current_row["Version"],
            "month_number": current_row["month_number"],
            "business_day": current_row["business_day"],
            "Report name": current_row["Report name"],
            "Indicator": current_row["Indicator"],

            # Current PNL values
            "daily": current_daily,
            "mtd": current_row["mtd"],
            "ytd": current_row["ytd"],

            # Daily trend parameters
            "history_count": history_count,
            "daily_mean_20": historical_mean,
            "daily_std_20": historical_std,
            "daily_zscore_20": daily_zscore,
            "daily_avg_5": daily_avg_5,
            "volatility_ratio_5_20": volatility_ratio,

            # MTD/YTD trend parameters
            "mtd_change": mtd_change,
            "ytd_change": ytd_change,
            "mtd_daily_error": mtd_daily_error,
            "ytd_daily_error": ytd_daily_error,
        }

    def get_trend_parameters(
        self,
        selected_date,
        version="Daily",
    ):
        """
        Generate TREND metrics for a selected COB date/version.

        Returns one row per Report name + Indicator.
        """
        selected_date = pd.Timestamp(
            selected_date
        ).normalize()

        if self.data.empty:
            self.load_data(selected_date)

        data = self.data.copy()

        current = data[
            (data["COB date"] == selected_date)
            & (
                data["Version"].str.lower()
                == version.lower()
            )
        ]

        if current.empty:
            raise ValueError(
                f"No records found for {selected_date.date()} "
                f"with version {version!r}. "
                "Check the COB date, version, and monthly file."
            )

        results = []

        for _, current_row in current.iterrows():
            mask = (
                (data["Report name"] == current_row["Report name"])
                & (data["Indicator"] == current_row["Indicator"])
            )
            group = data[mask].copy()

            result = self._calculate_one(
                group,
                current_row,
                selected_date,
            )
            results.append(result)

        self.trend_results = pd.DataFrame(results)

        return self.trend_results.copy()

    def save_results(
        self,
        output_path="trend_parameters.xlsx",
    ):
        """Save the latest TREND results to Excel."""
        if self.trend_results.empty:
            raise ValueError(
                "No TREND results available. "
                "Run get_trend_parameters() first."
            )

        self.trend_results.to_excel(
            output_path,
            index=False,
            sheet_name="TREND_Parameters",
        )

        return output_path






################################
# USAGE
################################


from trend_engine import TrendEngine

# Initialize the engine
engine = TrendEngine(
    data_dir="data",
    history_window=20,
    recent_window=5,
    min_history=5,
)

# Select the COB date and version
selected_date = "2026-09-09"
version = "Daily"

# Load the selected month's data and previous months
engine.load_data(
    selected_date=selected_date,
    months_back=3,
)

# Calculate TREND parameters
trend_df = engine.get_trend_parameters(
    selected_date=selected_date,
    version=version,
)

# Display the results
print(trend_df.to_string(index=False))

# Export results
engine.save_results(
    "trend_parameters.xlsx"
)

print("TREND parameters generated successfully.")






