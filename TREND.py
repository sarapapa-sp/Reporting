
from pathlib import Path
import calendar

import numpy as np
import pandas as pd


class TrendEngine:
    """
    PNL TREND calculation engine.

    Expected columns:
        COB date, Version, month_number, business_day,
        Report name, Indicator, daily, mtd, ytd

    Monthly files:
        Conso_Reports_Jan.xlsx
        Conso_Reports_Feb.xlsx
        ...
        Conso_Reports_Dec.xlsx

    Features:
        1. Reads multiple monthly files.
        2. Detects missing expected COB dates.
        3. Stops TREND calculations if COB dates are missing.
        4. Calculates Daily, MTD, and YTD TREND metrics.
        5. Exports results to Excel.
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

        self.loaded_start_date = None
        self.loaded_through = None
        self.missing_files = []

        if history_window < 1:
            raise ValueError("history_window must be at least 1.")

        if recent_window < 1:
            raise ValueError("recent_window must be at least 1.")

        if min_history < 1:
            raise ValueError("min_history must be at least 1.")

    # ---------------------------------------------------------
    # MONTHLY FILE HANDLING
    # ---------------------------------------------------------

    def _get_month_file(self, year, month):
        """
        Return the monthly file path.

        The provided filename does not include a year.
        Therefore, this method assumes one year's files per
        data directory.
        """
        month_abbr = calendar.month_abbr[month]

        return (
            self.data_dir
            / f"Conso_Reports_{month_abbr}.xlsx"
        )

    @staticmethod
    def _months_to_read(selected_date, months_back):
        """
        Return the current month and preceding months.

        Example:
            selected_date = 2026-09-10
            months_back = 2

        Reads July, August, and September.
        """
        selected_period = selected_date.to_period("M")

        periods = []

        for offset in range(months_back, -1, -1):
            period = selected_period - offset
            periods.append((period.year, period.month))

        return periods

    def load_data(
        self,
        selected_date,
        months_back=3,
        sheet_name=0,
    ):
        """
        Load the selected month's file and preceding months.

        Future COB dates are excluded.

        Missing monthly files are recorded. Missing expected
        COB dates are subsequently detected by validation.
        """
        selected_date = pd.Timestamp(
            selected_date
        ).normalize()

        if months_back < 0:
            raise ValueError("months_back cannot be negative.")

        periods = self._months_to_read(
            selected_date,
            months_back,
        )

        first_year, first_month = periods[0]

        self.loaded_start_date = pd.Timestamp(
            year=first_year,
            month=first_month,
            day=1,
        )

        frames = []
        self.missing_files = []

        for year, month in periods:
            file_path = self._get_month_file(year, month)

            if not file_path.exists():
                self.missing_files.append(
                    str(file_path)
                )
                continue

            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
            )

            df.columns = [
                str(column).strip()
                for column in df.columns
            ]

            missing_columns = (
                set(self.REQUIRED_COLUMNS)
                - set(df.columns)
            )

            if missing_columns:
                raise ValueError(
                    f"{file_path.name} is missing columns: "
                    f"{sorted(missing_columns)}"
                )

            df["COB date"] = pd.to_datetime(
                df["COB date"],
                errors="coerce",
            ).dt.normalize()

            for column in [
                "Version",
                "Report name",
                "Indicator",
            ]:
                df[column] = (
                    df[column]
                    .astype("string")
                    .str.strip()
                )

            for column in ["daily", "mtd", "ytd"]:
                df[column] = pd.to_numeric(
                    df[column],
                    errors="coerce",
                )

            df = df.dropna(
                subset=[
                    "COB date",
                    "Version",
                    "Report name",
                    "Indicator",
                ]
            )

            # Exclude future observations.
            df = df[
                df["COB date"] <= selected_date
            ]

            frames.append(df)

        if not frames:
            raise FileNotFoundError(
                f"No usable monthly Excel files found in "
                f"{self.data_dir}"
            )

        self.data = pd.concat(
            frames,
            ignore_index=True,
        )

        # One row per COB/version/report/indicator.
        # If duplicates exist, retain the last occurrence.
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
            [
                "Report name",
                "Indicator",
                "COB date",
                "Version",
            ]
        ).reset_index(drop=True)

        self.loaded_through = selected_date

        return self.data.copy()

    # ---------------------------------------------------------
    # MISSING COB VALIDATION
    # ---------------------------------------------------------

    def validate_missing_cobs(
        self,
        selected_date,
        holidays=None,
    ):
        """
        Check whether expected weekdays have any COB records
        between the start of the loaded period and selected date.

        If one or more dates are missing, raise ValueError.
        TREND calculations must not continue.

        holidays example:
            ["2026-08-15", "2026-08-28"]

        IMPORTANT:
        Weekdays are assumed to be expected COB days, except
        dates supplied in holidays. Use your official COB
        calendar if it differs from this assumption.
        """
        selected_date = pd.Timestamp(
            selected_date
        ).normalize()

        if self.data.empty:
            raise ValueError(
                "No data loaded. Run load_data() first."
            )

        if self.loaded_start_date is None:
            raise ValueError(
                "The loaded history start date is unavailable."
            )

        if selected_date > self.loaded_through:
            raise ValueError(
                "Selected date is later than the loaded data. "
                "Run load_data() again."
            )

        # Generate expected Monday-Friday dates.
        expected_dates = pd.bdate_range(
            start=self.loaded_start_date,
            end=selected_date,
        )

        # Exclude configured holidays.
        holiday_dates = pd.DatetimeIndex(
            pd.to_datetime(holidays or [])
        ).normalize()

        if len(holiday_dates):
            expected_dates = expected_dates.difference(
                holiday_dates
            )

        # A date counts as available if any row exists on it.
        available_dates = pd.DatetimeIndex(
            self.data["COB date"]
            .dropna()
            .unique()
        ).normalize()

        missing_dates = expected_dates.difference(
            available_dates
        )

        if len(missing_dates):
            missing_list = [
                date.strftime("%Y-%m-%d")
                for date in missing_dates
            ]

            message = (
                "\nTREND EXECUTION STOPPED\n"
                "Missing expected COB dates detected:\n"
                + "\n".join(
                    f"  - {date}"
                    for date in missing_list
                )
            )

            if self.missing_files:
                message += (
                    "\n\nMissing monthly files:\n"
                    + "\n".join(
                        f"  - {name}"
                        for name in self.missing_files
                    )
                )

            raise ValueError(message)

        # Missing files should also stop execution explicitly,
        # even if the missing period contains no expected dates.
        if self.missing_files:
            raise FileNotFoundError(
                "TREND execution stopped because monthly "
                "input files are missing:\n"
                + "\n".join(
                    f"  - {name}"
                    for name in self.missing_files
                )
            )

        print(
            "COB validation successful: "
            "no missing expected COB dates found."
        )

        return True

    # ---------------------------------------------------------
    # HISTORICAL TREND HELPERS
    # ---------------------------------------------------------

    @staticmethod
    def _safe_std(series, ddof=1):
        """Return standard deviation, or NaN if insufficient."""
        series = pd.to_numeric(
            series,
            errors="coerce",
        ).dropna()

        if len(series) <= ddof:
            return np.nan

        return float(series.std(ddof=ddof))

    def _get_history(self, group, selected_date):
        """
        Return previous Daily observations only.

        The selected COB is excluded.
        """
        history = group[
            (
                group["Version"].str.lower()
                == "daily"
            )
            & (
                group["COB date"]
                < selected_date
            )
        ].sort_values("COB date")

        return history.tail(self.history_window)

    def _get_previous_cob(self, group, current_row):
        """
        Find the previous available COB for the same
        report and indicator.

        Prefer the same version on the previous date.
        If unavailable, use Daily for that date.

        Global COB validation runs before this method is used.
        """
        prior = group[
            group["COB date"]
            < current_row["COB date"]
        ].copy()

        if prior.empty:
            return None

        previous_date = prior["COB date"].max()

        previous_rows = prior[
            prior["COB date"] == previous_date
        ]

        same_version = previous_rows[
            previous_rows["Version"].str.lower()
            == str(current_row["Version"]).lower()
        ]

        if not same_version.empty:
            return same_version.iloc[-1]

        daily_rows = previous_rows[
            previous_rows["Version"].str.lower()
            == "daily"
        ]

        if not daily_rows.empty:
            return daily_rows.iloc[-1]

        return previous_rows.iloc[-1]

    # ---------------------------------------------------------
    # TREND CALCULATION FOR ONE REPORT/INDICATOR
    # ---------------------------------------------------------

    def _calculate_one(
        self,
        group,
        current_row,
        selected_date,
    ):
        """Calculate TREND metrics for one report/indicator."""

        history = self._get_history(
            group,
            selected_date,
        )

        daily_history = history["daily"].dropna()
        history_count = len(daily_history)

        current_daily = current_row["daily"]

        historical_mean = (
            daily_history.mean()
            if history_count
            else np.nan
        )

        historical_std = self._safe_std(
            daily_history,
            self.ddof,
        )

        # Daily Z-score.
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

        # Recent volatility compared with historical volatility.
        recent_values = daily_history.tail(
            self.recent_window
        )

        recent_std = self._safe_std(
            recent_values,
            self.ddof,
        )

        if (
            pd.notna(recent_std)
            and pd.notna(historical_std)
            and historical_std > 0
        ):
            volatility_ratio = (
                recent_std / historical_std
            )
        else:
            volatility_ratio = np.nan

        daily_avg_recent = (
            recent_values.mean()
            if len(recent_values)
            else np.nan
        )

        previous_row = self._get_previous_cob(
            group,
            current_row,
        )

        if previous_row is None:
            mtd_change = np.nan
            ytd_change = np.nan
            mtd_daily_error = np.nan
            ytd_daily_error = np.nan

        else:
            # MTD movement from the previous available COB.
            if (
                pd.notna(current_row["mtd"])
                and pd.notna(previous_row["mtd"])
            ):
                mtd_change = (
                    current_row["mtd"]
                    - previous_row["mtd"]
                )
            else:
                mtd_change = np.nan

            # YTD movement from the previous available COB.
            if (
                pd.notna(current_row["ytd"])
                and pd.notna(previous_row["ytd"])
            ):
                ytd_change = (
                    current_row["ytd"]
                    - previous_row["ytd"]
                )
            else:
                ytd_change = np.nan

            # Cumulative movement should reconcile to Daily PNL.
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

            # Current values
            "daily": current_daily,
            "mtd": current_row["mtd"],
            "ytd": current_row["ytd"],

            # Daily TREND parameters
            "history_count": history_count,
            "daily_mean_20": historical_mean,
            "daily_std_20": historical_std,
            "daily_zscore_20": daily_zscore,
            "daily_avg_5": daily_avg_recent,
            "volatility_ratio_5_20": volatility_ratio,

            # MTD/YTD TREND parameters
            "mtd_change": mtd_change,
            "ytd_change": ytd_change,
            "mtd_daily_error": mtd_daily_error,
            "ytd_daily_error": ytd_daily_error,
        }

    # ---------------------------------------------------------
    # MAIN TREND METHOD
    # ---------------------------------------------------------

    def get_trend_parameters(
        self,
        selected_date,
        version="Daily",
        months_back=3,
        holidays=None,
        sheet_name=0,
    ):
        """
        Calculate TREND parameters for the selected COB/version.

        Validation runs before any TREND calculation.
        Missing expected COB dates stop execution.
        """
        selected_date = pd.Timestamp(
            selected_date
        ).normalize()

        # Clear old results so a failed run cannot leave
        # previous results available for accidental export.
        self.trend_results = pd.DataFrame()

        # Reload when data is absent or the requested date
        # is beyond the currently loaded date range.
        if (
            self.data.empty
            or self.loaded_through is None
            or selected_date > self.loaded_through
            or selected_date < self.loaded_start_date
        ):
            self.load_data(
                selected_date=selected_date,
                months_back=months_back,
                sheet_name=sheet_name,
            )

        # Critical: stop if COB dates are missing.
        self.validate_missing_cobs(
            selected_date=selected_date,
            holidays=holidays,
        )

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
                f"No records found for "
                f"{selected_date.date()} "
                f"with version {version!r}."
            )

        results = []

        for _, current_row in current.iterrows():
            mask = (
                (
                    data["Report name"]
                    == current_row["Report name"]
                )
                & (
                    data["Indicator"]
                    == current_row["Indicator"]
                )
            )

            group = data[mask].copy()

            results.append(
                self._calculate_one(
                    group=group,
                    current_row=current_row,
                    selected_date=selected_date,
                )
            )

        self.trend_results = pd.DataFrame(results)

        return self.trend_results.copy()

    # ---------------------------------------------------------
    # EXPORT
    # ---------------------------------------------------------

    def save_results(
        self,
        output_path="trend_parameters.xlsx",
    ):
        """Export the most recent successful TREND calculation."""

        if self.trend_results.empty:
            raise ValueError(
                "No TREND results available. "
                "Run get_trend_parameters() successfully first."
            )

        output_path = Path(output_path)

        self.trend_results.to_excel(
            output_path,
            index=False,
            sheet_name="TREND_Parameters",
        )

        return str(output_path)


###########################################
# Usage
###########################################

from trend_engine import TrendEngine


def main():
    engine = TrendEngine(
        data_dir="data",
        history_window=20,
        recent_window=5,
        min_history=5,
    )

    selected_date = "2026-09-10"
    version = "Daily"

    try:
        trend_df = engine.get_trend_parameters(
            selected_date=selected_date,
            version=version,
            months_back=3,
            holidays=[
                # Add official COB holidays here.
                # Example: "2026-08-15"
            ],
        )

        # Export only if validation and calculations succeed.
        engine.save_results(
            "trend_parameters.xlsx"
        )

        print("\nTREND parameters generated successfully.")
        print(trend_df.to_string(index=False))

    except (ValueError, FileNotFoundError) as error:
        print(f"\nERROR:\n{error}")
        print("TREND calculation was stopped.")


if __name__ == "__main__":
    main()








