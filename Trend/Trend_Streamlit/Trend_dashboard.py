"""Streamlit dashboard for Global Markets daily P&L."""

import os
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from export_utils import make_excel
from Heatmap_Trend import (
    generate_line_graph,
    generate_report_line_graph,
    generate_zscore_heatmap,
    history_to_pivot,
)
from Publication_Trendv2 import ReportTrendAnalyzer

TITLE = "Global Markets | Trend Analysis"
WINDOW_LABELS = {
    "short": "Short (5 COBs)",
    "long": "Long (60 COBs)",
    "broad": "Broad (100 COBs)",
    "month": "Month",
    "quarter": "Quarter",
    "fullyear": "Full year",
}
ANOMALY_STATUSES = ("Flagged", "Within range", "Insufficient history")


@st.cache_data(ttl=300)
def latest_daily_cob(root):
    return ReportTrendAnalyzer(root).get_latest_cob_by_version()["daily"]


@st.cache_data(ttl=300, show_spinner="Loading daily P&L...")
def load_trend(
    root,
    cob,
    comparison_cobs=5,
    months=3,
    period="n_business_days",
    *,
    heatmap_cobs=None,
):
    analyzer = ReportTrendAnalyzer(root)
    return analyzer.run_trend_analysis(
        cob, comparison_cobs, months, period, heatmap_cobs=heatmap_cobs
    )


def show_anomaly_overview(stats, cob, window, threshold):
    st.subheader("Anomaly Overview")
    st.caption(
        f"{cob:%d %b %Y} · {WINDOW_LABELS[window]} · Flagged at |z| ≥ {threshold:g}"
    )

    overview = stats.loc[stats["cob_date"].eq(cob)].copy()
    counts = overview["status"].value_counts()
    for column, status in zip(st.columns(3), ANOMALY_STATUSES):
        column.metric(status, counts.get(status, 0))
    if st.toggle("Flagged only"):
        overview = overview.loc[overview["status"].eq("Flagged")].copy()

    money_columns = {
        "daily": "Daily P&L (mEUR)",
        "mean": "Historical mean (mEUR)",
        "std": "Standard deviation (mEUR)",
    }
    overview[list(money_columns)] = overview[list(money_columns)].div(1000)
    headers = {
        "report_name": "Report",
        "indicator": "Indicator",
        **money_columns,
        "zscore": "Z-score",
        "status": "Status",
    }
    numbers = [*money_columns.values(), "Z-score"]
    st.dataframe(
        overview[list(headers)].rename(columns=headers),
        hide_index=True,
        use_container_width=True,
        column_config={
            name: st.column_config.NumberColumn(format="%.2f") for name in numbers
        },
    )


def show_line_graphs(pivot, n_cobs):
    st.subheader("Daily P&L by Report")
    st.caption(
        "One line per selected report. Daily P&L is summed across the selected indicators."
    )
    st.plotly_chart(generate_report_line_graph(pivot, n_cobs), use_container_width=True)
    st.subheader("Daily P&L by Report and Indicator")
    st.plotly_chart(
        generate_line_graph(pivot, "All Reports", n_cobs=n_cobs),
        use_container_width=True,
    )


def show_excel_export(root, cob):
    st.subheader("Save Trend Analysis")
    with st.form("export"):
        date_column, comparison_column, months_column, period_column = st.columns(4)
        export_cob = date_column.date_input("Export COB", cob.date())
        comparison = comparison_column.number_input(
            "Comparison COBs", min_value=1, value=5
        )
        months = months_column.number_input("Months to load", min_value=1, value=3)
        period = period_column.selectbox(
            "Trend period",
            ["n_business_days", "current_quarter", "current_year", "all"],
        )
        submitted = st.form_submit_button("Prepare Excel")

    if submitted:
        result = load_trend(root, str(export_cob), comparison, months, period)
        st.session_state["export_bytes"] = make_excel(result)
        st.session_state["export_date"] = str(export_cob)
    if st.session_state.get("export_bytes"):
        st.download_button(
            "Download Excel (mEUR)",
            st.session_state["export_bytes"],
            file_name=f"publication_trend_{st.session_state['export_date']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    st.caption(
        "Excel includes all six trend windows. The export period is recorded on the first audited run for a COB."
    )


def show_threshold_analysis(analyzer, trend, history, cob):
    st.subheader("Daily P&L Threshold Analysis")
    report_column, indicator_column, window_column = st.columns(3)
    report = report_column.selectbox("Report", sorted(history["report_name"].unique()))
    report_history = history.loc[history["report_name"].eq(report)]
    indicator = indicator_column.selectbox(
        "Indicator", sorted(report_history["indicator"].unique())
    )
    window = window_column.selectbox("Window", tuple(WINDOW_LABELS))
    matches = trend.loc[
        trend["report_name"].eq(report) & trend["indicator"].eq(indicator)
    ]
    if matches.empty:
        st.info("No trend result for selected report and indicator.")
        return

    row = matches.iloc[0]
    if row[f"{window}_status"] != "ok":
        st.info("Insufficient historical observations.")
        return

    pair_history = report_history.loc[report_history["indicator"].eq(indicator)]
    reference = analyzer._window(pair_history, cob, window)
    mean = row[f"{window}_avg"] / 1000
    distance = analyzer.anomaly_z_threshold * row[f"{window}_std"] / 1000
    markers = (
        (mean, "Mean", "#2563eb"),
        (mean - distance, "Lower threshold", "#dc2626"),
        (mean + distance, "Upper threshold", "#dc2626"),
        (row["daily"] / 1000, "Current COB", "#16a34a"),
    )
    figure = go.Figure(go.Histogram(x=reference["daily"] / 1000, name="Historical P&L"))
    for value, label, colour in markers:
        figure.add_vline(
            x=value, line_dash="dash", line_color=colour, annotation_text=label
        )
    figure.update_layout(
        template="plotly_white",
        xaxis_title="Daily P&L (mEUR)",
        yaxis_title="Observations",
    )
    st.plotly_chart(figure, use_container_width=True)
    st.caption(
        f"Reference: {row[f'{window}_start']} to {row[f'{window}_end']} (current COB excluded)"
    )


def main():
    st.set_page_config(page_title=TITLE, layout="wide")
    st.title(TITLE)
    demo_root = Path(__file__).resolve().parent / "demo_data"
    default_root = str(demo_root) if demo_root.is_dir() else "./data"
    root = st.sidebar.text_input(
        "Workbook root folder", os.getenv("TREND_ROOT", default_root)
    )
    st.sidebar.caption(f"Calculation logs: {(Path(root) / 'logs').resolve()}")

    recent = latest_daily_cob(root)
    if recent is None:
        st.warning("No Daily COB found. Check workbook root.")
        return

    report_column, indicator_column, count_column, date_column = st.columns(
        [3, 3, 1.4, 1.8]
    )
    cob = pd.Timestamp(
        date_column.date_input("COB Date (as-of)", value=pd.Timestamp(recent).date())
    )
    n_cobs = count_column.number_input(
        "Previous COBs", min_value=1, max_value=250, value=10
    )
    result = load_trend(
        root, cob.date().isoformat(), max(5, n_cobs), heatmap_cobs=n_cobs
    )
    history = result["history"]

    report_options = sorted(history["report_name"].unique())
    reports = report_column.multiselect(
        "Reports", report_options, default=report_options
    )
    report_history = history.loc[history["report_name"].isin(reports)]
    indicator_options = sorted(report_history["indicator"].unique())
    indicators = indicator_column.multiselect(
        "Indicators", indicator_options, default=indicator_options
    )
    selected_history = report_history.loc[report_history["indicator"].isin(indicators)]
    if selected_history.empty:
        st.info("Select at least one report and indicator.")
        return

    st.caption(f"Latest Daily COB: {recent} | P&L values: mEUR")
    window_column = st.columns(3)[0]
    window = window_column.selectbox(
        "Trend window", tuple(WINDOW_LABELS), format_func=WINDOW_LABELS.get
    )
    dates = pd.DatetimeIndex(history["cob_date"].drop_duplicates().nlargest(n_cobs))
    analyzer = ReportTrendAnalyzer(root)
    stats = analyzer.get_heatmap_statistics(selected_history, dates, window)
    threshold = analyzer.anomaly_z_threshold

    heatmap, anomaly, lines, export, distribution = st.tabs(
        [
            "Heatmap",
            "Anomaly Overview",
            "Line Graphs",
            "Save File",
            "Threshold Analysis",
        ]
    )
    with heatmap:
        st.caption(
            "Cells show daily P&L in mEUR and the signed z-score. Colour shows the magnitude of the z-score."
        )
        if stats.empty:
            st.info(
                "No P&L observations for the selected reports and indicators in these COBs."
            )
        else:
            st.plotly_chart(
                generate_zscore_heatmap(stats, dates),
                use_container_width=True,
            )
    with anomaly:
        show_anomaly_overview(stats, cob, window, threshold)
    with lines:
        # Keep the dashboard-wide COB range even if a selected pair has gaps.
        pivot = history_to_pivot(history, cob)
        pivot = pivot.loc[
            pivot["report_name"].isin(reports) & pivot["indicator"].isin(indicators)
        ]
        show_line_graphs(pivot, n_cobs)
    with export:
        show_excel_export(root, cob)
    with distribution:
        show_threshold_analysis(analyzer, result["trend"], selected_history, cob)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        st.error(f"Unable to load dashboard: {exc}")
