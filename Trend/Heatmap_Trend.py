
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go


def generate_heatmap(
    df: pd.DataFrame,
    report_name: str = "All Reports",
    n_cobs: int = None,
    title: str = "Global Markets P&L Heatmap"
) -> go.Figure:
    """
    Generate a Plotly heatmap from a pre-pivoted DataFrame.

    Expected structure:
        report_name | indicator | cob_date1 | cob_date2 | ...

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame containing report, indicator and P&L values.

    report_name : str
        Report to filter, or "All Reports".

    n_cobs : int, optional
        Number of most recent COB columns to display.

    title : str
        Heatmap title.

    Returns
    -------
    go.Figure
        Plotly heatmap figure.
    """

    required_cols = ["report_name", "indicator"]

    if df.empty:
        raise ValueError("Input DataFrame is empty.")

    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Identify and validate COB date columns
    date_cols = [c for c in df.columns if c not in required_cols]

    if not date_cols:
        raise ValueError("No COB date columns available.")

    parsed_dates = pd.to_datetime(
        [str(c) for c in date_cols],
        errors="coerce"
    )

    if pd.isna(parsed_dates).any():
        raise ValueError("Invalid COB date column names found.")

    if parsed_dates.duplicated().any():
        raise ValueError("Duplicate COB dates found.")

    # Sort dates in ascending order
    sorted_pairs = sorted(
        zip(parsed_dates, date_cols),
        key=lambda x: x[0]
    )

    date_cols = [col for _, col in sorted_pairs]

    # Select last N COB dates
    if n_cobs is not None:
        if not isinstance(n_cobs, int) or n_cobs < 1:
            raise ValueError("n_cobs must be a positive integer.")

        date_cols = date_cols[-n_cobs:]

    data = df.copy()

    # Filter selected report
    if report_name != "All Reports":
        data = data[data["report_name"] == report_name]

    if data.empty:
        raise ValueError("No data for selected report.")

    # Validate unique report / indicator combinations
    if data.duplicated(required_cols).any():
        raise ValueError(
            "Duplicate report_name and indicator combinations."
        )

    # Sort data
    data = data.sort_values(required_cols)

    # Create labels
    data["row_label"] = (
        data["report_name"].astype(str)
        + " | "
        + data["indicator"].astype(str)
    )

    # Numeric P&L values
    raw_values = data[date_cols]

    numeric_values = raw_values.apply(
        pd.to_numeric,
        errors="coerce"
    )

    invalid = numeric_values.isna() & raw_values.notna()

    if invalid.any().any():
        raise ValueError(
            "Non-numeric P&L values found in COB columns."
        )

    z_values = numeric_values.to_numpy(dtype=float)

    if not np.isfinite(z_values).any():
        raise ValueError("No valid P&L values available.")

    z_values[~np.isfinite(z_values)] = np.nan

    # Symmetric red / green color scale
    max_abs = max(
        1.0,
        float(np.nanmax(np.abs(z_values)))
    )

    # Format numbers displayed inside cells
    cell_text = [
        [
            f"{v:,.2f}" if np.isfinite(v) else ""
            for v in row
        ]
        for row in z_values
    ]

    # Format COB labels
    x_labels = [
        pd.to_datetime(str(c)).strftime("%d %b %Y")
        for c in date_cols
    ]

    y_labels = data["row_label"].tolist()

    # Build Plotly heatmap
    fig = go.Figure(
        data=go.Heatmap(
            z=z_values,
            x=x_labels,
            y=y_labels,
            text=cell_text,
            texttemplate="%{text}",
            colorscale=[
                [0.0, "#C0392B"],
                [0.5, "#F5F5F5"],
                [1.0, "#27AE60"]
            ],
            zmin=-max_abs,
            zmax=max_abs,
            xgap=2,
            ygap=2,
            colorbar=dict(
                title="P&L"
            ),
            hovertemplate=(
                "<b>%{y}</b><br>"
                "COB Date: %{x}<br>"
                "P&L: %{z:,.2f}"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(
        title=title,
        xaxis_title="COB Date",
        yaxis_title="Report / Indicator",
        height=max(
            400,
            min(1500, 120 + len(data) * 38)
        ),
        margin=dict(
            l=20, r=20, t=60, b=30
        ),
        template="plotly_white",
        xaxis=dict(
            type="category",
            side="top"
        ),
        yaxis=dict(
            type="category",
            autorange="reversed",
            automargin=True
        )
    )

    return fig


def display_heatmap(df: pd.DataFrame):
    """
    Streamlit interface for displaying the heatmap.
    """

    st.subheader("Global Markets Trend Analysis")

    reports = ["All Reports"] + sorted(
        df["report_name"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    date_cols = [
        c for c in df.columns
        if c not in ["report_name", "indicator"]
    ]

    col1, col2 = st.columns(2)

    with col1:
        selected_report = st.selectbox(
            "Select Report",
            options=reports
        )

    with col2:
        n_cobs = st.number_input(
            "Previous N COBs",
            min_value=1,
            max_value=len(date_cols),
            value=min(5, len(date_cols)),
            step=1
        )

    try:
        fig = generate_heatmap(
            df=df,
            report_name=selected_report,
            n_cobs=int(n_cobs),
            title="P&L Trend Heatmap"
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    except ValueError as error:
        st.error(str(error))


# ------------------------------------
# SAMPLE DATA
# Replace with your Trend class output
# ------------------------------------

if __name__ == "__main__":

    st.set_page_config(layout="wide")

    df = pd.DataFrame({
        "report_name": [
            "Revenue", "Revenue", "Revenue",
            "Costs", "Costs"
        ],
        "indicator": ["A", "B", "C", "D", "E"],
        "2026-11-04": [10, -1, 6, -10, 4],
        "2026-11-05": [14, 2, 8, -4, 7],
        "2026-11-06": [9, -3, -5, 5, -2],
        "2026-11-09": [18, -2, 12, -8, 9]
    })

    display_heatmap(df)



################# Calling
# Assume df_trend is returned by your trend-analysis class

fig = generate_heatmap(
    df=df_trend,
    report_name="All Reports",
    n_cobs=10,
    title="Daily P&L Trend Analysis"
)

st.plotly_chart(fig, use_container_width=True)


