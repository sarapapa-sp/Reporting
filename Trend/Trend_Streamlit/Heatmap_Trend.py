"""Plotly charts and the shared daily P&L pivot."""

import pandas as pd
import plotly.graph_objects as go
from plotly.colors import qualitative


def generate_line_graph(df, report_name, indicators=None, n_cobs=None):
    keys = ["report_name", "indicator"]
    date_cols = sorted(
        (c for c in df.columns if c not in keys),
        key=lambda x: pd.Timestamp(x),
        reverse=True,
    )
    if n_cobs:
        date_cols = date_cols[: int(n_cobs)]
    data = df if report_name == "All Reports" else df[df["report_name"] == report_name]
    if indicators:
        data = data[data["indicator"].isin(indicators)]
    fig = go.Figure()
    for i, (_, r) in enumerate(data.sort_values(keys).iterrows()):
        fig.add_trace(
            go.Scatter(
                x=date_cols,
                y=[r[c] for c in date_cols],
                mode="lines+markers",
                name=f"{r.report_name} | {r.indicator}",
                line=dict(color=qualitative.Plotly[i % len(qualitative.Plotly)]),
            )
        )
    fig.update_layout(
        template="plotly_white",
        xaxis_title="COB date",
        yaxis_title="Daily P&L (mEUR)",
        hovermode="x unified",
    )
    return fig


def history_to_pivot(history, as_of_date):
    df = history[history["cob_date"] <= pd.Timestamp(as_of_date)].copy()
    df["cob_date"] = df["cob_date"].dt.strftime("%Y-%m-%d")
    pivot = (
        df.assign(daily=df["daily"] / 1000)
        .pivot(index=["report_name", "indicator"], columns="cob_date", values="daily")
        .reset_index()
    )
    dates = sorted(
        (c for c in pivot.columns if c not in ("report_name", "indicator")),
        reverse=True,
    )
    return pivot[["report_name", "indicator"] + dates]


def generate_report_line_graph(df, n_cobs=None):
    """One line per report, summing the selected indicators for each COB."""
    keys = ["report_name", "indicator"]
    dates = sorted(
        (c for c in df.columns if c not in keys), key=pd.Timestamp, reverse=True
    )
    if n_cobs:
        dates = dates[: int(n_cobs)]
    dates = list(reversed(dates))  # plot chronological left to right
    grouped = df.groupby("report_name", as_index=False)[dates].sum(min_count=1)
    fig = go.Figure()
    for i, (_, r) in enumerate(grouped.iterrows()):
        fig.add_trace(
            go.Scatter(
                x=dates,
                y=[r[d] for d in dates],
                mode="lines+markers",
                name=str(r["report_name"]),
                line=dict(color=qualitative.Plotly[i % len(qualitative.Plotly)]),
            )
        )
    fig.update_layout(
        title="Daily P&L by Report",
        template="plotly_white",
        xaxis_title="COB date",
        yaxis_title="Daily P&L (mEUR)",
        hovermode="x unified",
    )
    return fig


def generate_zscore_heatmap(stats, dates):
    """Display daily P&L and signed z-score, with color based on score magnitude."""
    pairs = sorted(set(zip(stats["report_name"], stats["indicator"])))
    date_labels = [date.strftime("%d %b %Y") for date in dates]
    row_positions = {pair: i for i, pair in enumerate(pairs)}
    column_positions = {date: i for i, date in enumerate(dates)}
    scores = [[None] * len(dates) for _ in pairs]
    labels = [[""] * len(dates) for _ in pairs]
    tooltips = [[""] * len(dates) for _ in pairs]
    bands = [
        ("Unavailable", "#bfc5cd"),
        ("−1 ≤ z ≤ 1", "#dcf5dc"),
        ("1 < |z| ≤ 2", "#83c98d"),
        ("2 < |z| ≤ 3", "#b77916"),
        ("|z| > 3", "#b91c1c"),
    ]

    def amount(value):
        return f"{value / 1000:,.2f} mEUR" if pd.notna(value) else "—"

    for cell in stats.itertuples(index=False):
        if cell.cob_date not in column_positions:
            continue
        row = row_positions[cell.report_name, cell.indicator]
        column = column_positions[cell.cob_date]
        if pd.notna(cell.zscore):
            magnitude = abs(cell.zscore)
            if magnitude > 3:
                color = 4
            elif magnitude > 2:
                color = 3
            elif magnitude > 1:
                color = 2
            else:
                color = 1
            score_text = f"{cell.zscore:+.2f}"
        elif cell.status == "Flagged":
            color = 0
            score_text = "Undefined (zero standard deviation)"
        else:
            color = 0
            score_text = "—"

        scores[row][column] = color
        label_score = score_text if pd.notna(cell.zscore) else "—"
        labels[row][column] = f"{cell.daily / 1000:+,.2f} mEUR<br>z = {label_score}"
        tooltips[row][column] = (
            f"Daily P&L: {amount(cell.daily)}<br>"
            f"Historical mean: {amount(cell.mean)}<br>"
            f"Standard deviation: {amount(cell.std)}<br>"
            f"Z-score: {score_text}<br>Status: {cell.status}"
        )

    fig = go.Figure(
        go.Heatmap(
            x=date_labels,
            y=[f"{report} | {indicator}" for report, indicator in pairs],
            z=scores,
            text=labels,
            customdata=tooltips,
            texttemplate="%{text}",
            colorscale=[
                [edge / len(bands), shade]
                for index, (_, shade) in enumerate(bands)
                for edge in (index, index + 1)
            ],
            zmin=-0.5,
            zmax=len(bands) - 0.5,
            xgap=3,
            ygap=3,
            hoverongaps=False,
            colorbar=dict(
                title="Z-score",
                tickvals=list(range(len(bands))),
                ticktext=[label for label, _ in bands],
            ),
            hovertemplate="<b>%{y}</b> · %{x}<br>%{customdata}<extra></extra>",
        )
    )
    fig.update_layout(
        title="Global Markets | P&L Heatmap",
        template="plotly_white",
        height=max(420, min(1500, 170 + len(pairs) * 42)),
        xaxis=dict(
            side="top",
            type="category",
            tickmode="array",
            tickvals=date_labels,
            ticktext=[date.strftime("%d %b") for date in dates],
        ),
        yaxis=dict(autorange="reversed", automargin=True),
    )
    return fig
