"""Excel export: mEUR, newest COB first, string dates, visible heatmap colours."""

from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MONETARY_COLUMNS = {"daily", "previous_month_daily", "previous_month_change"}
MONETARY_SUFFIXES = ("_avg", "_std", "_min", "_max", "_slope", "_change")


def convert_to_meur(frame):
    out = frame.copy()
    for col in out.columns:
        if (
            col in MONETARY_COLUMNS or col.endswith(MONETARY_SUFFIXES)
        ) and pd.api.types.is_numeric_dtype(out[col]):
            out[col] = out[col] / 1000.0
    return out


def date_as_text(series):
    """Return ISO strings, not Excel date serials or pandas Timestamp objects."""
    return pd.to_datetime(series, errors="coerce").dt.strftime("%Y-%m-%d").fillna("")


def _style_header(ws):
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30


def _heat_colour(value, limit):
    """Seven visible, symmetric red-neutral-green steps, centered on zero."""
    if not isinstance(value, (int, float, np.integer, np.floating)) or not np.isfinite(
        value
    ):
        return None
    scaled = float(value) / limit if limit else 0
    if scaled <= -0.66:
        return "9C1C24"
    if scaled <= -0.33:
        return "D95757"
    if scaled < -0.05:
        return "F5B7B1"
    if scaled <= 0.05:
        return "F5F6F7"
    if scaled < 0.33:
        return "B8E6CC"
    if scaled < 0.66:
        return "59B989"
    return "137547"


def make_excel(result):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        trend = convert_to_meur(result["trend"]).drop(
            columns=["mtd", "ytd"], errors="ignore"
        )
        first = ["cob_date", "report_name", "indicator", "daily"]
        trend = trend[
            [c for c in first if c in trend.columns]
            + [c for c in trend.columns if c not in first]
        ]
        for c in trend.columns:
            if c == "cob_date" or c.endswith(("_cob", "_start", "_end")):
                trend[c] = date_as_text(trend[c])
        trend.to_excel(writer, index=False, sheet_name="Trend")
        ws = writer.sheets["Trend"]
        ws.freeze_panes = "E2"
        ws.auto_filter.ref = ws.dimensions
        _style_header(ws)
        for i, name in enumerate(trend.columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = (
                27 if name in ("report_name", "indicator") else 17
            )
            if name == "cob_date" or name.endswith(("_cob", "_start", "_end")):
                for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                    row[0].number_format = "@"
        if "history" in result and not result["history"].empty:
            from Heatmap_Trend import history_to_pivot

            hist = result["history"]
            pivot = history_to_pivot(hist, pd.to_datetime(hist["cob_date"]).max())
            dates = sorted(
                (c for c in pivot.columns if c not in ("report_name", "indicator")),
                key=pd.Timestamp,
                reverse=True,
            )
            pivot = pivot[["report_name", "indicator"] + dates]
            pivot.to_excel(writer, index=False, sheet_name="Heatmap Data")
            ws = writer.sheets["Heatmap Data"]
            ws.freeze_panes = "C2"
            ws.auto_filter.ref = ws.dimensions
            _style_header(ws)
            ws.column_dimensions["A"].width = 27
            ws.column_dimensions["B"].width = 20
            ws.row_dimensions[1].height = 100
            for i in range(3, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(i)].width = 11
                header = ws.cell(1, i)
                header.value = str(header.value)
                header.number_format = "@"
                header.alignment = Alignment(
                    textRotation=90, horizontal="center", vertical="bottom"
                )
            numeric = (
                pivot[dates].apply(pd.to_numeric, errors="coerce").to_numpy(dtype=float)
            )
            finite = numeric[np.isfinite(numeric)]
            limit = float(np.max(np.abs(finite))) if len(finite) else 1.0
            limit = max(limit, 1e-12)
            for row in ws.iter_rows(min_row=2, min_col=3):
                for cell in row:
                    colour = _heat_colour(cell.value, limit)
                    if colour:
                        cell.fill = PatternFill("solid", fgColor=colour)
                    cell.number_format = "#,##0.00;[Red](#,##0.00);–"
                    cell.alignment = Alignment(horizontal="center")
    return output.getvalue()
