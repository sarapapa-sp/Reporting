# Reporting Trends

Run these commands from the extracted `Trends/` folder.

```sh
python -m pip install -r requirements.txt
python -m streamlit run Trend_dashboard.py
```

The dashboard opens with the included synthetic data. Set **Workbook root
folder** in the sidebar or set the `TREND_ROOT` environment variable to use
your own workbooks.

```sh
python run_trend.py --root demo_data --version Daily --cob-date 2026-11-25 --output demo_trend.xlsx
```

## Try the demo

For **25 November 2026**, select **Short (5 COBs)**. Expect these results:

| Report | Indicator | Daily P&L (mEUR) | Z-score | Status |
| --- | --- | ---: | ---: | --- |
| Credit | Net P&L | +0.25 | — | Insufficient history |
| Equities | Net P&L | −2.50 | −2.64 | Flagged |
| FX | Trading P&L | +0.10 | +0.25 | Within range |
| Rates | Net P&L | +3.00 | +3.16 | Flagged |

The data is synthetic. To regenerate it, save and run the generator from the
companion HTML guide in this folder.

## Workbook format

Put monthly workbooks in `<root>/YYYY/Flat_Conso_Reports_Mon.xlsx`.
The first sheet needs `cob_date`, `version`, `report_name`, `indicator`,
`daily` in kEUR, and `bus_count` for the previous-month comparison. Each
COB, report, and indicator key must be unique within a version.

The calculation excludes the COB under evaluation from its own baseline.
Previous-month dates are aligned by business-day position, with adjustments
near the middle when the months have different lengths. The calendar uses
weekdays and the dates in `ReportTrendAnalyzer.EXCEPTION_COBS`. For production,
check those dates against the official reporting calendar.

## Source files

| File | Role |
| --- | --- |
| `Publication_Trendv2.py` | Excel loading, scoring, comparisons and audit logs |
| `Heatmap_Trend.py` | Plotly charts and daily P&L pivot |
| `Trend_dashboard.py` | Streamlit controls, five tabs and downloads |
| `export_utils.py` | Excel sheets, units and formatting |
| `run_trend.py` | Command-line entry point |
| `requirements.txt` | Python dependencies |

Both entry points use the same analyzer. The dashboard heatmap uses z-score
colours and the Excel heatmap uses P&L magnitude. The Excel workbook includes
all six trend windows. The export period option is an audit label. Each COB
writes one audit file under `<root>/logs/`, when that file is first created.
