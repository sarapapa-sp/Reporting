"""Browse trend.json with: streamlit run pnl_viewer.py"""
from datetime import date
from report_paths import REPORT_DIR, WINDOWS, report_path

import json
import sqlite3
import tempfile
from contextlib import closing
from pathlib import Path

import ijson
from ijson.common import ObjectBuilder
import streamlit as st

ARRAYS = ("findings", "branches", "regions", "leaves")


@st.cache_resource(show_spinner="Indexing the report…", max_entries=2)
def index_report(filename, modified, size):
    # An index per report revision. Child arrays stream directly to separate records.
    temp = tempfile.TemporaryDirectory()
    with closing(sqlite3.connect(Path(temp.name) / "report.db")) as db, open(filename, "rb") as src:
        db.execute("CREATE TABLE nodes(id INTEGER PRIMARY KEY,root TEXT,branch TEXT,region TEXT,depth INTEGER,impact INTEGER,payload TEXT)")
        node_prefixes = [".".join(k + ".item" for k in ARRAYS[:i]) for i in range(1, 5)]
        arrays = {p[:-5] for p in node_prefixes}
        header = ObjectBuilder()
        stack = []
        seq = 0
        for prefix, event, value in ijson.parse(src, use_float=True):
            if prefix in arrays or (event == "map_key" and value in ARRAYS):
                continue
            if prefix in node_prefixes and event == "start_map":
                seq += 1
                stack.append([seq, ObjectBuilder(), False])
            builder = stack[-1][1] if stack else header
            builder.event(event, value)
            if prefix in node_prefixes and event == "end_map":
                ident, record, child_impact = stack.pop()
                node = record.value
                impact = bool(node["alerts"] or node["is_anomalous"] or child_impact)
                if impact and stack:
                    stack[-1][2] = True
                path = node["path"]
                branch = path[1] if len(path) > 1 else None
                region = path[4] if len(path) == 6 else path[2] if len(path) == 3 else None
                db.execute("INSERT INTO nodes VALUES(?,?,?,?,?,?,?)", (
                    ident, path[0], branch, region, len(stack)+1, impact, json.dumps(node)))
        db.execute("CREATE INDEX impact_rows ON nodes(root,depth,id) WHERE impact=1")
        db.execute("CREATE INDEX hierarchy ON nodes(depth,root,branch,region)")
        db.commit()
    return temp, header.value


def fmt(value, signed=False):
    if value is None:
        return "—"
    return format(value, "+,.2f" if signed else ",.2f").rstrip("0").rstrip(".")


st.set_page_config(page_title="P&L movement explorer", layout="wide")
st.markdown("### P&L movement explorer")
available = []
for file in REPORT_DIR.glob("????/??/trend_????-??-??*.json"):
    try:
        file_day = date.fromisoformat(file.stem[6:16])
        if file in {report_path(file_day, w) for w in (None, *WINDOWS)}:
            available.append(file_day)
    except ValueError:
        continue
cob = st.date_input("COB date", value=max(available, default=date.today()), format="YYYY-MM-DD")
choices = [name for name in WINDOWS if report_path(cob, name).is_file()]
if report_path(cob).is_file():
    choices.append("Custom")
window = st.selectbox("Baseline window", choices) if choices else None
source = report_path(cob, window if window != "Custom" else None)
if not source.exists():
    st.info(f"No report for {cob}. Expected {source.relative_to(REPORT_DIR.parent)}.")
    st.stop()
stat = source.stat()
cache, report = index_report(str(source), stat.st_mtime_ns, stat.st_size)
if report["reporting_day"] != cob.isoformat():
    st.error(f"The reporting day in {source.name} differs from the selected date.")
    st.stop()
days = report["history_days"]
st.caption(f"COB {days[-1]}   ·   Compared with {days[-2]}   ·   {len(days)} COB balances")
if report.get("window"):
    st.caption(f"{window} baseline · First change COB {report['first_baseline_cob'] or '—'} · "
               f"Up to {report['max_baseline_observations']} prior changes · "
               f"Minimum {report['minimum_baseline_observations']} to score")
with closing(sqlite3.connect(Path(cache.name) / "report.db")) as db:
    roots = dict(db.execute("SELECT root,payload FROM nodes WHERE depth=1 ORDER BY root"))
    levels = ["COL1", "COL2", "Region", "COL5"]
    left, branch_control, region_control, right = st.columns(4)
    selected = left.multiselect("COL1", list(roots), default=list(roots))
    level = right.selectbox("Row level", levels)
    depth_selected = levels.index(level) + 1
    if not selected:
        st.info("Select a COL1 value to see movements.")
        st.stop()
    conditions = ["root IN (SELECT value FROM json_each(?))"]
    filters = [json.dumps(selected)]
    for depth, field, label, control in ((2, "branch", "COL2", branch_control),
                                         (3, "region", "Region", region_control)):
        options = [v for v, in db.execute(
            f"SELECT DISTINCT {field} FROM nodes WHERE depth=? AND " + " AND ".join(conditions)
            + f" ORDER BY {field}", (depth, *filters))]
        active = depth_selected >= depth
        chosen = control.multiselect(label, options, default=options, disabled=not active,
                                     help=f"Applies at {label} row level and below.")
        if active:
            conditions.append(f"{field} IN (SELECT value FROM json_each(?))")
            filters.append(json.dumps(chosen))
    scope = " AND ".join(conditions)
    # Aggregate at the finest active filter grain, including stable rows.
    summary_depth = min(3, depth_selected)
    summaries = [json.loads(payload) for payload, in db.execute(
        "SELECT payload FROM nodes WHERE depth=? AND " + scope, (summary_depth, *filters))]
    if not summaries:
        st.info("No matching hierarchy values are selected.")
        st.stop()
    combined = [sum(x for x in values if x is not None) if any(x is not None for x in values) else None
                for values in zip(*(n["daily_history"] for n in summaries))]
    movement = sum(n["change"] for n in summaries)
    pct = 100 * movement / abs(combined[-2]) if combined[-2] and combined[-1] is not None else None
    cards = st.columns(4)
    for card, label, value in zip(cards, ("Daily P&L", "Net movement", "Change (%)", ("Gross branch movement", "Gross region movement", "Gross item movement")[summary_depth-1]),
                                  (fmt(combined[-1]), fmt(movement, True), fmt(pct, True),
                                   fmt(sum(n["gross_child_change"] for n in summaries)))):
        card.metric(label, value)
    st.subheader("Impacting movements")
    st.caption("Each row represents the selected level. Summary metrics cover all movements matching the active filters.")
    where = " FROM nodes WHERE impact=1 AND depth=? AND " + scope
    params = (depth_selected, *filters)
    total = db.execute("SELECT count(*)" + where, params).fetchone()[0]
    if not total:
        st.info("No impacting movements for this selection.")
        st.stop()
    page_size = 200
    page_count = (total + page_size - 1) // page_size
    page = 1
    if page_count > 1:
        page = st.number_input(f"Page of {page_count:,}", min_value=1, max_value=page_count,
                               key=f"page-{filters}-{level}-{cob}-{stat.st_mtime_ns}")
    first = (page - 1) * page_size
    rows = []
    for (payload,) in db.execute("SELECT payload" + where + " ORDER BY id LIMIT ? OFFSET ?",
                                 (*params, page_size, first)):
        n = json.loads(payload)
        path = n["path"]
        depth = len(path)
        rows.append({
            "COL1": path[0],
            "COL2": path[1] if depth >= 2 else "",
            "Region": path[4] if depth == 6 else path[2] if depth == 3 else "",
            "COL5": path[-1] if depth == 6 else "",
            "Prior": n["daily_history"][-2], "Current": n["daily_history"][-1],
            "Move": n["change"], "Change %": n["change_pct"],
            "History": n["daily_history"],
            "Finding": ", ".join(a.replace("_", " ") for a in n["alerts"]) or ("Unusual move" if n["is_anomalous"] else "Contains flagged movement"),
            "Z-score": n["z_score"], "Prior changes": n["baseline_observation_count"],
            "Mean %": n["baseline"]["mean_change_pct"], "Std dev (pp)": n["baseline"]["sample_std_change_pct"],
            "Observed days": n["observed_days"], "Source path": " / ".join(path),
        })
        for column in levels[depth_selected:]:
            del rows[-1][column]
    st.dataframe(rows, hide_index=True, width="stretch", height=min(640, 38 + len(rows)*35),
                 column_config={"History": st.column_config.LineChartColumn("History"),
                                **{k: st.column_config.NumberColumn(format="%.2f")
                                   for k in ("Change %", "Z-score", "Mean %", "Std dev (pp)")}})
    if page_count > 1:
        st.caption(f"Rows {first+1:,}–{first+len(rows):,} of {total:,}")
    st.subheader("Filtered P&L history")
    st.line_chart([{"Date": d, "Daily P&L": v} for d,v in zip(days, combined)],
                  x="Date", y="Daily P&L", height=220)
