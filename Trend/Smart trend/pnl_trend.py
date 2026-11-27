"""Analyze daily P&L extracts ending on a COB date (YYYY-MM-DD)."""
from report_paths import WINDOWS, report_path

import argparse
from openpyxl import load_workbook
import json
import re
import sqlite3
from contextlib import closing, ExitStack
from itertools import groupby
from operator import itemgetter
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tempfile import NamedTemporaryFile, TemporaryDirectory


def parse_day(value):
    if isinstance(value, datetime):
        return value.date()
    if type(value) is date:
        return value
    if not isinstance(value, str) or not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", value):
        raise ValueError("COB date must use YYYY-MM-DD.")
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError("COB date must be a valid calendar date in YYYY-MM-DD format.") from exc


class PnlTrendAnalyzer:
    keys = ("COL1", "COL2", "COL3", "COL4", "REGION", "COL5")

    def __init__(self, source, minimum=10, z_threshold=3, history_days=6, min_observations=4):
        self.source = Path(source)
        if not isinstance(history_days, int) or history_days < 4:
            raise ValueError("History window must contain at least four calendar days.")
        if not isinstance(min_observations, int) or not 2 <= min_observations <= history_days - 2:
            raise ValueError("Minimum observations must be between 2 and history_days minus 2.")
        self.history_days = history_days
        self.min_observations = min_observations
        minimum = Decimal(str(minimum))
        if not minimum.is_finite() or minimum <= 0:
            raise ValueError("Minimum percentage change must be positive.")
        self.minimum = minimum
        self.z_threshold = Decimal(str(z_threshold))
        if not self.z_threshold.is_finite() or self.z_threshold <= 0:
            raise ValueError("Z threshold must be finite and positive.")

    def _generate_file_name(self, day):
        return f"day_{day}.xlsx"

    def input_files(self, days):
        if self.source.is_file():
            return [(self.source, None)]  # Consolidated input carries a Day column.
        files = [(self.source / self._generate_file_name(d), d) for d in days]
        for path, _ in files:
            if not path.is_file():
                raise FileNotFoundError(f"Missing daily extract: {path}")
        return files

    def _sheet_rows(self, path, required):
        if path.suffix.lower() != ".xlsx":
            raise ValueError(f"{path}: expected an XLSX workbook.")
        with closing(load_workbook(path, read_only=True, data_only=True, keep_links=False)) as book:
            sheet = book["PnL"] if "PnL" in book.sheetnames else book.worksheets[0]
            sheet.reset_dimensions()
            rows = sheet.iter_rows(values_only=True)
            headers = tuple(str(h).strip() if h is not None else "" for h in next(rows, ()))
            missing = required.difference(headers)
            if missing:
                raise ValueError(f"{path} [{sheet.title}]: missing columns {sorted(missing)}")
            if any(headers.count(h) > 1 for h in required | {"Day"}):
                raise ValueError(f"{path} [{sheet.title}]: duplicate input column")
            for number, values in enumerate(rows, 2):
                if any(v is not None for v in values):
                    yield number, dict(zip(headers, values))

    def available_days(self):
        if self.source.is_file():
            with closing(self._sheet_rows(self.source, {"Day"})) as rows:
                return sorted({parse_day(row.get("Day")) for _, row in rows})
        days = [parse_day(p.stem[4:]) for p in self.source.glob("day_*.xlsx")
                if re.fullmatch(r"day_[0-9]{4}-[0-9]{2}-[0-9]{2}", p.stem)]
        if not days:
            raise ValueError("Expected daily extracts named day_YYYY-MM-DD.xlsx.")
        return sorted(days)

    def latest_day(self):
        return max(self.available_days())

    def _rows(self, days):
        index = {d: i for i, d in enumerate(days)}
        seen = set()
        for path, file_day in self.input_files(days):
            required = set(self.keys) | {"Daily"}
            if file_day is None:
                required.add("Day")
            with closing(self._sheet_rows(path, required)) as records:
                for number, row in records:
                    try:
                        day = parse_day(row["Day"]) if "Day" in row else file_day
                        if file_day is not None and day != file_day:
                            raise ValueError("Day differs from filename")
                        if day not in index:
                            continue
                        key = tuple(str(row[k]).strip() if row.get(k) is not None else "" for k in self.keys)
                        value = Decimal(str(row.get("Daily")))
                        if not all(key) or not value.is_finite():
                            raise ValueError("Missing hierarchy or non-finite value")
                    except (ValueError, TypeError, AttributeError, InvalidOperation) as exc:
                        raise ValueError(f"{path}, worksheet row {number}: invalid input") from exc
                    a, b, c, d, region, leaf = key
                    yield a, b, region, c, d, leaf, index[day], str(value)
                    seen.add(day)
        if seen != set(days):
            raise ValueError(f"No rows for reporting days {sorted(set(days) - seen)}")

    def _describe(self, path, history, minimum_count):
        def percent(before, after):
            if before is None or before == 0 or after is None:
                return None
            return 100 * (after - before) / abs(before)

        prior = [pct for before, after in zip(history[:-2], history[1:-1])
                 if (pct := percent(before, after)) is not None]
        before, after = history[-2:]
        pct = percent(before, after)
        # Contribution to the reported balance includes entries and exits.
        change = (after or 0) - (before or 0)
        mean = sum(prior) / len(prior) if prior else None
        std = None
        if len(prior) >= 2:
            std = (sum((x - mean) ** 2 for x in prior) / (len(prior) - 1)).sqrt()
        z = None
        unusual = False
        status = "scored"
        if after is None:
            status = "missing_reporting_day"
        elif before is None:
            status = "missing_comparison_day"
        elif before == 0:
            status = "zero_comparison_balance"
        elif len(prior) < minimum_count:
            status = "insufficient_history"
        else:
            deviation = pct - mean
            z = deviation / std if std else None
            exceptional = abs(z) >= self.z_threshold if std else deviation != 0
            unusual = abs(pct) >= self.minimum and exceptional
            if not std:
                status = "flat_baseline"
        alerts = []
        if unusual:
            alerts.append("unusual_increase" if pct > 0 else "unusual_decrease")
        return dict(path=path, daily_history=history, change=change, change_pct=pct,
                    observed_days=sum(x is not None for x in history),
                    baseline_observation_count=len(prior),
                    baseline=dict(mean_change_pct=mean, sample_std_change_pct=std),
                    scoring_status=status, z_score=z, is_anomalous=unusual, alerts=alerts)

    def _stage(self, db, days):
        db.execute("PRAGMA temp_store=FILE")
        db.execute("PRAGMA cache_size=-8192")
        db.execute("CREATE TABLE pnl(a,b,region,c,d,leaf,day INTEGER,value TEXT)")
        db.executemany("INSERT INTO pnl VALUES(?,?,?,?,?,?,?,?)", self._rows(days))
        db.execute("CREATE INDEX report_order ON pnl(a,b,region,c,d,leaf,day)")
        db.commit()

    def _write_report(self, days, out, db, offset=0, window=None):
        labels = [d.isoformat() for d in days]
        header = dict(schema_version=3, metric="daily", reporting_day=labels[-1],
                      comparison_day=labels[-2], history_days=labels, max_baseline_observations=len(days)-2, minimum_baseline_observations=self.min_observations,
                      minimum_change_pct=self.minimum, z_threshold=self.z_threshold)
        if window:
            header.update(window=window, first_baseline_cob=labels[1] if len(days)>2 else None,
                          minimum_baseline_observations=WINDOWS[window])
        encode = json.JSONEncoder(default=json_number, allow_nan=False, separators=(",", ":")).encode
        out.write(encode(header)[:-1] + ',"findings":[')
        # Disk sorting makes each subtree contiguous, including unsorted and duplicate input rows.
        rows = db.execute("SELECT a,b,region,c,d,leaf,day-?,value FROM pnl WHERE day>=? "
                          "ORDER BY a,b,region,c,d,leaf,day", (offset, offset))
        for i, (business, group) in enumerate(groupby(rows, itemgetter(0))):
            if i:
                out.write(",")
            self._emit((business,), group, out, encode, len(days), header["minimum_baseline_observations"])
        out.write("]}\n")

    def _emit(self, path, rows, out, encode, length, minimum_count):
        """Write children first so only the current ancestors and two drivers are retained."""
        history = [None] * length
        if len(path) == 6:
            for row in rows:
                history[row[6]] = (history[row[6]] or 0) + Decimal(row[7])
            node = self._describe(path, history, minimum_count)
            out.write(encode(node))
            drivers = [(path, node["change"])] if node["change"] else []
            return history, drivers

        level = len(path) - 1
        child_name = ("branches", "regions", "leaves")[level]
        key = (itemgetter(1), itemgetter(2), itemgetter(3, 4, 5))[level]
        out.write('{"' + child_name + '":[')
        increases = decreases = Decimal(0)
        drivers = []
        for i, (identifier, group) in enumerate(groupby(rows, key)):
            if i:
                out.write(",")
            if level == 2:
                c, d, leaf = identifier
                child_path = (*path[:2], c, d, path[2], leaf)
            else:
                child_path = (*path, identifier)
            child, candidates = self._emit(child_path, group, out, encode, length, minimum_count)
            history = [a if b is None else (a or 0) + b for a, b in zip(history, child)]
            change = (child[-1] or 0) - (child[-2] or 0)
            increases += max(0, change)
            decreases += max(0, -change)
            if level and candidates:
                drivers = sorted(drivers + candidates, key=lambda x: (-abs(x[1]), x[0]))[:2]
        node = self._describe(path, history, minimum_count)
        node["gross_child_change"] = increases + decreases
        if history[-2] and 100 * min(increases, decreases) >= self.minimum * abs(history[-2]):
            name = ("branch", "region", "leaf")[level]
            node["alerts"].append(f"offsetting_{name}_moves")
        if level == 1:
            node["top_leaf_drivers"] = [dict(path=p, change=c) for p, c in drivers]
            node["other_leaf_change"] = node["change"] - sum(c for _, c in drivers)
        if level == 0:
            node["other_branch_change"] = node["change"] - (increases - decreases)
        out.write("]," + encode(node)[1:])
        return history, drivers

    def generate_windows(self, cob_date):
        """Stage the shared set of input COBs once, then stream five reports."""
        end = parse_day(cob_date)
        dates = [d for d in self.available_days() if d <= end]
        if len(dates) < 2 or dates[-1] != end:
            raise ValueError("Supply the selected COB and at least one earlier COB.")
        windows = {}
        for name, count in WINDOWS.items():
            if name in ("Month", "Quarter"):
                month = end.month if name == "Month" else 1 + 3 * ((end.month-1)//3)
                start = end.replace(month=month, day=1)
                count = sum(d >= start for d in dates[:-1])
            # One opening balance, the baseline changes, and the current observation.
            windows[name] = dates[-(count + 2):]
        days = max(windows.values(), key=len)
        targets = {name: report_path(end, name) for name in windows}
        report_path(end).parent.mkdir(parents=True, exist_ok=True)
        with TemporaryDirectory() as scratch, closing(sqlite3.connect(Path(scratch)/"sort.db")) as db:
            self._stage(db, days)
            with ExitStack() as cleanup:
                staged = {}
                for name, selected in windows.items():
                    with NamedTemporaryFile(mode="w", encoding="utf-8", dir=targets[name].parent,
                                            delete=False) as out:
                        temp = Path(out.name)
                        cleanup.callback(temp.unlink, missing_ok=True)
                        self._write_report(selected, out, db, len(days)-len(selected), name)
                        staged[name] = temp
                for name, temp in staged.items():
                    temp.replace(targets[name])
        return targets

    def trigger(self, cob_date, output=None):
        """Pass an output path for bounded-memory reporting, or omit it to return a dict."""
        end = parse_day(cob_date)
        days = [end - timedelta(days=i) for i in range(self.history_days - 1, -1, -1)]
        with TemporaryDirectory() as scratch:
            target = Path(output) if output is not None else Path(scratch) / "report.json"
            if target.resolve() == self.source.resolve() or (
                    self.source.is_dir() and target.resolve().parent == self.source.resolve()
                    and target.suffix.lower() == ".xlsx"):
                raise ValueError("Output would overwrite an input workbook.")
            with NamedTemporaryFile(mode="w", encoding="utf-8", dir=target.parent,
                                    delete=False) as out:
                temp = Path(out.name)
            try:
                with temp.open("w", encoding="utf-8") as out:
                    with closing(sqlite3.connect(Path(scratch) / "sort.db")) as db:
                        self._stage(db, days)
                        self._write_report(days, out, db)
                temp.replace(target)
            finally:
                temp.unlink(missing_ok=True)
            if output is None:
                with target.open(encoding="utf-8") as stream:
                    return json.load(stream, parse_float=Decimal)
            return target


def analyse(source, minimum=10):
    analyzer = PnlTrendAnalyzer(source, minimum)
    return analyzer.trigger(analyzer.latest_day())


def json_number(value):
    return int(value) if value == value.to_integral_value() else float(value)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path, nargs="?", help="Defaults to trend_YYYY-MM-DD.json under trend_params/YYYY/MM/")
    parser.add_argument("--minimum-pct", "--minimum", type=Decimal, default=10, help="Minimum absolute percentage move (10 means 10%%)")
    parser.add_argument("--history-days", type=int, default=6, help="Calendar days including the COB date")
    parser.add_argument("--min-observations", type=int, default=4, help="Minimum valid prior daily percentage changes for scoring")
    parser.add_argument("--cob-date", metavar="YYYY-MM-DD", help="End of the reporting window")
    parser.add_argument("--z-threshold", type=float, default=3)
    parser.add_argument("--windows", action="store_true", help="Generate Short, Long, Broad, Month and Quarter reports")
    args = parser.parse_args()
    try:
        analyzer = PnlTrendAnalyzer(args.source, args.minimum_pct, args.z_threshold, args.history_days, args.min_observations)
        end = analyzer.latest_day() if args.cob_date is None else parse_day(args.cob_date)
        if args.windows:
            if args.output:
                parser.error("--windows supplies its own five output filenames")
            analyzer.generate_windows(end)
        else:
            target = args.output or report_path(end)
            if args.output is None:
                target.parent.mkdir(parents=True, exist_ok=True)
            analyzer.trigger(end, target)
    except (ValueError, FileNotFoundError) as exc:
        parser.error(str(exc))
