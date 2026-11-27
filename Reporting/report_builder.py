import json
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory

import numpy as np

from ReportingEngine import generate_report, load_config
from Reporting_4 import ExcelReportGenerator


def _json_scalar(value):
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, date):
        return value.isoformat()
    raise TypeError(f"Unsupported report value: {type(value).__name__}")


def _write_json(path, data):
    with path.open("w", encoding="utf-8") as stream:
        json.dump(data, stream, default=_json_scalar, ensure_ascii=False,
                  allow_nan=False, indent=2)


class ReportBuilder:
    def __init__(self, project_root=None):
        self.root = Path(project_root or Path(__file__).parent).resolve()

    @staticmethod
    def _section_data(dataframe, section):
        rows = dataframe
        for rule in section.get("filters", []):
            column = rule["column"]
            operator = rule["operator"]
            if column not in rows.columns:
                raise ValueError(f"{section['id']}: missing filter column {column!r}")
            if operator == "not_null":
                mask = rows[column].notna()
            elif operator in ("in", "not_in") and isinstance(rule.get("values"), list):
                mask = rows[column].isin(rule["values"])
                if operator == "not_in":
                    mask = ~mask
            else:
                raise ValueError(f"{section['id']}: invalid filter for {column!r}: {rule}")
            rows = rows.loc[mask]
        return rows

    def build_report(self, report_id, report_date, version, dataframe):
        """Build one dated report from the supplied DataFrame and return its index path."""
        if version not in ("Daily", "V1", "V2"):
            raise ValueError("version must be Daily, V1, or V2")
        if isinstance(report_date, str):
            report_date = date.fromisoformat(report_date)
        if isinstance(report_date, datetime):
            report_date = report_date.date()
        definition = load_config(self.root / "report_config" / report_id / "report.yaml")
        context = dict(definition["versions"][version], report_id=report_id, date=report_date)
        output = definition["output"]
        directory = self.root / output["directory"].format(**context)
        config_directory = definition["config_directory"].format(**context)
        configs = self.root / config_directory
        mapping_path = self.root / definition["mapping"].format(config_directory=config_directory, **context)
        mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
        sections = definition["sections"]
        defaults = definition.get("excel", {}).get("defaults", {})
        workbook_name = output["workbook"].format(**context)
        index_name = output["web_index"].format(**context)
        names = [section.get("json_filename", output["section_json"]).format(
            section_id=section["id"], **context) for section in sections]
        files = [*names, workbook_name, index_name]
        if not sections or len({s["id"] for s in sections}) != len(sections):
            raise ValueError("Configure at least one section and use distinct section IDs")
        if (len({name.casefold() for name in files}) != len(files)
                or any(Path(name).name != name or name in ("", "..") for name in files)):
            raise ValueError("Each output must have a distinct filename within the report directory")

        first_sheet = (defaults | sections[0]["excel"]).get("sheet_name", "Global")
        generator = ExcelReportGenerator(sheet_name=first_sheet)
        next_rows = {}
        index = dict(report_id=report_id, title=definition["title"], version=version,
                     date=report_date.isoformat(), workbook=workbook_name, sections=[])
        directory.mkdir(parents=True, exist_ok=True)
        with TemporaryDirectory(dir=directory) as temporary:
            stage = Path(temporary)
            for section, name in zip(sections, names):
                rows = self._section_data(dataframe, section)
                tree = generate_report(rows, configs / section["config"])
                _write_json(stage / name, tree)
                options = defaults | section["excel"]
                sheet = options.setdefault("sheet_name", "Global")
                options.setdefault("start_row", next_rows.get(sheet, 3))
                heading = options.pop("merged_heading", None)
                if heading:
                    heading = {"row_number": options["start_row"], **heading}
                    row = heading["row_number"]
                    if type(row) is not int or row < 1:
                        raise ValueError(f"{section['id']}: merged_heading.row_number must be a positive integer")
                    if row < next_rows.get(sheet, 1):
                        raise ValueError(f"{section['id']}: heading row {row} precedes the next available row on {sheet}")
                    options["start_row"] = max(options["start_row"], row + 1)
                next_rows[sheet] = generator.create_report(
                    tree, mapping, report_title=section["title"], **options)
                if heading:
                    generator.write_merged_heading(**heading)
                index["sections"].append(dict(id=section["id"], title=section["title"],
                                              file=name, columns=list(mapping)))
            index["mapping"] = mapping
            generator.save(stage, workbook_name)
            _write_json(stage / index_name, index)
            self._publish(stage, directory, files)
        return directory / index_name

    @staticmethod
    def _publish(stage, directory, files):
        # Serialize replacements for this edition and restore prior files on failure.
        from shutil import copy2

        lock = directory / ("." + files[-1] + ".lock")
        lock.touch(exist_ok=False)
        try:
            with TemporaryDirectory(dir=directory) as backup:
                replaced = []
                try:
                    for name in files:
                        target = directory / name
                        if target.exists():
                            copy2(target, Path(backup) / name)
                        (stage / name).replace(target)
                        replaced.append(name)
                except OSError:
                    for name in reversed(replaced):
                        previous = Path(backup) / name
                        if previous.exists():
                            previous.replace(directory / name)
                        else:
                            (directory / name).unlink()
                    raise
        finally:
            lock.unlink()
