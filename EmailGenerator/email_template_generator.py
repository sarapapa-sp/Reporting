"""Generate plain corporate HTML emails from configured Excel ranges.

Public workflows: generate_html() and generate_email().
Requires openpyxl, Pillow and Jinja2. No Excel or LibreOffice installed.
"""
from __future__ import annotations

import base64
from datetime import date, datetime
from io import BytesIO
import sys
import tempfile
import uuid
import logging
from pathlib import Path
import re
from typing import Any
from urllib.parse import urlparse
from html import escape

from jinja2 import Environment, FileSystemLoader, select_autoescape
from markupsafe import Markup
from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import range_boundaries
from PIL import Image, ImageDraw, ImageFont


class EmailTemplateGenerator:
    """Single-sheet configuration controls component order for each report."""

    _FIELDS = (
        "ReportID", "Subject", "OpeningLine", "Body", "AdditionalComment",
        "Ending", "EmailLayout"
    )
    _TOKENS = {
        "COBDate", "Version", "OpeningLine", "Body", "Image1", "Image2",
        "Image3", "AdditionalComment", "Link", "Ending"
    }
    _DIMENSION_FIELDS = tuple(f"Image{i}{dim}" for i in range(1, 4) for dim in ("Width", "Height"))
    _TOKEN_RE = re.compile(r"{{\s*([A-Za-z][A-Za-z0-9]*)\s*}}")

    def __init__(self, config_path: str | Path, template_path: str | Path,
                 logger: logging.Logger | None = None):
        self._config_path = Path(config_path).expanduser().resolve()
        self._template_path = Path(template_path).expanduser().resolve()
        self._logger = logger or logging.getLogger(__name__)
        if not self._config_path.is_file():
            raise FileNotFoundError(f"Configuration workbook not found: {self._config_path}")
        if not self._template_path.is_file():
            raise FileNotFoundError(f"HTML template not found: {self._template_path}")

    def generate_html(
        self,
        filepath: str | Path,
        cobdate: str | date | datetime,
        version: str,
        reportid: str,
        imagedict: dict[str, dict[str, str]],
        link_url: str,
        link_text: str,
        output_path: str | Path,
    ) -> tuple[str, str]:
        """Build and save HTML, returning (html_string, subject_string).

        Image keys Image1, Image2, Image3 are optional. Each entry has
        Sheetname, Imagerange and Imageinfo. No temporary PNGs are written.
        """
        source = Path(filepath).expanduser().resolve()
        destination = Path(output_path).expanduser().resolve()
        if not source.is_file():
            raise FileNotFoundError(f"Report workbook not found: {source}")
        if source.suffix.lower() not in (".xlsx", ".xlsm"):
            raise ValueError("Source workbook must be .xlsx or .xlsm")
        if destination.suffix.lower() != ".html":
            raise ValueError("output_path must end with .html")
        if not str(reportid).strip():
            raise ValueError("reportid cannot be empty")
        version = str(version).strip().title()
        if version not in ("Daily", "Flash", "Final"):
            raise ValueError("version must be Daily, Flash, or Final")
        formatted_date = self._format_date(cobdate)
        self._validate_image_dict(imagedict)
        link_url = str(link_url or "").strip()
        link_text = str(link_text or "").strip()
        if link_url:
            self._validate_link(link_url)
        elif link_text:
            raise ValueError("link_text cannot be specified without link_url")

        config = self._load_configuration(str(reportid).strip())
        substitutions = {"COBDate": formatted_date, "Version": version}
        # Do not replace layout's structural placeholders prematurely.
        fields = {k: self._replace_tokens(v, substitutions)
                  for k, v in config.items() if k != "EmailLayout"}
        layout = config["EmailLayout"]
        used_tokens = set(self._TOKEN_RE.findall(layout))
        unknown = used_tokens - self._TOKENS
        if unknown:
            raise ValueError(f"Unrecognized EmailLayout placeholders: {sorted(unknown)}")

        image_blocks: dict[str, str] = {}
        needed = [key for key in ("Image1", "Image2", "Image3")
                  if key in imagedict and key in used_tokens]
        if needed:
            workbook = load_workbook(source, read_only=False, data_only=True)
            try:
                for key in needed:
                    details = imagedict[key]
                    if details["Sheetname"] not in workbook.sheetnames:
                        raise ValueError(f"Worksheet {details['Sheetname']!r} not found for {key}")
                    png = self._render_range_to_png(
                        workbook[details["Sheetname"]], details["Imagerange"]
                    )
                    data_uri = "data:image/png;base64," + base64.b64encode(png).decode("ascii")
                    caption = self._replace_tokens(str(details["Imageinfo"]), substitutions)
                    image_blocks[key] = self._image_html(
                        data_uri, caption,
                        width=config.get(f"{key}Width"),
                        height=config.get(f"{key}Height"),
                    )
            finally:
                workbook.close()

        replacements: dict[str, str] = {
            "COBDate": escape(formatted_date),
            "Version": escape(version),
            "OpeningLine": self._escaped_multiline(fields["OpeningLine"]),
            "Body": self._escaped_multiline(fields["Body"]),
            "AdditionalComment": self._escaped_multiline(fields["AdditionalComment"]),
            "Ending": self._escaped_multiline(fields["Ending"]),
            "Link": (f'<a href="{escape(link_url, quote=True)}" '
                     f'style="color:#0563c1;text-decoration:underline;">'
                     f'{escape(link_text or link_url)}</a>') if link_url else "",
            **{key: image_blocks.get(key, "") for key in ("Image1", "Image2", "Image3")},
        }
        content = self._render_layout(layout, replacements)
        env = Environment(
            loader=FileSystemLoader(str(self._template_path.parent)),
            autoescape=select_autoescape(["html", "xml"]),
        )
        html = env.get_template(self._template_path.name).render(
            Subject=fields["Subject"], Content=Markup(content)
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(html, encoding="utf-8")
        self._logger.info("Generated %s (%s %s), images=%d to %s",
                          reportid, formatted_date, version, len(image_blocks), destination)
        return html, fields["Subject"]

    def generate_email(
        self,
        html: str,
        subject: str,
        output_path: str | Path,
    ) -> str:
        """Create an Outlook .oft template from generated HTML and its subject.

        Windows and Outlook desktop are required for the .oft backend.
        This function never sends an email.
        """
        if not isinstance(html, str) or not html.strip():
            raise ValueError("html must be a nonempty HTML string")
        if not isinstance(subject, str) or not subject.strip():
            raise ValueError("subject must be a nonempty string")
        target = Path(output_path).expanduser().resolve()
        if target.suffix.lower() != ".oft":
            raise ValueError("output_path must end with .oft")
        return self._generate_oft(html, subject, target)

    def _generate_oft(self, html: str, subject: str, destination: Path) -> str:
        """Windows Outlook COM backend; insert PNGs as CID inline attachments.

        Outlook's Attachments.Add API requires a file, so the .oft-specific
        packaging stage uses a temporary directory and removes it afterwards.
        The HTML generation stage always keeps PNGs in memory.
        """
        if sys.platform != "win32":
            raise OSError("Outlook .oft generation requires Windows and Outlook desktop")
        try:
            import win32com.client as win32  # pip install pywin32 on Windows
        except ImportError as exc:
            raise RuntimeError("Install pywin32 on Windows to generate .oft files") from exc

        destination.parent.mkdir(parents=True, exist_ok=True)
        # Replace image data URI with CID references before Outlook stores mail.
        image_tag_pattern = re.compile(r"<img\b[^>]*>", re.IGNORECASE | re.DOTALL)
        data_uri_pattern = re.compile(
            r"\bsrc\s*=\s*([\"\'])(data:image/png;base64,([A-Za-z0-9+/=\s]+))\1",
            re.IGNORECASE,
        )
        attachments: list[tuple[str, bytes]] = []

        def replace_img_tag(match: re.Match[str]) -> str:
            tag = match.group(0)
            uri = data_uri_pattern.search(tag)
            if uri is None:
                return tag
            encoded = re.sub(r"\s+", "", uri.group(3))
            try:
                binary = base64.b64decode(encoded, validate=True)
            except (ValueError, base64.binascii.Error) as exc:
                raise ValueError("Invalid inline PNG data in HTML") from exc
            if not binary.startswith(b"\x89PNG\r\n\x1a\n"):
                raise ValueError("Image src is not a valid PNG")
            cid = f"gm-report-{uuid.uuid4().hex}@inline"
            attachments.append((cid, binary))
            return tag[:uri.start(2)] + f"cid:{cid}" + tag[uri.end(2):]

        outlook_html = image_tag_pattern.sub(replace_img_tag, html)
        mail = None
        try:
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # olMailItem
            mail.Subject = subject
            with tempfile.TemporaryDirectory(prefix="gm_oft_") as tmp:
                for idx, (cid, binary) in enumerate(attachments, start=1):
                    temp_png = Path(tmp) / f"image_{idx}.png"
                    temp_png.write_bytes(binary)
                    attachment = mail.Attachments.Add(str(temp_png))
                    attachment.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
                    )
                    attachment.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True
                    )
                mail.HTMLBody = outlook_html
                mail.SaveAs(str(destination), 2)  # olTemplate
            if not destination.is_file():
                raise RuntimeError("Outlook did not produce the expected .oft file")
            self._logger.info("Generated Outlook template: %s", destination)
            return str(destination)
        except Exception as exc:
            raise RuntimeError(f"Could not generate Outlook template: {exc}") from exc
        finally:
            # Release COM references. No email is displayed or sent.
            mail = None

    @staticmethod
    def _format_date(value: str | date | datetime) -> str:
        if isinstance(value, (date, datetime)):
            return value.strftime("%d %b %Y")
        try:
            return date.fromisoformat(str(value).strip()[:10]).strftime("%d %b %Y")
        except ValueError as exc:
            raise ValueError("cobdate must be YYYY-MM-DD or a date object") from exc

    @staticmethod
    def _replace_tokens(value: Any, replacements: dict[str, str]) -> str:
        output = "" if value is None else str(value)
        for key, replacement in replacements.items():
            output = output.replace("{{" + key + "}}", replacement)
        return output

    @staticmethod
    def _escaped_multiline(text: str) -> str:
        return "<br>".join(escape(str(text)).splitlines())

    @staticmethod
    def _dimension(value: Any, name: str) -> int | None:
        """Blank means automatic; all explicit dimensions must be positive pixels."""
        if value is None or str(value).strip() == "":
            return None
        try:
            number = float(value)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{name} must be a positive integer (pixels) or blank") from exc
        if not number.is_integer() or not 1 <= number <= 5000:
            raise ValueError(f"{name} must be an integer between 1 and 5000 pixels")
        return int(number)

    @staticmethod
    def _image_html(data_uri: str, caption: str, width: int | None = None,
                    height: int | None = None) -> str:
        """Configured width/height are HTML display dimensions, not PNG resolution."""
        width = 650 if width is None else width
        caption_html = (f'<p style="margin:8px 0 4px 0;"><strong>{escape(caption)}</strong></p>'
                        if caption.strip() else "")
        height_attr = f' height="{height}"' if height is not None else ""
        height_css = f"{height}px" if height is not None else "auto"
        return (caption_html +
                f'<p style="margin:0 0 12px 0;"><img src="{data_uri}" '
                f'alt="{escape(caption, quote=True)}" width="{width}"{height_attr} '
                f'style="display:block;max-width:100%;width:{width}px;'
                f'height:{height_css};border:0;">'
                '</p>')

    def _render_layout(self, layout: str, replacements: dict[str, str]) -> str:
        """Each nonempty config line becomes a paragraph unless it is an image."""
        blocks = []
        for line in layout.splitlines():
            line = line.strip()
            if not line:
                continue
            standalone = self._TOKEN_RE.fullmatch(line)
            if standalone:
                key = standalone.group(1)
                if key not in self._TOKENS:
                    raise ValueError(f"Unknown layout token: {key}")
                fragment = replacements.get(key, "")
                if not fragment:
                    continue
                if key.startswith("Image"):
                    blocks.append(fragment)
                else:
                    blocks.append(f'<p style="margin:0 0 12px 0;">{fragment}</p>')
                continue
            # Mixed prose/text placeholders, support custom positioning on the same line.
            output = []
            start = 0
            for match in self._TOKEN_RE.finditer(line):
                output.append(escape(line[start:match.start()]))
                key = match.group(1)
                if key not in self._TOKENS:
                    raise ValueError(f"Unknown layout token: {key}")
                output.append(replacements.get(key, ""))
                start = match.end()
            output.append(escape(line[start:]))
            rendered = "".join(output).strip()
            if rendered:
                blocks.append(f'<p style="margin:0 0 12px 0;">{rendered}</p>')
        return "\n".join(blocks)

    @staticmethod
    def _validate_link(value: str) -> None:
        parsed = urlparse(value)
        if parsed.scheme.lower() not in ("https", "http", "mailto"):
            raise ValueError("link_url must start with https://, http://, or mailto:")
        if parsed.scheme.lower() in ("http", "https") and not parsed.netloc:
            raise ValueError("link_url must contain a hostname")

    @staticmethod
    def _validate_image_dict(imagedict: Any) -> None:
        if not isinstance(imagedict, dict):
            raise TypeError("imagedict must be a dictionary ({} for no images)")
        unknown = set(imagedict) - {"Image1", "Image2", "Image3"}
        if unknown:
            raise ValueError(f"Unsupported image keys: {sorted(unknown)}")
        for key, entry in imagedict.items():
            if not isinstance(entry, dict):
                raise TypeError(f"{key} must be a dictionary")
            required = {"Sheetname", "Imagerange", "Imageinfo"}
            if not required.issubset(entry):
                raise ValueError(f"{key} missing fields: {sorted(required - set(entry))}")
            if not str(entry["Sheetname"]).strip() or not str(entry["Imagerange"]).strip():
                raise ValueError(f"{key} Sheetname and Imagerange cannot be empty")
            try:
                min_c, min_r, max_c, max_r = range_boundaries(entry["Imagerange"])
                if not all((min_c, min_r, max_c, max_r)) or (max_c - min_c + 1) * (max_r - min_r + 1) > 4000:
                    raise ValueError("Invalid or oversized range")
            except (TypeError, ValueError) as exc:
                raise ValueError(f"Invalid Imagerange for {key}: {entry['Imagerange']}") from exc

    def _load_configuration(self, reportid: str) -> dict[str, str]:
        wb = load_workbook(self._config_path, data_only=True, read_only=True)
        try:
            if len(wb.sheetnames) != 1:
                raise ValueError("Configuration workbook must have exactly one sheet")
            rows = wb.active.iter_rows(values_only=True)
            first = next(rows, None)
            if not first:
                raise ValueError("Configuration workbook is empty")
            headers = [str(v).strip() if v is not None else "" for v in first]
            missing = set(self._FIELDS) - set(headers)
            if missing:
                raise ValueError(f"Missing configuration columns: {sorted(missing)}")
            missing_dimensions = set(self._DIMENSION_FIELDS) - set(headers)
            if missing_dimensions:
                raise ValueError(f"Missing image dimension columns: {sorted(missing_dimensions)}")
            found = None
            for row in rows:
                record = {name: (row[idx] if idx < len(row) else None)
                          for idx, name in enumerate(headers)}
                if str(record.get("ReportID") or "").strip().lower() == reportid.lower():
                    if found is not None:
                        raise ValueError(f"Duplicate ReportID: {reportid}")
                    found = {key: "" if record[key] is None else str(record[key])
                             for key in self._FIELDS}
                    for key in self._DIMENSION_FIELDS:
                        found[key] = self._dimension(record.get(key), key)
            if found is None:
                raise ValueError(f"ReportID not found: {reportid}")
            if not found["Subject"].strip() or not found["EmailLayout"].strip():
                raise ValueError("Subject and EmailLayout are required")
            return found
        finally:
            wb.close()

    @staticmethod
    def _argb_to_rgb(color: Any, default: str) -> str:
        if color is None:
            return default
        if color.type == "rgb" and isinstance(color.rgb, str):
            return "#" + color.rgb[-6:]
        # Theme colors are not fully resolved by openpyxl; use legible fallback.
        return default

    @staticmethod
    def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
        names = (["DejaVuSans-Bold.ttf", "Arial Bold.ttf", "arialbd.ttf"] if bold else
                 ["DejaVuSans.ttf", "Arial.ttf", "arial.ttf"])
        for name in names:
            try:
                return ImageFont.truetype(name, size)
            except OSError:
                continue
        return ImageFont.load_default()

    @staticmethod
    def _display_value(cell: Any) -> str:
        value = cell.value
        if value is None:
            return ""
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        if isinstance(value, (datetime, date)):
            return value.strftime("%d %b %Y")
        if isinstance(value, (int, float)):
            fmt = str(cell.number_format or "General")
            if "%" in fmt:
                match = re.search(r"\.([0#]+)%", fmt)
                places = len(match.group(1)) if match else 0
                return f"{value:.{places}%}"
            if is_date_format(fmt):
                # In normal workbooks openpyxl converts numeric Excel dates to datetime.
                return str(value)
            decimals = 0
            pattern = re.search(r"\.(0+|#+)", fmt)
            if pattern:
                decimals = len(pattern.group(1))
            grouping = "," in fmt.split(";")[0]
            spec = f",.{decimals}f" if grouping else f".{decimals}f"
            result = format(value, spec)
            if value < 0 and "(" in fmt and ")" in fmt:
                result = f"({format(abs(value), spec)})"
            if "€" in fmt:
                result = "€" + result
            elif "$" in fmt:
                result = "$" + result
            return result
        return str(value)

    @staticmethod
    def _wrap(text: str, font: Any, max_width: int, draw: ImageDraw.ImageDraw) -> list[str]:
        if not text:
            return []
        result: list[str] = []
        for paragraph in text.split("\n"):
            if paragraph == "":
                result.append("")
                continue
            current = ""
            for word in paragraph.split(" "):
                trial = f"{current} {word}".strip()
                if current and draw.textbbox((0, 0), trial, font=font)[2] > max_width:
                    result.append(current)
                    current = word
                else:
                    current = trial
            if current:
                result.append(current)
        return result

    def _render_range_to_png(self, ws: Any, cell_range: str) -> bytes:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        scale = 1.7
        margin = 3
        cols = list(range(min_col, max_col + 1))
        rows = list(range(min_row, max_row + 1))
        col_sizes = []
        row_sizes = []
        for col in cols:
            from openpyxl.utils import get_column_letter
            letter = get_column_letter(col)
            dimension = ws.column_dimensions[letter]
            width = dimension.width if dimension.width is not None else 13
            col_sizes.append(max(48, min(320, int((width * 7 + 7) * scale))))
        for row in rows:
            dimension = ws.row_dimensions[row]
            height = dimension.height if dimension.height is not None else 19
            row_sizes.append(max(29, min(125, int(height * (96 / 72) * scale))))
        x_offsets, y_offsets = [margin], [margin]
        for w in col_sizes:
            x_offsets.append(x_offsets[-1] + w)
        for h in row_sizes:
            y_offsets.append(y_offsets[-1] + h)
        image = Image.new("RGB", (x_offsets[-1] + margin, y_offsets[-1] + margin), "white")
        draw = ImageDraw.Draw(image)
        merged_topleft = {}
        covered = set()
        for merged in ws.merged_cells.ranges:
            c1, r1, c2, r2 = merged.bounds
            if c1 <= max_col and c2 >= min_col and r1 <= max_row and r2 >= min_row:
                for r in range(max(r1, min_row), min(r2, max_row) + 1):
                    for c in range(max(c1, min_col), min(c2, max_col) + 1):
                        if (r, c) != (r1, c1):
                            covered.add((r, c))
                if min_row <= r1 <= max_row and min_col <= c1 <= max_col:
                    merged_topleft[(r1, c1)] = (min(r2, max_row), min(c2, max_col))
        for row in rows:
            for col in cols:
                if (row, col) in covered:
                    continue
                cell = ws.cell(row, col)
                bottom, right = merged_topleft.get((row, col), (row, col))
                x0, y0 = x_offsets[col - min_col], y_offsets[row - min_row]
                x1, y1 = x_offsets[right - min_col + 1], y_offsets[bottom - min_row + 1]
                background = "#FFFFFF"
                if cell.fill.patternType == "solid":
                    background = self._argb_to_rgb(cell.fill.fgColor, "#FFFFFF")
                draw.rectangle((x0, y0, x1, y1), fill=background)
                # Draw default subtle grid only when explicit border not specified.
                sides = [("top", (x0, y0, x1, y0)), ("bottom", (x0, y1, x1, y1)),
                         ("left", (x0, y0, x0, y1)), ("right", (x1, y0, x1, y1))]
                for label, line in sides:
                    border = getattr(cell.border, label, None)
                    explicit = border is not None and border.style is not None
                    line_color = self._argb_to_rgb(border.color, "#8192A8") if explicit else "#E6EBF1"
                    stroke = 2 if explicit and border.style in ("medium", "thick", "double") else 1
                    draw.line(line, fill=line_color, width=stroke)
                text = self._display_value(cell)
                if not text:
                    continue
                size = max(10, int((cell.font.sz or 10) * scale * 0.82))
                font = self._font(size, bold=bool(cell.font.bold))
                color = self._argb_to_rgb(cell.font.color, "#172B4D")
                if cell.font.color is None and background in ("#15324F", "#17365D", "#102A43"):
                    color = "#FFFFFF"
                pad = int(8 * scale)
                max_text = max(8, x1 - x0 - 2 * pad)
                lines = self._wrap(text, font, max_text, draw)
                line_height = size + 4
                max_lines = max(1, int((y1 - y0 - 2 * pad) / line_height))
                if len(lines) > max_lines:
                    lines = lines[:max_lines]
                    last = lines[-1]
                    while last and draw.textbbox((0, 0), last + "…", font=font)[2] > max_text:
                        last = last[:-1]
                    lines[-1] = last + "…"
                total_h = len(lines) * line_height
                val = cell.alignment.vertical
                yy = (y0 + pad if val == "top" else y1 - pad - total_h if val == "bottom" else
                      y0 + (y1 - y0 - total_h) / 2)
                horizontal = cell.alignment.horizontal or ("right" if isinstance(cell.value, (int, float)) else "left")
                for line in lines:
                    text_w = draw.textbbox((0, 0), line, font=font)[2]
                    xx = (x0 + (x1 - x0 - text_w) / 2 if horizontal in ("center", "centerContinuous") else
                          x1 - pad - text_w if horizontal == "right" else x0 + pad)
                    draw.text((xx, yy), line, font=font, fill=color)
                    yy += line_height
        buffer = BytesIO()
        image.save(buffer, format="PNG", optimize=True)
        return buffer.getvalue()
