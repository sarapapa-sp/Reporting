import json


from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


class ExcelReportGenerator:

    COLORS = {
        # -----------------------------------------------------
        # Column heading grading
        # -----------------------------------------------------
        "green": "3C8D5A",
        "dark_grey": "56616B",
        "light_grey": "D9DDE2",

        # -----------------------------------------------------
        # Grand Total
        # -----------------------------------------------------
        "grand_total_green": "3C8D5A",
        "grand_total_blue": "4472C4",

        # -----------------------------------------------------
        # Level 1
        # -----------------------------------------------------
        "level_1_fill": "E5E7E9",
        "level_1_border": "3C8D5A",

        # -----------------------------------------------------
        # Section heading
        # -----------------------------------------------------
        "header_display_fill": "EAF4EE",

        # -----------------------------------------------------
        # General
        # -----------------------------------------------------
        "white": "FFFFFF",
        "text": "12263A",
        "navy": "17365D",
        "header_border": "7F8C8D",
    }

    # =========================================================
    # ROW HEIGHTS
    #
    # Excel uses points for row height.
    #
    # Approximate:
    # 35 pixels = 26.25 points
    # 25 pixels = 18.75 points
    #  3 pixels =  2.25 points
    # =========================================================

    GRAND_TOTAL_ROW_HEIGHT = 26.25
    LEVEL_1_ROW_HEIGHT = 26.25
    CHILD_ROW_HEIGHT = 18.75
    SPACER_ROW_HEIGHT = 2.25

    # =========================================================
    # COLUMN WIDTHS
    #
    # These are approximate because Excel column widths are
    # not directly specified in pixels.
    # =========================================================

    SECTION_HEADING_WIDTH = 8.5       # ~60 pixels
    SMALL_SPACER_WIDTH = 0.5          # ~5 pixels
    LABEL_COLUMN_WIDTH = 30
    VALUE_COLUMN_WIDTH = 16

    # =========================================================
    # PERCENTAGE LIMITS
    # =========================================================

    PERCENTAGE_MAX = 200
    PERCENTAGE_MIN = -200

    def __init__(
        self,
        file_name="global_markets_report.xlsx",
        sheet_name="Global Markets Report"
    ):

        self.file_name = file_name
        self.sheet_name = sheet_name

        self.workbook = None
        self.worksheet = None

        self._create_file()

    # =========================================================
    # FILE CREATION
    # =========================================================

    def _create_file(self):

        self.workbook = Workbook()

        self.worksheet = self.workbook.active

        self.worksheet.title = self.sheet_name

        # No frozen panes
        self.worksheet.freeze_panes = None

        # Hide gridlines
        self.worksheet.sheet_view.showGridLines = False

    # =========================================================
    # CREATE REPORT
    # =========================================================

    def create_report(
        self,
        data,
        column_mapping,
        opening_statement,
        report_name,
        header_display=None,
        grant_total=False,
        total_color=True,
        write_column_headings=True,
        start_row=3,
        start_col=3
    ):
        """
        Create one report section.

        Column layout:

            start_col
                |
                +-- Section heading
                |
                +-- 5px spacer
                |
                +-- Labels
                |
                +-- 5px spacer
                |
                +-- Value columns
                    |
                    +-- part-based separators

        Parameters
        ----------
        data : dict
            Nested report dictionary.

        column_mapping : dict
            Example:

            {
                "PnL": {
                    "mapping_name": "PNL",
                    "type": "Actuals",
                    "category": "sum",
                    "part": "Financial",
                    "color_grading": "green"
                },

                "Revenue": {
                    "mapping_name": "Revenue",
                    "type": "Actuals",
                    "category": "sum",
                    "part": "Financial",
                    "color_grading": "green"
                }
            }

        opening_statement : str
            Example:
                "Meur. As of 09.09.2026"

        report_name : str
            Grand Total label.

        header_display : str
            Section heading.

        grant_total : bool
            Whether to write Grand Total.

        total_color : bool
            True  -> green
            False -> blue

        write_column_headings : bool
            True:
                Write opening statement and column headings.

            False:
                Write neither.

        start_row : int
            Starting row.

        start_col : int
            Section heading column.
        """

        # -----------------------------------------------------
        # Calculate column positions
        # -----------------------------------------------------

        header_display_col = start_col

        header_display_spacer_col = (
            start_col + 1
        )

        label_col = (
            start_col + 2
        )

        label_spacer_col = (
            start_col + 3
        )

        value_start_col = (
            start_col + 4
        )

        # -----------------------------------------------------
        # Prepare mapping with part separators
        # -----------------------------------------------------

        column_layout = self._build_column_layout(
            column_mapping=column_mapping,
            value_start_col=value_start_col
        )

        # -----------------------------------------------------
        # Header
        # -----------------------------------------------------

        if write_column_headings:

            self._write_header(
                opening_statement=opening_statement,
                column_layout=column_layout,
                row=start_row,
                label_col=label_col
            )

            # Spacer after heading
            header_spacer_row = start_row + 1

            self._write_spacer_row(
                row=header_spacer_row,
                start_col=header_display_col,
                number_of_columns=len(column_layout) + 5
            )

            next_row = (
                header_spacer_row + 1
            )

        else:

            # No opening statement
            # No column headings
            # No header spacer row

            next_row = start_row

        # -----------------------------------------------------
        # Section boundaries
        # -----------------------------------------------------

        first_section_row = None
        last_section_row = None

        # =====================================================
        # GRAND TOTAL
        # =====================================================

        if grant_total:

            grand_total_row = next_row

            next_row = self._write_node(
                node=data,
                level=0,
                row=grand_total_row,
                label_col=label_col,
                column_layout=column_layout,
                report_name=report_name,
                total_color=total_color,
                write_children=False
            )

            # 35 pixels
            self.worksheet.row_dimensions[
                grand_total_row
            ].height = self.GRAND_TOTAL_ROW_HEIGHT

            # Spacer after Grand Total
            self._write_spacer_row(
                row=next_row,
                start_col=header_display_col,
                number_of_columns=len(column_layout) + 5
            )

            next_row += 1

        # =====================================================
        # LEVEL 1 + CHILDREN
        # =====================================================

        for child in data.get("child", []):

            if first_section_row is None:

                first_section_row = next_row

            next_row = self._write_node(
                node=child,
                level=1,
                row=next_row,
                label_col=label_col,
                column_layout=column_layout,
                report_name=report_name,
                total_color=total_color,
                write_children=True
            )

            last_section_row = next_row - 1

        # =====================================================
        # SECTION HEADING
        # =====================================================

        if (
            header_display
            and first_section_row is not None
            and last_section_row is not None
        ):

            self._write_header_display(
                header_display=header_display,
                start_row=first_section_row,
                end_row=last_section_row,
                col=header_display_col,
                total_color=total_color
            )

        # =====================================================
        # LAYOUT
        # =====================================================

        self._apply_layout(
            column_mapping=column_mapping,
            column_layout=column_layout,
            start_row=start_row,
            start_col=start_col,
            label_col=label_col,
            end_row=next_row - 1
        )

        return next_row

    # =========================================================
    # BUILD COLUMN LAYOUT
    # =========================================================

    def _build_column_layout(
        self,
        column_mapping,
        value_start_col
    ):
        """
        Builds the physical Excel column layout.

        If:

            A -> part A
            B -> part A
            C -> part B
            D -> part B
            E -> part C

        Physical layout becomes:

            A
            B
            blank
            C
            D
            blank
            E

        Returns a list containing both value columns and
        separator columns.
        """

        layout = []

        previous_part = None
        current_col = value_start_col

        for key, config in column_mapping.items():

            current_part = config.get(
                "part",
                None
            )

            # -------------------------------------------------
            # Insert separator when part changes
            # -------------------------------------------------

            if (
                previous_part is not None
                and current_part != previous_part
            ):

                layout.append({
                    "kind": "separator",
                    "column": current_col
                })

                current_col += 1

            # -------------------------------------------------
            # Add actual value column
            # -------------------------------------------------

            layout.append({
                "kind": "value",
                "column": current_col,
                "key": key,
                "config": config
            })

            current_col += 1

            previous_part = current_part

        return layout

    # =========================================================
    # WRITE HEADER
    # =========================================================

    def _write_header(
        self,
        opening_statement,
        column_layout,
        row,
        label_col
    ):

        ws = self.worksheet

        # -----------------------------------------------------
        # Opening statement
        # -----------------------------------------------------

        cell = ws.cell(
            row=row,
            column=label_col,
            value=opening_statement
        )

        cell.fill = PatternFill(
            fill_type=None
        )

        cell.border = Border()

        cell.font = Font(
            name="Aptos",
            size=11,
            bold=True,
            color=self.COLORS["navy"]
        )

        cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        # -----------------------------------------------------
        # Spacer between label and values
        # -----------------------------------------------------

        separator_cell = ws.cell(
            row=row,
            column=label_col + 1
        )

        separator_cell.value = None
        separator_cell.fill = PatternFill(
            fill_type=None
        )
        separator_cell.border = Border()

        # -----------------------------------------------------
        # Column headings
        # -----------------------------------------------------

        for item in column_layout:

            col = item["column"]

            # -------------------------------------------------
            # Part separator
            # -------------------------------------------------

            if item["kind"] == "separator":

                cell = ws.cell(
                    row=row,
                    column=col
                )

                cell.value = None

                cell.fill = PatternFill(
                    fill_type=None
                )

                cell.border = Border()

                continue

            # -------------------------------------------------
            # Actual heading
            # -------------------------------------------------

            config = item["config"]

            cell = ws.cell(
                row=row,
                column=col,
                value=config["mapping_name"]
            )

            self._apply_column_heading_style(
                cell=cell,
                color_grading=config.get(
                    "color_grading"
                ),
                column_type=config.get(
                    "type"
                )
            )

            # -------------------------------------------------
            # Word wrap by default
            # -------------------------------------------------

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            cell.border = self._get_thin_border(
                self.COLORS["header_border"]
            )

        ws.row_dimensions[row].height = 28

    # =========================================================
    # COLUMN HEADING STYLE
    # =========================================================

    def _apply_column_heading_style(
        self,
        cell,
        color_grading,
        column_type=None
    ):
        """
        color_grading takes priority.

        Supported:
            green
            dark_grey
            light_grey

        If color_grading isn't supplied, the older
        type-based behavior is used as a fallback.
        """

        # -----------------------------------------------------
        # Explicit color grading
        # -----------------------------------------------------

        if color_grading:

            grading = color_grading.lower()

            if grading == "green":

                fill = self.COLORS["green"]
                font_color = self.COLORS["white"]

            elif grading in (
                "dark_grey",
                "dark grey"
            ):

                fill = self.COLORS["dark_grey"]
                font_color = self.COLORS["white"]

            elif grading in (
                "light_grey",
                "light grey"
            ):

                fill = self.COLORS["light_grey"]
                font_color = self.COLORS["text"]

            else:

                raise ValueError(
                    f"Unsupported color_grading: "
                    f"{color_grading}. "
                    f"Use green, dark_grey or light_grey."
                )

        # -----------------------------------------------------
        # Backward-compatible type-based fallback
        # -----------------------------------------------------

        else:

            if column_type == "Actuals":

                fill = self.COLORS["green"]
                font_color = self.COLORS["white"]

            elif column_type == "Previous":

                fill = self.COLORS["light_grey"]
                font_color = self.COLORS["text"]

            elif column_type == "Budget":

                fill = self.COLORS["dark_grey"]
                font_color = self.COLORS["white"]

            else:

                fill = self.COLORS["white"]
                font_color = self.COLORS["text"]

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=fill
        )

        cell.font = Font(
            name="Aptos",
            size=11,
            bold=True,
            color=font_color
        )

    # =========================================================
    # SECTION HEADING
    # =========================================================

    def _write_header_display(
        self,
        header_display,
        start_row,
        end_row,
        col,
        total_color
    ):
        """
        Vertically merged section heading.

        The cell covers all Level 1 rows and their
        descendants.
        """

        ws = self.worksheet

        # -----------------------------------------------------
        # Border colour based on total_color
        # -----------------------------------------------------

        border_color = (
            self.COLORS["grand_total_green"]
            if total_color
            else self.COLORS["grand_total_blue"]
        )

        side = Side(
            style="thin",
            color=border_color
        )

        # -----------------------------------------------------
        # Apply fill/borders before merge
        # -----------------------------------------------------

        for row in range(
            start_row,
            end_row + 1
        ):

            cell = ws.cell(
                row=row,
                column=col
            )

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor=self.COLORS[
                    "header_display_fill"
                ]
            )

            cell.border = Border(
                left=side,
                right=side
            )

        # -----------------------------------------------------
        # Merge
        # -----------------------------------------------------

        ws.merge_cells(
            start_row=start_row,
            start_column=col,
            end_row=end_row,
            end_column=col
        )

        # -----------------------------------------------------
        # Configure merged cell
        # -----------------------------------------------------

        cell = ws.cell(
            row=start_row,
            column=col
        )

        cell.value = header_display

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=self.COLORS[
                "header_display_fill"
            ]
        )

        cell.font = Font(
            name="Aptos",
            size=15,
            bold=True,
            color=border_color
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            textRotation=90,
            wrap_text=True
        )

        cell.border = Border(
            left=side,
            right=side
        )

        # -----------------------------------------------------
        # Reapply vertical borders
        # -----------------------------------------------------

        for row in range(
            start_row,
            end_row + 1
        ):

            merged_cell = ws.cell(
                row=row,
                column=col
            )

            merged_cell.border = Border(
                left=side,
                right=side
            )

    # =========================================================
    # SPACER ROW
    # =========================================================

    def _write_spacer_row(
        self,
        row,
        start_col,
        number_of_columns
    ):

        ws = self.worksheet

        ws.row_dimensions[row].height = (
            self.SPACER_ROW_HEIGHT
        )

        for offset in range(
            number_of_columns
        ):

            cell = ws.cell(
                row=row,
                column=start_col + offset
            )

            cell.value = None

            cell.fill = PatternFill(
                fill_type=None
            )

            cell.border = Border()

    # =========================================================
    # WRITE NODE
    # =========================================================

    def _write_node(
        self,
        node,
        level,
        row,
        label_col,
        column_layout,
        report_name,
        total_color,
        write_children=True
    ):

        ws = self.worksheet

        values = node.get(
            "values",
            {}
        )

        children = node.get(
            "child",
            []
        )

        # -----------------------------------------------------
        # Label
        # -----------------------------------------------------

        if level == 0:

            label = report_name

        else:

            label = node.get(
                "root_value",
                ""
            )

        label_cell = ws.cell(
            row=row,
            column=label_col,
            value=label
        )

        # -----------------------------------------------------
        # Alignment
        #
        # Level 0 = left
        # Level 1 = left
        # Level 2+ = right
        # -----------------------------------------------------

        if level <= 1:

            horizontal_alignment = "left"

        else:

            horizontal_alignment = "right"

        label_cell.alignment = Alignment(
            horizontal=horizontal_alignment,
            vertical="center",
            indent=level
        )

        # =====================================================
        # LEVEL 0 / GRAND TOTAL
        # =====================================================

        if level == 0:

            total_fill = (
                self.COLORS["grand_total_green"]
                if total_color
                else self.COLORS["grand_total_blue"]
            )

            label_cell.fill = PatternFill(
                fill_type="solid",
                fgColor=total_fill
            )

            label_cell.font = Font(
                name="Aptos",
                size=12,
                bold=True,
                color=self.COLORS["white"]
            )

            label_cell.border = Border()

            ws.row_dimensions[
                row
            ].height = self.GRAND_TOTAL_ROW_HEIGHT

        # =====================================================
        # LEVEL 1
        # =====================================================

        elif level == 1:

            label_cell.fill = PatternFill(
                fill_type="solid",
                fgColor=self.COLORS["level_1_fill"]
            )

            label_cell.font = Font(
                name="Aptos",
                size=12,
                bold=True,
                color=self.COLORS["text"]
            )

            label_cell.border = self._get_bottom_border(
                self.COLORS["level_1_border"]
            )

            ws.row_dimensions[
                row
            ].height = self.LEVEL_1_ROW_HEIGHT

        # =====================================================
        # LEVEL 2+
        # =====================================================

        else:

            label_cell.fill = PatternFill(
                fill_type="solid",
                fgColor=self.COLORS["white"]
            )

            label_cell.font = Font(
                name="Aptos",
                size=11,
                bold=False,
                color=self.COLORS["text"]
            )

            label_cell.border = Border()

            ws.row_dimensions[
                row
            ].height = self.CHILD_ROW_HEIGHT

        # -----------------------------------------------------
        # Spacer between labels and values
        # -----------------------------------------------------

        separator_cell = ws.cell(
            row=row,
            column=label_col + 1
        )

        separator_cell.value = None

        separator_cell.fill = PatternFill(
            fill_type="solid",
            fgColor=self.COLORS["white"]
        )

        separator_cell.border = Border()

        # =====================================================
        # VALUE COLUMNS
        # =====================================================

        for item in column_layout:

            col = item["column"]

            # -------------------------------------------------
            # Part separator
            # -------------------------------------------------

            if item["kind"] == "separator":

                separator = ws.cell(
                    row=row,
                    column=col
                )

                separator.value = None

                separator.fill = PatternFill(
                    fill_type="solid",
                    fgColor=self.COLORS["white"]
                )

                separator.border = Border()

                continue

            # -------------------------------------------------
            # Value column
            # -------------------------------------------------

            key = item["key"]
            config = item["config"]

            raw_value = values.get(
                key
            )

            value, number_format = (
                self._format_value(
                    value=raw_value,
                    category=config["category"]
                )
            )

            cell = ws.cell(
                row=row,
                column=col,
                value=value
            )

            cell.number_format = number_format

            cell.alignment = Alignment(
                horizontal="right",
                vertical="center"
            )

            # =================================================
            # GRAND TOTAL
            # =================================================

            if level == 0:

                total_fill = (
                    self.COLORS["grand_total_green"]
                    if total_color
                    else self.COLORS["grand_total_blue"]
                )

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor=total_fill
                )

                cell.font = Font(
                    name="Aptos",
                    size=12,
                    bold=True,
                    color=self.COLORS["white"]
                )

                cell.border = Border()

            # =================================================
            # LEVEL 1
            # =================================================

            elif level == 1:

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor=self.COLORS["level_1_fill"]
                )

                cell.font = Font(
                    name="Aptos",
                    size=12,
                    bold=True,
                    color=self.COLORS["text"]
                )

                cell.border = self._get_bottom_border(
                    self.COLORS["level_1_border"]
                )

            # =================================================
            # LEVEL 2+
            # =================================================

            else:

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor=self.COLORS["white"]
                )

                cell.font = Font(
                    name="Aptos",
                    size=11,
                    bold=False,
                    color=self.COLORS["text"]
                )

                cell.border = Border()

        # -----------------------------------------------------
        # Excel outline
        # -----------------------------------------------------

        if children:

            ws.row_dimensions[
                row
            ].outlineLevel = level

        # -----------------------------------------------------
        # Do not recursively write Grand Total children
        # -----------------------------------------------------

        if not write_children:

            return row + 1

        # -----------------------------------------------------
        # Write children
        # -----------------------------------------------------

        next_row = row + 1

        for child in children:

            next_row = self._write_node(
                node=child,
                level=level + 1,
                row=next_row,
                label_col=label_col,
                column_layout=column_layout,
                report_name=report_name,
                total_color=total_color,
                write_children=True
            )

        return next_row

    # =========================================================
    # VALUE FORMATTING
    # =========================================================

    def _format_value(
        self,
        value,
        category
    ):
        """
        Returns:

            value_to_write,
            Excel number format
        """

        if value is None:

            return None, "General"

        # -----------------------------------------------------
        # SUM
        # -----------------------------------------------------

        if category == "sum":

            return (
                value,
                '#,##0;[Red](#,##0);-'
            )

        # -----------------------------------------------------
        # PERCENTAGE
        # -----------------------------------------------------

        elif category == "percentage":

            return self._parse_percentage(
                value
            )

        # -----------------------------------------------------
        # PERCENTAGE - 100
        # -----------------------------------------------------

        elif category == "percentage-100":

            result = 100 - value

            return self._parse_percentage(
                result
            )

        # -----------------------------------------------------
        # DEFAULT
        # -----------------------------------------------------

        else:

            return (
                value,
                "General"
            )

    # =========================================================
    # PERCENTAGE PARSER
    # =========================================================

    def _parse_percentage(
        self,
        value,
        max_percentage=None,
        min_percentage=None
    ):
        """
        Parse a percentage value.

        Input is expected as a number such as:

            25
            150
            250
            -50
            -120

        Results:

            25   -> Excel 0.25 -> 25.00%
            150  -> Excel 1.50 -> 150.00%
            250  -> ">200"
            -50  -> Excel -0.50 -> (50.00%) depending
                     on number format
            -120 -> "NA"

        Thresholds can be overridden.

        Parameters
        ----------
        max_percentage : number
            Default = 200

        min_percentage : number
            Default = -100
        """

        if max_percentage is None:

            max_percentage = self.PERCENTAGE_MAX

        if min_percentage is None:

            min_percentage = self.PERCENTAGE_MIN

        # -----------------------------------------------------
        # Greater than maximum
        # -----------------------------------------------------

        if value > max_percentage:

            return (
                f">{max_percentage}",
                "General"
            )

        # -----------------------------------------------------
        # Less than minimum
        # -----------------------------------------------------

        if value < min_percentage:

            return (
                "NA",
                "General"
            )

        # -----------------------------------------------------
        # Normal percentage
        # -----------------------------------------------------

        return (
            value / 100,
            '0%;[Red](0%);-'
        )

    # =========================================================
    # BORDER HELPERS
    # =========================================================

    def _get_thin_border(
        self,
        color
    ):

        side = Side(
            style="thin",
            color=color
        )

        return Border(
            left=side,
            right=side,
            top=side,
            bottom=side
        )

    def _get_bottom_border(
        self,
        color
    ):

        side = Side(
            style="thin",
            color=color
        )

        return Border(
            bottom=side
        )

    # =========================================================
    # LAYOUT
    # =========================================================

    def _apply_layout(
        self,
        column_mapping,
        column_layout,
        start_row,
        start_col,
        label_col,
        end_row
    ):

        ws = self.worksheet

        # -----------------------------------------------------
        # Section heading
        #
        # Approximately 60 pixels.
        # -----------------------------------------------------

        ws.column_dimensions[
            get_column_letter(start_col)
        ].width = self.SECTION_HEADING_WIDTH

        # -----------------------------------------------------
        # Spacer between section heading and labels
        #
        # Approximately 5 pixels.
        # -----------------------------------------------------

        ws.column_dimensions[
            get_column_letter(start_col + 1)
        ].width = self.SMALL_SPACER_WIDTH

        # -----------------------------------------------------
        # Label column
        # -----------------------------------------------------

        ws.column_dimensions[
            get_column_letter(label_col)
        ].width = self.LABEL_COLUMN_WIDTH

        # -----------------------------------------------------
        # Spacer between labels and values
        #
        # Approximately 5 pixels.
        # -----------------------------------------------------

        ws.column_dimensions[
            get_column_letter(label_col + 1)
        ].width = self.SMALL_SPACER_WIDTH

        # -----------------------------------------------------
        # Value columns + part separators
        # -----------------------------------------------------

        for item in column_layout:

            col = item["column"]

            if item["kind"] == "separator":

                ws.column_dimensions[
                    get_column_letter(col)
                ].width = self.SMALL_SPACER_WIDTH

            else:

                ws.column_dimensions[
                    get_column_letter(col)
                ].width = self.VALUE_COLUMN_WIDTH

        # -----------------------------------------------------
        # Print setup
        # -----------------------------------------------------

        ws.page_setup.orientation = "landscape"

        ws.page_setup.fitToWidth = 1

        ws.page_setup.fitToHeight = 0

        ws.sheet_properties.pageSetUpPr.fitToPage = True

        # -----------------------------------------------------
        # No frozen panes
        # -----------------------------------------------------

        ws.freeze_panes = None

    # =========================================================
    # SAVE
    # =========================================================

    def save(self):

        self.workbook.save(
            self.file_name
        )

        return Path(
            self.file_name
        )



# -------------------------------------------
# Usage
# -------------------------------------------
mapping_file_name = r'D:\Shubham\GMP\Mapping.json'

with open(mapping_file_name, 'r') as f:
    column_mapping = json.load(f)


with open("dict.json", "r") as f:
    data = json.load(f)

report = ExcelReportGenerator(
    file_name="global_markets_report.xlsx",
    sheet_name="Global Markets Report"
)

next_row = report.create_report(
    data=data,
    column_mapping=column_mapping,
    opening_statement="Meur. As of 09.09.2026",
    report_name="Global Market at NBI",
    header_display="Global",
    grant_total=True,
    total_color=True,
    write_column_headings=True,
    start_row=3,
    start_col=3
)

next_row_after_apac = report.create_report(
    data=data,
    column_mapping=column_mapping,
    opening_statement="Meur. As of 09.09.2026",
    report_name="Global Market APAC at NBI",
    header_display="APAC",
    grant_total=True,
    total_color=False,
    write_column_headings=False,
    start_row=next_row+3,
    start_col=3
)

report.save()

print(f"Next available row: {next_row_after_apac}")