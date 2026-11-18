


def write_dataframe(
    self,
    df,
    sheet_name,
    start_row=1,
    start_col=1,
    include_header=True
):
    """
    Write a Pandas DataFrame to a worksheet
    in the existing workbook.

    Creates the worksheet if it does not exist.
    """

    import pandas as pd
    from openpyxl.utils.dataframe import dataframe_to_rows

    if not isinstance(df, pd.DataFrame):
        raise TypeError("df must be a Pandas DataFrame.")

    if not isinstance(sheet_name, str) or not sheet_name.strip():
        raise ValueError("A valid sheet_name is required.")

    if start_row < 1 or start_col < 1:
        raise ValueError("start_row and start_col must be >= 1.")

    # Reuse existing worksheet management
    ws = self._select_worksheet(sheet_name)

    # Write DataFrame columns and values
    for row_idx, values in enumerate(
        dataframe_to_rows(
            df,
            index=False,
            header=include_header
        ),
        start=start_row
    ):
        for col_idx, value in enumerate(
            values,
            start=start_col
        ):
            if pd.api.types.is_scalar(value) and pd.isna(value):
                value = None

            cell = ws.cell(
                row=row_idx,
                column=col_idx
            )
            cell.value = value

            # Apply header formatting
            if include_header and row_idx == start_row:
                cell.font = Font(
                    name="Aptos",
                    size=11,
                    bold=True,
                    color=self.COLORS["white"]
                )
                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor=self.COLORS["green"]
                )
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

    return start_row + len(df) + int(include_header)




#### Usage
report.write_dataframe(
    df=df,
    sheet_name="Raw Data"
)

