from sqlalchemy import create_engine
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

def get_connection() -> create_engine:
    """Establish a connection to the database and return the engine."""
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def fetch_data_from_db(engine: create_engine, query: str) -> pd.DataFrame:
    """Fetch data from the database using the provided query."""
    raw_connection = engine.raw_connection()
    try:
        cursor = raw_connection.cursor()
        try:
            cursor.execute(query)
            columns = None
            rows = None
            while True:
                if cursor.description:
                    columns = [col[0] for col in cursor.description]
                    rows = cursor.fetchall()
                    break
                if not cursor.nextset():
                    raise RuntimeError("Query did not return a result set.")
        finally:
            cursor.close()
    finally:
        raw_connection.close()

    if columns is None or rows is None:
        return pd.DataFrame()

    return pd.DataFrame.from_records(rows, columns=columns)

def save_to_pickle(df, filename):
    df.to_pickle(filename)
    print(f"Pickle File saved to {filename}")

def build_column_totals(df, columns):
    """Return a dict of column totals for the given columns."""
    return {col: df[col].sum() for col in columns if col in df.columns}


def _parse_history_date(value):
    if not isinstance(value, str):
        return None
    for date_format in ("%m-%d-%Y", "%Y-%m-%d"):
        parsed = pd.to_datetime(value, format=date_format, errors="coerce")
        if not pd.isna(parsed):
            return parsed
    return None


def save_to_excel(
    df,
    filename,
    summary=None,
    format_cols=None,
    decimal_cols=None,
    rolling_main_layout=False,
    pre_date_column_count=18,
    summary_label_col_idx=None,
    top_row_groups=None,
    header_overrides=None,
    header_fill_overrides=None,
    column_width_overrides=None,
    format_blank_summary_cells=True,
    title_block=None,
    header_row_override=None,
    show_weeknum_label=True,
    history_top_labels=None,
    weeknum_label_fill="#DDD9C4",
    column_format_overrides=None,
    integer_accounting_no_symbol=False,
):
    filename = pd.io.common.stringify_path(filename)
    from pathlib import Path

    output_path = Path(filename)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            worksheet_name = 'Sheet1'
            header_row = header_row_override if header_row_override is not None else (4 if rolling_main_layout else 3)
            data_start_row = header_row + 1
            df.to_excel(writer, sheet_name=worksheet_name, startrow=header_row, index=False)
            worksheet = writer.sheets[worksheet_name]
            workbook = writer.book
            last_data_row = len(df) + header_row
            last_col = len(df.columns) - 1

            integer_num_format = (
                '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
                if integer_accounting_no_symbol
                else '#,##0'
            )
            decimal_num_format = (
                '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                if integer_accounting_no_symbol
                else '#,##0.00'
            )

            default_header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D8E4BC',
                'border': 1,
            })
            pre_date_header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#B8CCE4',
                'border': 1,
            })
            weeknum_label_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': weeknum_label_fill,
                'border': 1,
            })
            odd_month_header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D8E4BC',
                'border': 1,
            })
            even_month_header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#F2DCDB',
                'border': 1,
            })
            summary_label_format = workbook.add_format({
                'bold': True,
                'align': 'left',
                'valign': 'vcenter',
                'bg_color': '#CCC0DA',
                'border': 1,
            })
            summary_number_format = workbook.add_format({
                'bold': True,
                'num_format': integer_num_format,
                'align': 'right',
                'valign': 'vcenter',
                'bg_color': '#E4DFEC',
                'border': 1,
            })
            summary_decimal_format = workbook.add_format({
                'bold': True,
                'num_format': decimal_num_format,
                'align': 'right',
                'valign': 'vcenter',
                'bg_color': '#E4DFEC',
                'border': 1,
            })
            blank_summary_format = workbook.add_format({
                'bg_color': '#E4DFEC',
                'border': 1,
            })

        header_fill_formats = {
            '#D8E4BC': default_header_format,
            '#B8CCE4': pre_date_header_format,
            '#DDD9C4': weeknum_label_format,
            '#F2DCDB': even_month_header_format,
        }

        def get_header_format(base_format, col_idx):
            if not header_fill_overrides or col_idx not in header_fill_overrides:
                return base_format
            color = header_fill_overrides[col_idx]
            if color in header_fill_formats:
                return header_fill_formats[color]
            override_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': color,
                'border': 1,
            })
            header_fill_formats[color] = override_format
            return override_format

        group_ranges: set[int] = set()
        if top_row_groups:
            for group in top_row_groups:
                start_col = int(group["start_col"])
                end_col = int(group.get("end_col", start_col))
                for col_idx in range(start_col, end_col + 1):
                    group_ranges.add(col_idx)

        for col_idx, col_name in enumerate(df.columns):
            header_format = default_header_format
            history_date = _parse_history_date(col_name) if rolling_main_layout else None
            if rolling_main_layout:
                if col_idx < pre_date_column_count:
                    header_format = pre_date_header_format
                elif history_date is not None:
                    header_format = (
                        odd_month_header_format
                        if history_date.month % 2 == 1
                        else even_month_header_format
                    )
            header_format = get_header_format(header_format, col_idx)

            display_name = header_overrides.get(col_idx, col_name) if header_overrides else col_name
            worksheet.write(header_row, col_idx, display_name, header_format)

            if rolling_main_layout and history_date is not None:
                if history_top_labels and col_idx in history_top_labels:
                    worksheet.write(
                        header_row - 2,
                        col_idx,
                        history_top_labels[col_idx],
                        pre_date_header_format,
                    )
                worksheet.write(
                    header_row - 1,
                    col_idx,
                    history_date.isocalendar().week,
                    header_format,
                )
            elif rolling_main_layout and col_idx in group_ranges:
                worksheet.write_blank(header_row - 1, col_idx, None, header_format)

        if (
            rolling_main_layout
            and show_weeknum_label
            and pre_date_column_count > 0
            and len(df.columns) > pre_date_column_count
        ):
            worksheet.write(
                header_row - 1,
                pre_date_column_count - 1,
                'WeekNum',
                weeknum_label_format,
            )

        if rolling_main_layout and top_row_groups:
            for group in top_row_groups:
                start_col = int(group["start_col"])
                end_col = int(group.get("end_col", start_col))
                label = str(group["label"])
                group_format = get_header_format(pre_date_header_format, start_col)
                if start_col == end_col:
                    worksheet.write(header_row - 1, start_col, label, group_format)
                else:
                    worksheet.merge_range(header_row - 1, start_col, header_row - 1, end_col, label, group_format)

        if title_block:
            title_align = title_block.get("align", "center")
            title_format = workbook.add_format({
                'bold': False,
                'align': title_align,
                'valign': 'vcenter',
                'bg_color': '#C4BD97',
                'border': 1,
                'font_size': 16,
            })
            subtitle_format = workbook.add_format({
                'bold': False,
                'align': title_align,
                'valign': 'vcenter',
                'bg_color': '#C4BD97',
                'border': 1,
                'font_size': 16,
            })
            start_row = int(title_block["start_row"])
            end_row = int(title_block.get("end_row", start_row))
            start_col = int(title_block["start_col"])
            end_col = int(title_block["end_col"])
            title_text = str(title_block["title"])
            subtitle_text = str(title_block["subtitle"])
            merge_cells = bool(title_block.get("merge_cells", True))
            if not merge_cells:
                worksheet.write(start_row, start_col, title_text, title_format)
                worksheet.write(end_row, start_col, subtitle_text, subtitle_format)
            else:
                if end_row == start_row:
                    worksheet.merge_range(start_row, start_col, end_row, end_col, f"{title_text}\n{subtitle_text}", title_format)
                else:
                    worksheet.merge_range(start_row, start_col, start_row, end_col, title_text, title_format)
                    worksheet.merge_range(end_row, start_col, end_row, end_col, subtitle_text, subtitle_format)

        # Add total and subtotal rows above the header row.
        if summary:
            label_col_idx = (
                summary_label_col_idx
                if summary_label_col_idx is not None
                else (7 if len(df.columns) > 7 else 0)
            )
            unformatted_summary_cols = {10, 12}
            if format_blank_summary_cells:
                for row_idx in (0, 1):
                    for col_idx in range(label_col_idx, len(df.columns)):
                        if col_idx in unformatted_summary_cols:
                            continue
                        worksheet.write_blank(row_idx, col_idx, None, blank_summary_format)

            worksheet.write(0, label_col_idx, 'Total', summary_label_format)
            worksheet.write(1, label_col_idx, 'Subtotal', summary_label_format)

            for col_idx, col in enumerate(df.columns):
                if col not in summary or col_idx == label_col_idx:
                    continue

                start_cell = xl_rowcol_to_cell(data_start_row, col_idx)
                end_cell = xl_rowcol_to_cell(last_data_row, col_idx)
                total_formula = f'=SUM({start_cell}:{end_cell})'
                subtotal_formula = f'=SUBTOTAL(9,{start_cell}:{end_cell})'

                value = summary[col]
                if isinstance(value, float) and not float(value).is_integer():
                    formula_format = summary_decimal_format
                else:
                    formula_format = summary_number_format

                worksheet.write_formula(0, col_idx, total_formula, formula_format, value)
                worksheet.write_formula(1, col_idx, subtotal_formula, formula_format, value)

        worksheet.set_row(0, None)
        worksheet.set_row(1, None)
        worksheet.autofilter(header_row, 0, last_data_row, last_col)
        worksheet.freeze_panes(data_start_row, 0)

        # Format integer columns
        if format_cols:
            number_format = workbook.add_format({'num_format': integer_num_format})
            for col in format_cols:
                if col in df.columns:
                    col_idx = df.columns.get_loc(col)
                    width = None
                    if rolling_main_layout and pre_date_column_count == 18 and col_idx == 16:
                        width = 10.5
                    elif rolling_main_layout and pre_date_column_count == 18 and col_idx == 17:
                        width = 11
                    worksheet.set_column(col_idx, col_idx, width, number_format)

        # Format decimal columns
        if decimal_cols:
            decimal_format = workbook.add_format({'num_format': decimal_num_format})
            for col in decimal_cols:
                if col in df.columns:
                    col_idx = df.columns.get_loc(col)
                    worksheet.set_column(col_idx, col_idx, None, decimal_format)

        if "ISBN" in df.columns:
            isbn_format = workbook.add_format({'num_format': '0'})
            col_idx = df.columns.get_loc("ISBN")
            worksheet.set_column(col_idx, col_idx, 13.5, isbn_format)

        if "Title" in df.columns:
            col_idx = df.columns.get_loc("Title")
            worksheet.set_column(col_idx, col_idx, 41)

        if rolling_main_layout and pre_date_column_count == 18 and len(df.columns) > 17:
            number_format = workbook.add_format({'num_format': '#,##0'})
            worksheet.set_column(17, 17, 11, number_format)

        if column_width_overrides:
            integer_format = workbook.add_format({'num_format': '#,##0'})
            for col_idx, width in column_width_overrides.items():
                worksheet.set_column(int(col_idx), int(col_idx), width, integer_format)

        if column_format_overrides:
            for col_idx, override in column_format_overrides.items():
                fmt = workbook.add_format(override["format"])
                width = override.get("width")
                worksheet.set_column(int(col_idx), int(col_idx), width, fmt)

            print(f"Excel saved to: {output_path}")
    except PermissionError as exc:
        raise PermissionError(
            f"Could not write Excel file because it is locked or not writable: {output_path}. "
            "If the workbook is open in Excel, close it and run the process again."
        ) from exc
