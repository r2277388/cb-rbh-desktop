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

def save_to_excel(df, filename, summary=None, format_cols=None, decimal_cols=None):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        worksheet_name = 'Sheet1'
        df.to_excel(writer, sheet_name=worksheet_name, startrow=3, index=False)
        worksheet = writer.sheets[worksheet_name]
        workbook = writer.book
        last_data_row = len(df) + 3
        last_col = len(df.columns) - 1

        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D8E4BC',
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
            'num_format': '#,##0',
            'align': 'right',
            'valign': 'vcenter',
            'bg_color': '#E4DFEC',
            'border': 1,
        })
        summary_decimal_format = workbook.add_format({
            'bold': True,
            'num_format': '#,##0.00',
            'align': 'right',
            'valign': 'vcenter',
            'bg_color': '#E4DFEC',
            'border': 1,
        })
        blank_summary_format = workbook.add_format({
            'bg_color': '#E4DFEC',
            'border': 1,
        })

        for col_idx, col_name in enumerate(df.columns):
            worksheet.write(3, col_idx, col_name, header_format)

        # Add total and subtotal rows above the header row.
        if summary:
            label_col_idx = 7 if len(df.columns) > 7 else 0
            for row_idx in (0, 1):
                for col_idx in range(label_col_idx, len(df.columns)):
                    worksheet.write_blank(row_idx, col_idx, None, blank_summary_format)

            worksheet.write(0, label_col_idx, 'Total', summary_label_format)
            worksheet.write(1, label_col_idx, 'Subtotal', summary_label_format)

            for col_idx, col in enumerate(df.columns):
                if col not in summary or col_idx == label_col_idx:
                    continue

                start_cell = xl_rowcol_to_cell(4, col_idx)
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
        worksheet.autofilter(3, 0, last_data_row, last_col)
        worksheet.freeze_panes(4, 0)

        # Format integer columns
        if format_cols:
            number_format = workbook.add_format({'num_format': '#,##0'})
            for col in format_cols:
                if col in df.columns:
                    col_idx = df.columns.get_loc(col)
                    worksheet.set_column(col_idx, col_idx, None, number_format)

        # Format decimal columns
        if decimal_cols:
            decimal_format = workbook.add_format({'num_format': '#,##0.00'})
            for col in decimal_cols:
                if col in df.columns:
                    col_idx = df.columns.get_loc(col)
                    worksheet.set_column(col_idx, col_idx, None, decimal_format)

        if "ISBN" in df.columns:
            isbn_format = workbook.add_format({'num_format': '0'})
            col_idx = df.columns.get_loc("ISBN")
            worksheet.set_column(col_idx, col_idx, None, isbn_format)

        print(f"Excel saved to: {filename}")
