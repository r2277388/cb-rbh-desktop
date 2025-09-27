from sqlalchemy import create_engine
import pandas as pd

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

        # Write summary/totals in row 2 (Excel row index 1)
        if summary:
            for col_idx, col in enumerate(df.columns):
                if col in summary:
                    worksheet.write(1, col_idx, summary[col])

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

        print(f"Excel saved to: {filename}")