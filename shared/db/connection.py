from sqlalchemy import create_engine


def get_connection():
    """Return the shared SQL Server engine for CBQ2."""
    return create_engine("mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server")

