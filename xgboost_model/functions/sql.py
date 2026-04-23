from sqlalchemy import create_engine

def get_connection():
    engine = create_engine('mssql+pyodbc://sql-2-db/CBQ2?driver=SQL+Server')
    return engine

def read_sql_query(file_path):
    with open(file_path, 'r') as file:
        return file.read()