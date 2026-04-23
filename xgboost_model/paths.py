from pathlib import Path

ROOT_FOLDER = Path(__file__).parent
DATAWAREHOUSE_ROOT = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier XGBoost")
DATAWAREHOUSE_RAW_DATA_DIR = DATAWAREHOUSE_ROOT / "save_raw_data"

def get_pickle_path():
    if DATAWAREHOUSE_PICKLE_PATH.exists():
        return DATAWAREHOUSE_PICKLE_PATH
    return LOCAL_PICKLE_PATH


def get_parquet_path():
    if DATAWAREHOUSE_PARQUET_PATH.exists():
        return DATAWAREHOUSE_PARQUET_PATH
    return LOCAL_PARQUET_PATH


LOCAL_PICKLE_PATH = ROOT_FOLDER / "save_raw_data" / "df_pickle.pkl"
DATAWAREHOUSE_PICKLE_PATH = DATAWAREHOUSE_RAW_DATA_DIR / "df_pickle.pkl"

LOCAL_PARQUET_PATH = ROOT_FOLDER / "save_raw_data" / "df_parquet.parquet"
DATAWAREHOUSE_PARQUET_PATH = DATAWAREHOUSE_RAW_DATA_DIR / "df_parquet.parquet"


def main():
    print(f"Root folder: {ROOT_FOLDER}")
    print(f"DataWarehouse root: {DATAWAREHOUSE_ROOT}")
    print()
    print(f"Local pickle path: {LOCAL_PICKLE_PATH}")
    print(f"DataWarehouse pickle path: {DATAWAREHOUSE_PICKLE_PATH}")
    print(f"Selected pickle path: {get_pickle_path()}")
    print()
    print(f"Local parquet path: {LOCAL_PARQUET_PATH}")
    print(f"DataWarehouse parquet path: {DATAWAREHOUSE_PARQUET_PATH}")
    print(f"Selected parquet path: {get_parquet_path()}")


if __name__ == "__main__":
    main()
