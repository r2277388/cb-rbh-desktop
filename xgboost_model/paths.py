from pathlib import Path

ROOT_FOLDER = Path(__file__).parent
DATAWAREHOUSE_ROOT = Path(r"F:\ANALYSIS\Finance\DataWarehouse\Atelier XGBoost")
DATAWAREHOUSE_RAW_DATA_DIR = DATAWAREHOUSE_ROOT / "save_raw_data"
DATAWAREHOUSE_PICKLE_PATH = DATAWAREHOUSE_RAW_DATA_DIR / "df_pickle.pkl"
DATAWAREHOUSE_PARQUET_PATH = DATAWAREHOUSE_RAW_DATA_DIR / "df_parquet.parquet"


def get_pickle_path():
    return DATAWAREHOUSE_PICKLE_PATH


def get_parquet_path():
    return DATAWAREHOUSE_PARQUET_PATH


def main():
    print(f"Root folder: {ROOT_FOLDER}")
    print(f"DataWarehouse root: {DATAWAREHOUSE_ROOT}")
    print()
    print(f"DataWarehouse pickle path: {DATAWAREHOUSE_PICKLE_PATH}")
    print(f"Selected pickle path: {get_pickle_path()}")
    print()
    print(f"DataWarehouse parquet path: {DATAWAREHOUSE_PARQUET_PATH}")
    print(f"Selected parquet path: {get_parquet_path()}")


if __name__ == "__main__":
    main()
