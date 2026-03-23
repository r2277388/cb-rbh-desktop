import pandas as pd
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path

def _load_shared_paths():
    shared_path = Path(__file__).resolve().parents[1] / "paths" / "process_paths.py"
    spec = spec_from_file_location("_shared_process_paths", shared_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load shared process paths from {shared_path}")
    module = module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


default_file_path = _load_shared_paths().CHRONICLE_ASIN_MAPPING_FILE

def load_asin_mapping(file_path=default_file_path):
    df = pd.read_excel(
        file_path,
        usecols=['Asin', 'Isbn13'],
        sheet_name='Sheet1',
        header=0,
        engine='openpyxl'
    )
    df.columns = df.columns.str.lower()
    df.rename(columns={'asin': 'ASIN', 'isbn13': 'ISBN'}, inplace=True)
    df['ASIN'] = df['ASIN'].astype(str).str.zfill(10)
    return df

if __name__ == '__main__':
    asin_mapping = load_asin_mapping()
    print(asin_mapping.info())
    print(asin_mapping.head())
    # Save to a pickle file if needed
