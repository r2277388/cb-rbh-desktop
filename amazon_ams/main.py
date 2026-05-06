import sys
import time
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from processor import run_full_rebuild

pd.set_option("future.no_silent_downcasting", True)


def main():
    pd.reset_option("display.max_columns")
    start_time = time.time()

    _, errors, archived = run_full_rebuild()

    if archived:
        print("Archived previous outputs:")
        for path in archived:
            print(f"- {path}")

    end_time = time.time()
    print(f"Finished in {end_time - start_time:.2f} seconds.")

    if errors:
        print("Some months failed during processing. See processing_errors.log.")


if __name__ == "__main__":
    main()
