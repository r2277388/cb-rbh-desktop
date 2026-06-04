from pathlib import Path
import sys


REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from monthly_campaigns import main as run_monthly_campaigns


def main():
    run_monthly_campaigns()


if __name__ == "__main__":
    main()
