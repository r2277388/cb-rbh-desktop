from __future__ import annotations

import argparse

from reprint_forecast.config import load_config
from reprint_forecast.pipelines.run_pipeline import refresh_raw_data, run_full_pipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Reprint forecasting pipeline")
    sub = parser.add_subparsers(dest="command", required=True)

    p_refresh = sub.add_parser("refresh-data", help="Run SQL and refresh cached PKLs")
    p_refresh.add_argument("--config", default="config.yaml")
    p_refresh.add_argument("--force", action="store_true", help="Ignore existing cache")

    p_run = sub.add_parser("run", help="Train/evaluate and generate 6-month forecast")
    p_run.add_argument("--config", default="config.yaml")
    p_run.add_argument("--refresh", action="store_true", help="Refresh SQL before running")

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    cfg = load_config(args.config)

    if args.command == "refresh-data":
        refresh_raw_data(cfg, force=args.force)
        return

    if args.command == "run":
        run_full_pipeline(cfg, refresh_first=args.refresh)
        return

    raise ValueError(f"Unknown command: {args.command}")


if __name__ == "__main__":
    main()
