from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv


@dataclass
class AppConfig:
    raw: dict[str, Any]
    root: Path

    @property
    def seed(self) -> int:
        return int(self.raw["project"]["seed"])

    @property
    def output_dir(self) -> Path:
        return self.root / self.raw["project"]["output_dir"]


def load_config(config_path: str | Path) -> AppConfig:
    config_path = Path(config_path).resolve()
    load_dotenv(config_path.parent / ".env")
    with config_path.open("r", encoding="utf-8") as f:
        raw = yaml.safe_load(f)
    return AppConfig(raw=raw, root=config_path.parent)


def get_db_env() -> dict[str, str]:
    return {
        "driver": os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server"),
        "server": os.getenv("DB_SERVER", ""),
        "database": os.getenv("DB_DATABASE", ""),
        "trusted_connection": os.getenv("DB_TRUSTED_CONNECTION", "yes"),
        "uid": os.getenv("DB_UID", ""),
        "pwd": os.getenv("DB_PWD", ""),
    }
