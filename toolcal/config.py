"""Settings stored in config.json next to the program (not committed to git)."""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass

from .recognize import DEFAULT_EFFORT, DEFAULT_MODEL

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(ROOT, "config.json")


@dataclass
class Settings:
    excel_path: str = ""
    anthropic_api_key: str = ""  # empty -> ANTHROPIC_API_KEY environment variable
    model: str = DEFAULT_MODEL
    effort: str = DEFAULT_EFFORT
    archive_photos: bool = True
    archive_dir: str = ""  # empty -> "<excel folder>/검교정사진"
    backup: bool = True
    backup_dir: str = ""  # empty -> "<excel folder>/backup"
    warn_days: int = 30

    def api_key(self) -> str:
        return self.anthropic_api_key or os.environ.get("ANTHROPIC_API_KEY", "")

    def resolved_archive_dir(self, excel_path: str) -> str | None:
        if not self.archive_photos:
            return None
        return self.archive_dir or os.path.join(os.path.dirname(os.path.abspath(excel_path)), "검교정사진")

    def resolved_backup_dir(self, excel_path: str) -> str | None:
        if not self.backup:
            return None
        return self.backup_dir or os.path.join(os.path.dirname(os.path.abspath(excel_path)), "backup")


def load_settings(path: str = CONFIG_PATH) -> Settings:
    if not os.path.exists(path):
        return Settings()
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    known = Settings.__dataclass_fields__
    return Settings(**{k: v for k, v in data.items() if k in known})


def save_settings(settings: Settings, path: str = CONFIG_PATH) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(asdict(settings), fh, ensure_ascii=False, indent=2)
