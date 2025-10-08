from __future__ import annotations

import configparser
import sys
from pathlib import Path
from typing import Dict


def _resolve_root() -> Path:
    """Return the application root directory.

    When packaged (``sys.frozen``) we rely on the executable location,
    otherwise we walk up from the source tree (``src/app`` -> project root).
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


ROOT = _resolve_root()
SRC_DIR = ROOT / "src"
DATA_DIR = ROOT / "data"
SORTIES_DIR = DATA_DIR / "sorties"
DB_PATH = DATA_DIR / "app.db"
LAST_IMPORT_DIR = DATA_DIR / "last_import"
LAST_IMPORT_METADATA = DATA_DIR / "last_import.json"
DEFAULT_EXPIRATION = "31/12/2026"
CONFIG_PATH = ROOT / "config.ini"


def load_config() -> Dict[str, str]:
    """Load the optional ``config.ini`` file.

    Only the settings used by the GUI are exposed; sensible defaults are
    returned when the file (or individual keys) is missing so that the GUI can
    operate out of the box.
    """
    cfg = configparser.ConfigParser()
    if CONFIG_PATH.exists():
        cfg.read(CONFIG_PATH, encoding="utf-8")

    return {
        "backend": cfg.get("impression", "backend", fallback="win32print"),
        "device": cfg.get("impression", "device", fallback="Brother QL-570"),
        "label": cfg.get("impression", "label", fallback="62"),
        "default_expire": cfg.get("app", "default_expire", fallback=""),
        "auto_import_file": cfg.get("app", "auto_import_file", fallback="deja_imprimes.csv"),
        "splash_image": cfg.get("app", "splash_image", fallback=""),
        "show_reset_db_button": cfg.get("app", "show_reset_db_button", fallback="false"),
        "rotate": cfg.get("impression", "rotate", fallback="0"),
    }


__all__ = [
    "ROOT",
    "SRC_DIR",
    "DATA_DIR",
    "SORTIES_DIR",
    "DB_PATH",
    "LAST_IMPORT_DIR",
    "LAST_IMPORT_METADATA",
    "DEFAULT_EXPIRATION",
    "CONFIG_PATH",
    "load_config",
]
