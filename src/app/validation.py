from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import re
from pathlib import Path
from typing import Iterable, Mapping

import pandas as pd

from .config import VALIDATION_EXPORT_DIR
from .io_utils import normalize_name, strip_accents


@dataclass(slots=True)
class DbExpiration:
    """Represent the latest expiration stored for a person in the DB."""

    expire: str
    printed_at: str


def find_latest_validation_export(directory: Path | str = VALIDATION_EXPORT_DIR) -> Path | None:
    """Return the most recent validation CSV located in ``directory``."""

    folder = Path(directory)
    if not folder.exists():
        return None

    candidates = [
        path
        for path in folder.glob("*_validation.csv")
        if path.is_file()
    ]
    if not candidates:
        return None

    return max(candidates, key=lambda item: item.stat().st_mtime)


def load_validation_export(path: Path | str) -> list[dict[str, str]]:
    """Load a previously exported validation CSV and return its rows."""

    source = Path(path)
    if not source.exists():
        return []

    try:
        df = pd.read_csv(source, sep=";", dtype=str, encoding="utf-8")
    except Exception as exc:  # pragma: no cover - surface informative error
        raise RuntimeError(f"Import validation: échec lecture {source.name} → {exc}") from exc

    for column in ("nom", "prenom", "valide_par"):
        if column not in df.columns:
            df[column] = ""

    df = df.fillna("")
    return df[["nom", "prenom", "valide_par"]].to_dict(orient="records")


def build_validation_lookup(rows: Iterable[Mapping[str, object]]) -> dict[tuple[str, str], dict[str, str]]:
    """Return a ``(nom, prenom)`` → row lookup from validation ``rows``."""

    lookup: dict[tuple[str, str], dict[str, str]] = {}
    for row in rows or []:
        nom = normalize_name(str(row.get("nom", "")))
        prenom = normalize_name(str(row.get("prenom", "")))
        if not nom and not prenom:
            continue
        lookup[(nom, prenom)] = {
            "nom": nom,
            "prenom": prenom,
            "valide_par": str(row.get("valide_par", "")),
        }
    return lookup


def parse_validator_names(raw: str) -> set[str]:
    """Return the normalized first names defined in ``ffessm_validators``."""

    tokens = re.split(r"[;,/\n\r]+", raw or "")
    cleaned = {strip_accents(token).strip().upper() for token in tokens if token and token.strip()}
    return {token for token in cleaned if token}


def _normalize_expire(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(text, fmt)
        except ValueError:
            continue
        return dt.strftime("%d/%m/%Y")
    return text


def _tokenize(text: str) -> set[str]:
    if not text:
        return set()
    normalized = strip_accents(text).upper()
    return {token for token in re.split(r"[^A-Z0-9]+", normalized) if token}


def _validator_present(valide_par: str, validators: set[str]) -> bool:
    if not validators:
        return False
    tokens = _tokenize(valide_par)
    return any(validator in tokens for validator in validators)


def compute_validation_status(
    key: tuple[str, str],
    db_lookup: Mapping[tuple[str, str], DbExpiration],
    validation_lookup: Mapping[tuple[str, str], Mapping[str, str]],
    default_expire: str,
    validators: set[str],
) -> str:
    """Return the status to display for ``key`` according to the ruleset."""

    if not key or not any(key):
        return ""

    db_entry = db_lookup.get(key)
    if not db_entry:
        return ""

    validation_entry = validation_lookup.get(key)
    if not validation_entry:
        return "question"

    cfg_expire = _normalize_expire(default_expire)
    db_expire = _normalize_expire(db_entry.expire)

    if cfg_expire and db_expire and cfg_expire != db_expire:
        return "red"
    if cfg_expire and not db_expire:
        return "red"

    if not validators:
        return "green"

    if _validator_present(validation_entry.get("valide_par", ""), validators):
        return "green"
    return "orange"


def load_latest_expiration_by_person(conn) -> dict[tuple[str, str], DbExpiration]:
    """Return the latest expiration stored per person from the ``prints`` table."""

    lookup: dict[tuple[str, str], DbExpiration] = {}
    try:
        cur = conn.execute(
            "SELECT nom, prenom, expire, printed_at FROM prints ORDER BY printed_at DESC"
        )
    except Exception:
        return lookup

    for row in cur.fetchall():
        nom = normalize_name(row["nom"] if isinstance(row, Mapping) else row[0])
        prenom = normalize_name(row["prenom"] if isinstance(row, Mapping) else row[1])
        if not nom and not prenom:
            continue
        key = (nom, prenom)
        if key in lookup:
            continue
        expire = (row["expire"] if isinstance(row, Mapping) else row[2]) or ""
        printed_at = (row["printed_at"] if isinstance(row, Mapping) else row[3]) or ""
        lookup[key] = DbExpiration(str(expire).strip(), str(printed_at).strip())
    return lookup


__all__ = [
    "DbExpiration",
    "build_validation_lookup",
    "compute_validation_status",
    "find_latest_validation_export",
    "load_latest_expiration_by_person",
    "load_validation_export",
    "parse_validator_names",
]

