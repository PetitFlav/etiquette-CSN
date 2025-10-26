from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
import re
from pathlib import Path
from typing import Iterable, Mapping

import pandas as pd

from .config import PREINSCRIPTION_DIR, VALIDATION_EXPORT_DIR
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


def find_latest_preinscription_export(directory: Path | str | None = None) -> Path | None:
    """Return the most recent pre-inscription CSV located in ``directory``."""

    folder = Path(directory or PREINSCRIPTION_DIR)
    if not folder.exists():
        return None

    candidates = [
        path
        for pattern in ("*.csv", "*.CSV")
        for path in folder.glob(pattern)
        if path.is_file()
    ]
    if not candidates:
        return None

    return max(candidates, key=lambda item: item.stat().st_mtime)


def _normalize_column_name(value: str) -> str:
    return strip_accents((value or "").strip()).lower()


def _parse_amount(raw: object) -> Decimal | None:
    text = str(raw or "").strip()
    if not text:
        return Decimal("0")
    cleaned = (
        text.replace("€", "")
        .replace(" ", "")
        .replace("\u00A0", "")
        .replace("\u202F", "")
        .replace(",", ".")
    )
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def _format_amount(amount: Decimal) -> str:
    return f"{amount.quantize(Decimal('0.01'))}"


def _load_preinscription_lookup(path: Path | str) -> dict[tuple[str, str], Decimal]:
    source = Path(path)
    if not source.exists():
        return {}

    try:
        df = pd.read_csv(
            source,
            sep=None,
            engine="python",
            dtype=str,
            encoding="utf-8",
        )
    except UnicodeDecodeError:
        df = pd.read_csv(
            source,
            sep=None,
            engine="python",
            dtype=str,
            encoding="latin-1",
        )
    except Exception:
        return {}

    if df.empty:
        return {}

    df = df.fillna("")
    column_lookup = {_normalize_column_name(column): column for column in df.columns}

    nom_column = column_lookup.get("nom adherent")
    prenom_column = column_lookup.get("prenom adherent")
    montant_column = column_lookup.get("montant tarif")

    if not (nom_column and prenom_column and montant_column):
        return {}

    lookup: dict[tuple[str, str], Decimal] = {}
    for _, row in df.iterrows():
        nom = normalize_name(str(row.get(nom_column, "")))
        prenom = normalize_name(str(row.get(prenom_column, "")))
        if not nom and not prenom:
            continue

        montant = _parse_amount(row.get(montant_column, ""))
        if montant is None or not montant:
            continue

        key = (nom, prenom)
        amount = montant.quantize(Decimal("0.01"))
        if key in lookup:
            lookup[key] = (lookup[key] + amount).quantize(Decimal("0.01"))
        else:
            lookup[key] = amount

    return lookup


def load_validation_export(
    path: Path | str,
    *,
    preinscriptions_dir: Path | str | None = None,
) -> list[dict[str, str]]:
    """Load a previously exported validation CSV and return its rows."""

    source = Path(path)
    if not source.exists():
        return []

    try:
        df = pd.read_csv(source, sep=";", dtype=str, encoding="utf-8")
    except Exception as exc:  # pragma: no cover - surface informative error
        raise RuntimeError(f"Import validation: échec lecture {source.name} → {exc}") from exc

    for column in ("nom", "prenom", "valide_par", "montant"):
        if column not in df.columns:
            df[column] = ""

    df = df.fillna("")
    records = df[["nom", "prenom", "valide_par", "montant"]].to_dict(orient="records")

    preinscription_lookup: dict[tuple[str, str], Decimal] = {}
    latest_preinscription = find_latest_preinscription_export(preinscriptions_dir)
    if latest_preinscription:
        preinscription_lookup = _load_preinscription_lookup(latest_preinscription)

    if preinscription_lookup:
        for record in records:
            nom = normalize_name(str(record.get("nom", "")))
            prenom = normalize_name(str(record.get("prenom", "")))
            if not nom and not prenom:
                continue

            supplement = preinscription_lookup.get((nom, prenom))
            if not supplement:
                continue

            base_amount = _parse_amount(record.get("montant"))
            if base_amount is None:
                continue

            total = (base_amount + supplement).quantize(Decimal("0.01"))
            record["montant"] = _format_amount(total)

    return records


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
            "montant": str(row.get("montant", "")),
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


def compute_validation_status(
    key: tuple[str, str],
    db_lookup: Mapping[tuple[str, str], DbExpiration],
    validation_lookup: Mapping[tuple[str, str], Mapping[str, str]],
    default_expire: str,
    validators: set[str],
) -> str:
    """Return the status to display for ``key`` according to the ruleset."""

    # ``validators`` is kept for API compatibility although it no longer affects the ruleset.

    if not key or not any(key):
        return ""

    db_entry = db_lookup.get(key)
    validation_entry = validation_lookup.get(key)

    if not db_entry and not validation_entry:
        return ""

    if validation_entry and not db_entry:
        return ""

    if db_entry and not validation_entry:
        return "red"

    if not db_entry or not validation_entry:
        return ""

    cfg_expire = _normalize_expire(default_expire)
    db_expire = _normalize_expire(db_entry.expire)

    if cfg_expire == db_expire:
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
    "find_latest_preinscription_export",
    "find_latest_validation_export",
    "load_latest_expiration_by_person",
    "load_validation_export",
    "parse_validator_names",
]

