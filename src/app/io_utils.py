from __future__ import annotations
from pathlib import Path
import re
import unicodedata
import pandas as pd

COLS_REQUIRED = [
    "Nom",
    "Prénom",
    "Date_de_naissance",
    "Expire_le",
    "Email",
]

COLS_OPTIONAL = [
    "Montant",
    "ErreurValide",
]

COLS_WANTED = COLS_REQUIRED + COLS_OPTIONAL


_HEADER_MAP = {
    "nom": "Nom",
    "nom de famille": "Nom",
    "nom usage": "Nom",
    "prenom": "Prénom",
    "prénom": "Prénom",
    "prenom usage": "Prénom",
    "prenom usuel": "Prénom",
    "date de naissance": "Date_de_naissance",
    "date naissance": "Date_de_naissance",
    "date_naissance": "Date_de_naissance",
    "ddn": "Date_de_naissance",
    "expire le": "Expire_le",
    "date de fin": "Expire_le",
    "date fin": "Expire_le",
    "date expiration": "Expire_le",
    "expiration": "Expire_le",
    "date limite": "Expire_le",
    "adresse mail": "Email",
    "adresse email": "Email",
    "email": "Email",
    "courriel": "Email",
    "montant": "Montant",
    "montant regle": "Montant",
    "montant payé": "Montant",
    "montant paye": "Montant",
    "montant verse": "Montant",
    "erreur valide": "ErreurValide",
    "erreur validée": "ErreurValide",
    "erreur valider": "ErreurValide",
    "validation": "ErreurValide",
}


def _normalize_header_label(value: object) -> str:
    text = strip_accents(str(value or ""))
    text = text.replace("\u00a0", " ")
    normalized = re.sub(r"[^0-9A-Za-z]+", " ", text).strip().lower()
    return re.sub(r"\s+", " ", normalized)


def strip_accents(text: str) -> str:
    text = str(text or "")
    normalized = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")


def normalize_name(text: str) -> str:
    normalized = strip_accents(text)
    # Collapse spaces around hyphens so that "Jean - Pierre" matches "Jean-Pierre"
    normalized = re.sub(r"\s*-\s*", "-", normalized)
    # Reduce consecutive whitespace to a single space for consistent comparisons
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip().upper()

def _is_zip(path: Path) -> bool:
    # .xlsx est un ZIP (commence par 'PK\x03\x04')
    with open(path, "rb") as f:
        return f.read(4) == b"PK\x03\x04"

def lire_tableau(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {path}")

    try:
        if path.suffix.lower() in {".xlsx"} or _is_zip(path):
            df = pd.read_excel(path, dtype=str, engine="openpyxl", skiprows=3)
        elif path.suffix.lower() == ".xls":
            df = pd.read_excel(path, dtype=str, engine="xlrd", skiprows=3)
        else:
            df = pd.read_csv(path, sep=",", dtype=str, encoding="utf-8", skiprows=3)
    except Exception as e:
        hint = ""
        if path.suffix.lower() == ".xls":
            hint = " Astuce: essaye de l’ouvrir et de l’enregistrer en .xlsx, ou convertis-le: `soffice --headless --convert-to xlsx fichier.xls`."
        raise RuntimeError(f"Import: échec lecture {path.name} → {e}.{hint}") from e

    column_lookup: dict[str, str] = {}
    for column in df.columns:
        normalized = _normalize_header_label(column)
        if not normalized or normalized in column_lookup:
            continue
        column_lookup[normalized] = column

    rename_map: dict[str, str] = {}
    for normalized, target in _HEADER_MAP.items():
        source = column_lookup.get(normalized)
        if not source:
            continue
        if source == target:
            continue
        if target in df.columns and source != target:
            continue
        rename_map[source] = target

    if rename_map:
        df = df.rename(columns=rename_map)

    missing = [c for c in COLS_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(
            f"Colonnes manquantes dans {path.name}: {missing}. "
            f"Colonnes disponibles: {list(df.columns)}"
        )

    for optional_col in COLS_OPTIONAL:
        if optional_col not in df.columns:
            df[optional_col] = ""

    for col in ("Nom", "Prénom"):
        df[col] = df[col].fillna("").map(normalize_name)

    return df[COLS_WANTED].fillna("")
