from __future__ import annotations
from pathlib import Path
import unicodedata
import pandas as pd

COLS_REQUIRED = [
    "Nom",
    "Prénom",
    "Date_de_naissance",
    "Expire_le",
    "Email",
COLS_OPTIONAL = [
    "Montant",
    "ErreurValide",
]
COLS_WANTED = COLS_REQUIRED + COLS_OPTIONAL

def strip_accents(text: str) -> str:
    text = str(text or "")
    normalized = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")


def normalize_name(text: str) -> str:
    return strip_accents(text).strip().upper()

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

    mapping = {
        "Date de naissance": "Date_de_naissance",
        "Expire le": "Expire_le",
        "Adresse mail": "Email",
    }
    df = df.rename(columns=mapping)

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
