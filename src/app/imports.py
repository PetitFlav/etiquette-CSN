from __future__ import annotations

import csv
import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import pandas as pd

from .config import DB_PATH, LAST_IMPORT_DIR, LAST_IMPORT_METADATA
from .db import connect, record_print
try:
    from .io_utils import lire_tableau, normalize_name, strip_accents
except ModuleNotFoundError:  # pragma: no cover - pour l'exécutable PyInstaller
    try:
        from io_utils import lire_tableau, normalize_name, strip_accents  # type: ignore
    except ModuleNotFoundError:  # pragma: no cover - repli final
        from app.io_utils import lire_tableau, normalize_name, strip_accents  # type: ignore


Row = dict[str, object]
LookupKey = tuple[str, str]


def build_ddn_lookup_from_rows(rows: Iterable[Row]) -> dict[LookupKey, str | None]:
    """Build a lookup ``(nom_lower, prenom_lower) -> ddn`` from imported rows."""
    tmp: dict[LookupKey, set[str]] = {}
    for r in rows or []:
        nom_normalized = normalize_name(str(r.get("Nom") or ""))
        prenom_normalized = normalize_name(str(r.get("Prénom") or ""))
        nom = nom_normalized.lower()
        prenom = prenom_normalized.lower()
        ddn = (str(r.get("Date_de_naissance") or "").strip())
        if not nom and not prenom:
            continue
        key = (nom, prenom)
        tmp.setdefault(key, set())
        if ddn:
            tmp[key].add(ddn)

    out: dict[LookupKey, str | None] = {}
    for key, values in tmp.items():
        if len(values) == 1:
            out[key] = next(iter(values))
        else:
            out[key] = None
    return out


def import_already_printed_csv(
    csv_path: Path,
    expire: str,
    rows_ddn_lookup: dict[LookupKey, str | None] | None = None,
    *,
    db_path: Path | None = None,
    connect_fn: Callable[[Path], Any] = connect,
) -> tuple[int, int]:
    """Import a ``nom;prenom`` CSV and log entries as already printed."""
    if not csv_path.exists():
        return (0, 0)

    imported, skipped = 0, 0
    db_target = db_path or DB_PATH
    with csv_path.open("r", encoding="utf-8", newline="") as fh:
        reader = csv.reader(fh, delimiter=";")
        with connect_fn(db_target) as cn:  # type: ignore[call-arg]
            for nom, prenom in reader:
                nom = normalize_name((nom or "").strip())
                prenom = normalize_name((prenom or "").strip())
                if not nom and not prenom:
                    skipped += 1
                    continue

                key = (nom.lower(), prenom.lower())

                ddns = cn.execute(
                    """
                    SELECT DISTINCT ddn
                    FROM prints
                    WHERE LOWER(nom)=? AND LOWER(prenom)=?
                    """,
                    key,
                ).fetchall()
                ddn_candidates = [
                    (row[0] or "").strip()
                    for row in ddns
                    if (row[0] or "").strip()
                ]
                if len(ddn_candidates) == 1:
                    ddn = ddn_candidates[0]
                elif len(ddn_candidates) > 1:
                    ddn = ""
                else:
                    ddn = (rows_ddn_lookup.get(key) or "") if rows_ddn_lookup else ""
                    ddn = ddn or ""

                record_print(cn, nom, prenom, ddn=ddn, expire=expire, zpl=None, status="printed")
                imported += 1

    return (imported, skipped)


def _normalize_header_label(label: object) -> str:
    raw = str(label or "").strip()
    if not raw:
        return ""
    normalized = strip_accents(raw).lower()
    normalized = re.sub(r"[^a-z0-9]+", " ", normalized)
    return re.sub(r"\s+", " ", normalized).strip()


_VALIDATION_HEADER_MAP: dict[str, str] = {
    "nom": "Nom",
    "nom usage": "Nom",
    "nom de famille": "Nom",
    "prenom": "Prénom",
    "prénom": "Prénom",
    "prenom usuel": "Prénom",
    "prenom usage": "Prénom",
    "date de naissance": "Date_de_naissance",
    "date naissance": "Date_de_naissance",
    "date_naissance": "Date_de_naissance",
    "ddn": "Date_de_naissance",
    "date naissance ddn": "Date_de_naissance",
    "expire le": "Expire_le",
    "date de fin": "Expire_le",
    "date fin": "Expire_le",
    "date expiration": "Expire_le",
    "expiration": "Expire_le",
    "date limite": "Expire_le",
    "courriel": "Email",
    "email": "Email",
    "mail": "Email",
    "adresse mail": "Email",
    "adresse email": "Email",
    "montant": "Montant",
    "montant regle": "Montant",
    "montant payé": "Montant",
    "montant paye": "Montant",
    "montant verse": "Montant",
    "validation": "ErreurValide",
    "erreur valide": "ErreurValide",
    "erreur validée": "ErreurValide",
    "erreur valider": "ErreurValide",
    "valide": "ErreurValide",
    "erreur ok": "ErreurValide",
}


def _format_validation_date(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        text = value.strip()
        if not text or text.lower() == "nat":
            return ""
        normalized = text.replace("\u00a0", " ").strip()
        normalized_slash = normalized.replace(".", "/").replace("-", "/")
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", normalized_slash):
            parsed = pd.to_datetime(normalized_slash, format="%d/%m/%Y", errors="coerce")
        elif re.fullmatch(r"\d{4}/\d{2}/\d{2}", normalized_slash):
            parsed = pd.to_datetime(normalized_slash, format="%Y/%m/%d", errors="coerce")
        elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", normalized):
            parsed = pd.to_datetime(normalized, format="%Y-%m-%d", errors="coerce")
        elif re.fullmatch(r"\d{2}-\d{2}-\d{4}", normalized):
            parsed = pd.to_datetime(normalized, format="%d-%m-%Y", errors="coerce")
        else:
            try:
                parsed = pd.to_datetime(normalized, dayfirst=True, errors="coerce")
            except Exception:
                parsed = pd.NaT
        if pd.isna(parsed):
            return text
        return parsed.strftime("%d/%m/%Y")
    try:
        parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    except Exception:
        parsed = pd.NaT
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%d/%m/%Y")


def parse_validation_workbook(path: Path | str) -> list[Row]:
    """Read a validation Excel workbook and return normalized rows.

    The returned rows follow the same structure as :func:`lire_tableau` with the
    columns ``Nom``, ``Prénom``, ``Date_de_naissance``, ``Expire_le``, ``Email``,
    ``Montant`` et ``ErreurValide``. Header names in the workbook are matched in a
    permissive fashion (accents/spacing ignored) so that slightly different
    source files can be imported without manual tweaks.
    """

    source = Path(path)
    if not source.exists():
        raise FileNotFoundError(f"Fichier introuvable: {source}")

    try:
        df = pd.read_excel(source, dtype=object)
    except ValueError:
        # Older .xls files sometimes require the xlrd engine.
        df = pd.read_excel(source, dtype=object, engine="xlrd")
    except Exception as exc:  # pragma: no cover - depends on pandas backends
        raise RuntimeError(f"Import validation: échec lecture {source.name} → {exc}") from exc

    if df.empty:
        return []

    column_lookup: dict[str, str] = {}
    for column in df.columns:
        normalized = _normalize_header_label(column)
        if not normalized:
            continue
        target = _VALIDATION_HEADER_MAP.get(normalized)
        if target and target not in column_lookup:
            column_lookup[target] = column

    required = ["Nom", "Prénom"]
    missing = [col for col in required if col not in column_lookup]
    if missing:
        raise ValueError(
            f"Colonnes manquantes dans {source.name}: {missing}. "
            f"Colonnes disponibles: {list(df.columns)}"
        )

    selected_columns: dict[str, str] = {target: column_lookup[target] for target in column_lookup}
    df = df.rename(columns={orig: target for target, orig in selected_columns.items()})
    df = df[list(selected_columns.keys())]

    # Ensure the optional columns exist so that downstream logic receives a
    # consistent schema.
    for optional in ("Date_de_naissance", "Expire_le", "Email", "Montant", "ErreurValide"):
        if optional not in df.columns:
            df[optional] = ""

    df = df[[
        "Nom",
        "Prénom",
        "Date_de_naissance",
        "Expire_le",
        "Email",
        "Montant",
        "ErreurValide",
    ]]

    df["Nom"] = df["Nom"].map(_to_clean_string).map(normalize_name)
    df["Prénom"] = df["Prénom"].map(_to_clean_string).map(normalize_name)
    df["Date_de_naissance"] = df["Date_de_naissance"].map(_format_validation_date)
    df["Expire_le"] = df["Expire_le"].map(_format_validation_date)
    df["Email"] = df["Email"].map(_to_clean_string)
    df["Montant"] = df["Montant"].map(_to_clean_string)
    df["ErreurValide"] = df["ErreurValide"].map(_to_clean_string)

    df = df[(df["Nom"].str.strip() != "") | (df["Prénom"].str.strip() != "")]

    return df.to_dict(orient="records")


def apply_validation_updates(rows: Sequence[Row], updates: Sequence[Row]) -> tuple[list[Row], int, int]:
    """Merge validation ``updates`` into ``rows``.

    ``rows`` is typically the dataset currently loaded in the GUI. The function
    returns the updated list along with the number of modified rows and the
    number of new entries that were added.
    """

    rows_list: list[Row] = [dict(row) for row in (rows or [])]
    updated, added = 0, 0

    def _norm(value: object) -> str:
        return normalize_name(str(value or ""))

    for update in updates or []:
        nom = _norm(update.get("Nom"))
        prenom = _norm(update.get("Prénom"))
        if not nom and not prenom:
            continue

        ddn_update = str(update.get("Date_de_naissance") or "").strip()

        target_index = None
        for idx, row in enumerate(rows_list):
            if _norm(row.get("Nom")) != nom:
                continue
            if _norm(row.get("Prénom")) != prenom:
                continue
            ddn_existing = str(row.get("Date_de_naissance") or "").strip()
            if ddn_update and ddn_existing and ddn_existing != ddn_update:
                continue
            target_index = idx
            break

        if target_index is None:
            new_row: Row = {
                "Nom": nom,
                "Prénom": prenom,
                "Date_de_naissance": ddn_update,
                "Expire_le": str(update.get("Expire_le") or "").strip(),
                "Email": str(update.get("Email") or "").strip(),
                "Montant": str(update.get("Montant") or "").strip(),
                "ErreurValide": str(update.get("ErreurValide") or "").strip(),
                "Derniere": "",
                "Compteur": 0,
            }
            rows_list.append(new_row)
            added += 1
            continue

        row = rows_list[target_index]
        changed = False

        if row.get("Nom") != nom:
            row["Nom"] = nom
            changed = True
        if row.get("Prénom") != prenom:
            row["Prénom"] = prenom
            changed = True

        for key in ("Date_de_naissance", "Expire_le", "Email", "Montant", "ErreurValide"):
            value = update.get(key)
            if value is None:
                continue
            value_str = str(value).strip()
            if not value_str:
                continue
            if str(row.get(key) or "").strip() != value_str:
                row[key] = value_str
                changed = True

        if changed:
            updated += 1

    return rows_list, updated, added


def _to_clean_string(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):  # type: ignore[arg-type]
            return ""
    except TypeError:
        pass
    return str(value).strip()


def persist_last_import(source: Path) -> dict[str, object]:
    """Cache the latest imported file and return the stored metadata."""

    if not source.exists():
        raise FileNotFoundError(f"Fichier introuvable: {source}")

    LAST_IMPORT_DIR.mkdir(parents=True, exist_ok=True)
    target = LAST_IMPORT_DIR / source.name
    shutil.copy2(source, target)

    metadata: dict[str, object] = {
        "source_path": str(source.resolve()),
        "source_name": source.name,
        "cached_path": str(target.resolve()),
        "cached_name": target.name,
        "stored_at": datetime.now().isoformat(timespec="seconds"),
    }

    LAST_IMPORT_METADATA.parent.mkdir(parents=True, exist_ok=True)
    LAST_IMPORT_METADATA.write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return metadata


def load_last_import() -> tuple[list[dict], dict[str, object]]:
    """Return cached rows and metadata when a previous import exists."""

    if not LAST_IMPORT_METADATA.exists():
        return ([], {})

    try:
        metadata = json.loads(LAST_IMPORT_METADATA.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return ([], {})

    cached_path = Path(metadata.get("cached_path") or "")
    if not cached_path.exists():
        return ([], metadata)

    df = lire_tableau(cached_path)
    rows = df.to_dict(orient="records")
    metadata.setdefault("source_name", cached_path.name)
    metadata.setdefault("cached_name", cached_path.name)
    return (rows, metadata)


__all__ = [
    "build_ddn_lookup_from_rows",
    "import_already_printed_csv",
    "persist_last_import",
    "load_last_import",
    "parse_validation_workbook",
    "apply_validation_updates",
]
