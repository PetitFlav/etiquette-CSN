from __future__ import annotations

import csv
import json
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
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


@dataclass(slots=True)
class ValidationParseResult:
    rows: list[Row]
    export_path: Path | None = None


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


def _extract_amount(value: object) -> str:
    raw = _to_clean_string(value)
    if not raw:
        return ""

    normalized = raw.replace("\u00a0", " ").replace("€", "").replace(",", ".")
    match = re.search(r"(\d+(?:[\s\.]\d{3})*(?:\.\d+)?)", normalized)
    if not match:
        return ""

    number = match.group(1).replace(" ", "")
    try:
        value_decimal = Decimal(number)
    except InvalidOperation:
        return number

    return f"{value_decimal.quantize(Decimal('0.01'))}"


def _looks_like_member_name(cell: str) -> bool:
    text = _to_clean_string(cell)
    if not text:
        return False

    normalized = strip_accents(text).upper()
    if normalized in {"NOM", "PRENOM", "PRÉNOM", "MEMBRE", "STATUT"}:
        return False
    if normalized.startswith("TOTAL"):
        return False
    if re.search(r"\badhesion\b", normalized, re.IGNORECASE):
        return False
    if re.match(r"^\d", text):
        return False
    tokens = [tok for tok in re.split(r"\s+", text.strip()) if tok]
    if len(tokens) < 2:
        return False
    has_upper = any(strip_accents(tok).upper() == strip_accents(tok) for tok in tokens[:2])
    return has_upper and re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", text) is not None


_NAME_CONNECTORS = {
    "DE",
    "DES",
    "DU",
    "D",
    "LE",
    "LA",
    "LES",
    "L",
    "ST",
    "STE",
    "SAINT",
    "SAINTE",
}


def _split_name_parts(text: str) -> tuple[str, str]:
    cleaned = _to_clean_string(text)
    if not cleaned:
        return ("", "")

    first_line = cleaned.splitlines()[0].strip()
    tokens = [tok for tok in re.split(r"\s+", first_line) if tok]
    if not tokens:
        return ("", "")

    surname_tokens: list[str] = []
    first_name_tokens: list[str] = []

    for idx, token in enumerate(tokens):
        ascii_token = strip_accents(token).replace("'", "").upper()
        if not first_name_tokens and (
            (any(ch.isupper() for ch in token if ch.isalpha()) and not any(ch.islower() for ch in token if ch.isalpha()))
            or ascii_token in _NAME_CONNECTORS
        ):
            surname_tokens.append(token)
            continue

        first_name_tokens = tokens[idx:]
        break

    if not first_name_tokens and surname_tokens:
        first_name_tokens = [surname_tokens.pop()]

    surname = normalize_name(" ".join(surname_tokens).strip()) if surname_tokens else ""
    firstname = normalize_name(" ".join(first_name_tokens).strip()) if first_name_tokens else ""
    return (surname, firstname)


def _extract_confirmation_date(text: str) -> str:
    cleaned = _to_clean_string(text)
    if not cleaned:
        return ""

    normalized = cleaned.replace("\u00a0", " ")
    patterns = [
        (r"\d{2}/\d{2}/\d{4}", "%d/%m/%Y"),
        (r"\d{2}-\d{2}-\d{4}", "%d-%m-%Y"),
        (r"\d{4}-\d{2}-\d{2}", "%Y-%m-%d"),
        (r"\d{4}/\d{2}/\d{2}", "%Y/%m/%d"),
    ]
    for pattern, fmt in patterns:
        match = re.search(pattern, normalized)
        if not match:
            continue
        candidate = match.group(0)
        parsed = pd.to_datetime(candidate, format=fmt, errors="coerce")
        if pd.isna(parsed):
            continue
        return parsed.strftime("%d/%m/%Y")

    try:
        parsed = pd.to_datetime(normalized, dayfirst=True, errors="coerce")
    except Exception:
        parsed = pd.NaT
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%d/%m/%Y")


def _extract_confirmation_person(text: str) -> str:
    cleaned = _to_clean_string(text)
    if not cleaned:
        return ""

    match = re.search(r"par\s+([^\d,;]+)", cleaned, flags=re.IGNORECASE)
    if not match:
        return ""

    person = match.group(1)
    person = re.split(r"\b(valide|validée|validé|non fourni)\b", person, flags=re.IGNORECASE)[0]
    person = person.split("(")[0]
    person = re.split(r"\s{2,}", person)[0]
    person = person.replace("\u00a0", " ")
    person = re.sub(r"[,;].*", "", person).strip()
    return normalize_name(person)


def _parse_validation_columnar(df: pd.DataFrame) -> list[Row]:
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
        available = list(df.columns)
        raise ValueError(
            f"Colonnes manquantes: {missing}. Colonnes disponibles: {available}"
        )

    selected_columns: dict[str, str] = {target: column_lookup[target] for target in column_lookup}
    df = df.rename(columns={orig: target for target, orig in selected_columns.items()})
    df = df[list(selected_columns.keys())]

    for optional in ("Date_de_naissance", "Expire_le", "Email", "Montant", "ErreurValide"):
        if optional not in df.columns:
            df[optional] = ""

    df = df[
        [
            "Nom",
            "Prénom",
            "Date_de_naissance",
            "Expire_le",
            "Email",
            "Montant",
            "ErreurValide",
        ]
    ]

    df["Nom"] = df["Nom"].map(_to_clean_string).map(normalize_name)
    df["Prénom"] = df["Prénom"].map(_to_clean_string).map(normalize_name)
    df["Date_de_naissance"] = df["Date_de_naissance"].map(_format_validation_date)
    df["Expire_le"] = df["Expire_le"].map(_format_validation_date)
    df["Email"] = df["Email"].map(_to_clean_string)
    df["Montant"] = df["Montant"].map(_to_clean_string)
    df["ErreurValide"] = df["ErreurValide"].map(_to_clean_string)

    df = df[(df["Nom"].str.strip() != "") | (df["Prénom"].str.strip() != "")]

    rows = df.to_dict(orient="records")
    for row in rows:
        row.setdefault("Validation_confirmee_le", "")
        row.setdefault("Validation_confirmee_par", "")
    return rows


def _build_row_from_block(block: list[list[str]]) -> Row | None:
    cleaned_block = [line for line in block if any(line)]
    if not cleaned_block:
        return None

    trimmed = cleaned_block[:3]
    name_line = trimmed[0][0] if trimmed and trimmed[0] else ""
    nom, prenom = _split_name_parts(name_line)

    montant_cell = ""
    if len(trimmed) >= 3 and len(trimmed[2]) > 5:
        montant_cell = trimmed[2][5]
    montant = _extract_amount(montant_cell)

    confirm_cells = [line[1] for line in trimmed if len(line) > 1 and line[1]]
    confirm_text = " ".join(confirm_cells)
    confirmation_date = _extract_confirmation_date(confirm_text)
    confirmation_person = _extract_confirmation_person(confirm_text)

    if not any([nom, prenom, montant, confirmation_date, confirmation_person]):
        return None

    return {
        "Nom": nom,
        "Prénom": prenom,
        "Date_de_naissance": "",
        "Expire_le": "",
        "Email": "",
        "Montant": montant,
        "ErreurValide": "",
        "Validation_confirmee_le": confirmation_date,
        "Validation_confirmee_par": confirmation_person,
    }


def _parse_validation_multiline(df: pd.DataFrame) -> list[Row]:
    if df.empty:
        return []

    df = df.fillna("")
    rows: list[Row] = []
    current_block: list[list[str]] = []

    def _flush(block: list[list[str]]) -> None:
        if not block:
            return
        trimmed_block = block[:3]
        row = _build_row_from_block(trimmed_block)
        if row:
            rows.append(row)

    for _, series in df.iterrows():
        values = [_to_clean_string(val) for val in series.tolist()]
        if not any(values):
            _flush(current_block)
            current_block = []
            continue

        first_cell = values[0] if values else ""
        if _looks_like_member_name(first_cell):
            _flush(current_block)
            current_block = [values]
        else:
            if current_block:
                current_block.append(values)

    _flush(current_block)
    return rows


def _write_validation_export(rows: list[Row], *, export_dir: Path, now: datetime) -> Path:
    export_dir.mkdir(parents=True, exist_ok=True)
    timestamp = now.strftime("%Y%m%d-%H%M%S")
    export_path = export_dir / f"FichierValidation_{timestamp}.xlsx"

    desired_columns = [
        "Nom",
        "Prénom",
        "Date_de_naissance",
        "Expire_le",
        "Email",
        "Montant",
        "ErreurValide",
        "Validation_confirmee_le",
        "Validation_confirmee_par",
    ]

    df = pd.DataFrame(rows)
    for col in desired_columns:
        if col not in df.columns:
            df[col] = ""
    df = df[desired_columns]
    df.to_excel(export_path, index=False)
    return export_path


def parse_validation_workbook(
    path: Path | str,
    *,
    export_dir: Path | None = None,
    now: datetime | None = None,
) -> ValidationParseResult:
    """Read a validation workbook and normalise its content."""

    source = Path(path)
    if not source.exists():
        raise FileNotFoundError(f"Fichier introuvable: {source}")

    try:
        df = pd.read_excel(source, dtype=object)
    except ValueError:
        df = pd.read_excel(source, dtype=object, engine="xlrd")
    except Exception as exc:  # pragma: no cover - depends on pandas backends
        raise RuntimeError(f"Import validation: échec lecture {source.name} → {exc}") from exc

    try:
        rows = _parse_validation_columnar(df)
    except ValueError:
        try:
            df_raw = pd.read_excel(source, header=None, dtype=object)
        except ValueError:
            df_raw = pd.read_excel(source, header=None, dtype=object, engine="xlrd")
        rows = _parse_validation_multiline(df_raw)

    if not rows:
        return ValidationParseResult([], None)

    export_path = _write_validation_export(
        rows,
        export_dir=export_dir or source.parent,
        now=now or datetime.now(),
    )

    return ValidationParseResult(rows, export_path)


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
                "Validation_confirmee_le": str(update.get("Validation_confirmee_le") or "").strip(),
                "Validation_confirmee_par": str(update.get("Validation_confirmee_par") or "").strip(),
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

        for key in (
            "Date_de_naissance",
            "Expire_le",
            "Email",
            "Montant",
            "ErreurValide",
            "Validation_confirmee_le",
            "Validation_confirmee_par",
        ):
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
    "ValidationParseResult",
    "build_ddn_lookup_from_rows",
    "import_already_printed_csv",
    "persist_last_import",
    "load_last_import",
    "parse_validation_workbook",
    "apply_validation_updates",
]
