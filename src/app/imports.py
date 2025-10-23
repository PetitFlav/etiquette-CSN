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

from .config import (
    DB_PATH,
    LAST_IMPORT_DIR,
    LAST_IMPORT_METADATA,
    VALIDATION_EXPORT_DIR,
)
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


NAME_AT_START = re.compile(
    r"^([A-ZÉÈÀÂÎÔÙÛÇ'`-]{2,})\s+([A-ZÉÈÀÂÎÔÙÛÇ][A-Za-zÀ-ÖØ-öø-ÿ'`-]*(?:\s+[A-ZÉÈÀÂÎÔÙÛÇ][A-Za-zÀ-ÖØ-öø-ÿ'`-]*)*)"
)
VALIDATOR_RX = re.compile(
    r"\bpar\s+([A-ZÉÈÀÂÎÔÙÛÇ'`-]{2,})\s+([A-ZÉÈÀÂÎÔÙÛÇ][A-Za-zÀ-ÖØ-öø-ÿ'`-]*(?:\s+[A-ZÉÈÀÂÎÔÙÛÇ][A-Za-zÀ-ÖØ-öø-ÿ'`-]*)*)",
    re.IGNORECASE,
)
MONTANT_RX = re.compile(r"montant\s*[:\-]?\s*([0-9]+(?:[.,][0-9]{1,2})?)", re.IGNORECASE)


def _norm_space(text: object) -> str:
    value = str(text or "")
    if not value:
        return ""
    value = (
        value.replace("\ufeff", "")
        .replace("\u00A0", " ")
        .replace("\u202F", " ")
        .replace("|", " ")
        .replace(";", " ")
    )
    value = re.sub(r"\s*-\s*", "-", value)
    return re.sub(r"\s+", " ", value).strip()


def _row_to_text(row: Sequence[object]) -> str:
    parts: list[str] = []
    for value in row:
        text = _norm_space(value)
        if text:
            parts.append(text)
    return " ".join(parts)


def _format_first_name(value: str) -> str:
    base = _norm_space(value)
    if not base:
        return ""
    tokens = base.split(" ")
    formatted_tokens: list[str] = []
    for token in tokens:
        if not token:
            continue
        hyphen_parts = token.split("-")
        formatted_parts = [
            part[:1].upper() + part[1:].lower() if part else ""
            for part in hyphen_parts
        ]
        formatted_tokens.append("-".join(formatted_parts))
    return " ".join(formatted_tokens)


COMMON_FIRST_NAME_TOKENS = {
    "ADAM",
    "ALAIN",
    "ALBERT",
    "ALEXANDRE",
    "ALICE",
    "ANNA",
    "ANNE",
    "ANTOINE",
    "ARTHUR",
    "BENJAMIN",
    "CAMILLE",
    "CAROLINE",
    "CHARLES",
    "CHRISTIAN",
    "CLAIRE",
    "CLAUDE",
    "DANIEL",
    "DAVID",
    "ELISE",
    "EMILIE",
    "EMMA",
    "FLORENT",
    "FRANCIS",
    "FRANCOIS",
    "FRANÇOIS",
    "FRANÇOISE",
    "GABRIEL",
    "HENRI",
    "ISABELLE",
    "JACQUES",
    "JEAN",
    "JEANNE",
    "JEROME",
    "JULIE",
    "JULIEN",
    "LAURENT",
    "LUC",
    "LUCIE",
    "LOUIS",
    "LOUISE",
    "MARIA",
    "MARIE",
    "MARION",
    "MARTIN",
    "MATHIEU",
    "MICHEL",
    "MICHELLE",
    "NICOLAS",
    "NOEL",
    "NOËL",
    "PAUL",
    "PIERRE",
    "REMY",
    "RENÉ",
    "ROBERT",
    "SARAH",
    "SIMON",
    "SOPHIE",
    "STEPHANE",
    "THIERRY",
    "THOMAS",
    "VALENTIN",
    "VALERIE",
    "VICTOR",
}


def _token_starts_with_two_uppercase(token: str) -> bool:
    if "-" in token:
        # Hyphenated first names (e.g. "JEAN-PAUL") should not be merged into the
        # last name even if they appear in uppercase in the source document.
        return False

    letters = [char for char in token if char.isalpha()]
    if len(letters) < 2:
        return False
    if token.upper() in COMMON_FIRST_NAME_TOKENS:
        return False
    return letters[0].isupper() and letters[1].isupper()


def _split_compound_last_name(last_name: str, first_name: str) -> tuple[str, str]:
    last_tokens = [token for token in last_name.split(" ") if token]
    first_tokens = [token for token in first_name.split(" ") if token]

    if not first_tokens:
        return " ".join(last_tokens), ""

    max_move = 0
    for token in first_tokens:
        if _token_starts_with_two_uppercase(token):
            max_move += 1
        else:
            break

    def score(candidate_last: list[str], candidate_first: list[str], moved: int) -> int:
        if not candidate_first:
            return -1_000

        score_value = 0

        if any(any(ch.islower() for ch in token) for token in candidate_first):
            score_value += 5

        score_value += len(candidate_first)

        if candidate_first and candidate_first[0].upper() in COMMON_FIRST_NAME_TOKENS:
            score_value += 3

        if any(token.upper() in COMMON_FIRST_NAME_TOKENS for token in candidate_last):
            score_value -= 4

        if moved == 0:
            score_value += 1

        return score_value

    best_last = list(last_tokens)
    best_first = list(first_tokens)
    best_score = score(best_last, best_first, 0)
    best_move = 0

    for move_count in range(1, max_move + 1):
        candidate_last = last_tokens + first_tokens[:move_count]
        candidate_first = first_tokens[move_count:]
        candidate_score = score(candidate_last, candidate_first, move_count)

        if candidate_score > best_score or (
            candidate_score == best_score
            and len(candidate_first) < len(best_first)
        ) or (
            candidate_score == best_score and len(candidate_first) == len(best_first)
            and move_count > best_move
        ):
            best_score = candidate_score
            best_last = candidate_last
            best_first = candidate_first
            best_move = move_count

    return " ".join(best_last), " ".join(best_first)


def _clean_output(text: str) -> str:
    normalized = strip_accents(str(text or "")).replace("€", "")
    return _norm_space(normalized)


def _normalize_compound_name(text: str) -> str:
    """Return ``text`` in uppercase while preserving compound separators."""

    cleaned = _clean_output(text)
    if not cleaned:
        return ""

    # Collapse any spacing around hyphens so that tokens are contiguous.
    normalized_hyphen = re.sub(r"\s*-\s*", "-", cleaned)

    formatted_tokens: list[str] = []
    for token in normalized_hyphen.split(" "):
        if not token:
            continue
        if "-" in token:
            parts = [part for part in token.split("-") if part]
            if not parts:
                continue
            formatted_tokens.append("-".join(part.upper() for part in parts))
        else:
            formatted_tokens.append(token.upper())

    return " ".join(formatted_tokens)


def _extract_validator(*lines: str) -> str:
    for candidate in lines:
        if not candidate:
            continue
        match = VALIDATOR_RX.search(candidate)
        if not match:
            continue
        base_nom, base_prenom = _split_compound_last_name(
            _norm_space(match.group(1)), _norm_space(match.group(2))
        )
        nom = _normalize_compound_name(base_nom)
        prenom = _format_first_name(base_prenom)
        formatted = f"{prenom} {nom}".strip()
        return _clean_output(formatted)
    return ""


def _extract_amount(*lines: str) -> str:
    for candidate in lines:
        if not candidate:
            continue
        if "montant" not in candidate.lower():
            continue
        match = MONTANT_RX.search(candidate)
        if not match:
            continue
        raw = match.group(1)
        cleaned = (
            raw.replace(" ", "")
            .replace("\u00A0", "")
            .replace("\u202F", "")
            .replace(",", ".")
        )
        try:
            amount = Decimal(cleaned)
        except InvalidOperation:
            continue
        return f"{amount.quantize(Decimal('0.01'))}"
    return ""


def _load_validation_lines(source: Path) -> list[str]:
    suffix = source.suffix.lower()
    lines: list[str]
    if suffix in {".xlsx", ".xls"}:
        try:
            df = pd.read_excel(source, header=None, dtype=str).fillna("")
        except ValueError:
            df = pd.read_excel(source, header=None, dtype=str, engine="xlrd").fillna("")
        lines = [_row_to_text(row) for row in df.itertuples(index=False, name=None)]
    else:
        try:
            text = source.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            text = source.read_text(encoding="latin-1")
        normalized = text.replace("\r\n", "\n").replace("\r", "\n")
        lines = [_norm_space(line) for line in normalized.split("\n")]
    return [line for line in lines if line]


def parse_validation_three_line_file(
    path: Path | str, *, output_dir: Path | None = None
) -> ValidationParseResult:
    """Extract members from a three-line validation file and export a CSV.

    The input file can be a CSV or Excel workbook where each member spans three
    consecutive lines (``Nom Prenom`` / ``Payé ...`` / ``Montant : xxx €``).
    Accents are stripped in the exported CSV and the output is saved to the
    ``Validation`` sub-directory of the application's ``data`` folder unless a
    custom ``output_dir`` is provided.
    """

    source = Path(path)
    if not source.exists():
        raise FileNotFoundError(f"Fichier introuvable: {source}")

    raw_lines = _load_validation_lines(source)
    records: list[dict[str, str]] = []

    i = 0
    while i < len(raw_lines):
        line1 = raw_lines[i]
        match = NAME_AT_START.search(line1)
        if not match:
            i += 1
            continue

        line2 = raw_lines[i + 1] if i + 1 < len(raw_lines) else ""
        line3 = raw_lines[i + 2] if i + 2 < len(raw_lines) else ""

        if not line2 or not line3:
            i += 1
            continue

        pay_lines = [candidate for candidate in (line2, line3) if candidate]
        if not any("pay" in candidate.lower() for candidate in pay_lines):
            i += 1
            continue

        montant = _extract_amount(line2, line3)
        if not montant:
            i += 1
            continue

        base_nom, base_prenom = _split_compound_last_name(
            _norm_space(match.group(1)), _norm_space(match.group(2))
        )
        nom = normalize_name(base_nom)
        prenom = normalize_name(base_prenom)
        valide_par = _extract_validator(line1, line2, line3)

        records.append(
            {
                "nom": nom,
                "prenom": prenom,
                "valide_par": valide_par,
                "montant": montant,
            }
        )
        i += 3

        # Ignore any trailing lines linked to the current member until the
        # next "Nom Prenom" line is found.  Some exports insert additional
        # informational rows that should be skipped entirely.
        while i < len(raw_lines) and not NAME_AT_START.search(raw_lines[i]):
            i += 1
        continue

    columns = ["nom", "prenom", "valide_par", "montant"]
    df = pd.DataFrame(records, columns=columns)
    if not df.empty:
        df = df.drop_duplicates()
        records = df.to_dict(orient="records")
    else:
        records = []

    target_dir = output_dir or VALIDATION_EXPORT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target_path = target_dir / f"{source.stem}_{timestamp}_validation.csv"
    df.to_csv(target_path, index=False, sep=";", encoding="utf-8", columns=columns)

    return ValidationParseResult(rows=records, export_path=target_path)


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


def _flatten_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return ``df`` with a single level header.

    LibreOffice may export validation workbooks with drawings that generate an
    additional header level in ``pandas.read_excel``.  Downstream code expects a
    flat index, so we coerce the header back to strings here.
    """

    if not isinstance(df.columns, pd.MultiIndex):
        return df

    flattened = []
    for column in df.columns:  # type: ignore[union-attr]
        parts = [str(part).strip() for part in column if str(part).strip() and str(part).strip().lower() != "nan"]
        flattened.append(" ".join(parts) if parts else "")

    df = df.copy()
    df.columns = flattened
    return df


def _drop_image_like_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Remove columns that only contain embedded image markers.

    LibreOffice stores pictures linked to rows as additional columns that only
    contain OLE/embedded markers.  They should be ignored when normalizing the
    dataset.
    """

    image_like_columns: list[str] = []
    for column in df.columns:
        normalized = _normalize_header_label(column)
        if normalized.startswith("image") or normalized.startswith("photo"):
            image_like_columns.append(column)
            continue

        values = df[column].dropna()
        if values.empty:
            continue

        def _looks_like_image_marker(value: object) -> bool:
            text = str(value or "").strip().lower()
            if not text:
                return True
            return (
                text.startswith("=embed(")
                or text.startswith("oleobject")
                or text.startswith("picture")
                or text.startswith("bitmap")
                or text.endswith(".png")
                or text.endswith(".jpg")
                or text.endswith(".jpeg")
            )

        if values.map(_looks_like_image_marker).all():
            image_like_columns.append(column)

    if image_like_columns:
        df = df.drop(columns=image_like_columns)
    return df


def _expand_single_row_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Expand a single-row dataframe where values are separated by newlines.

    Some LibreOffice exports flatten the table into a single row whose cells
    contain the column title followed by the data separated with newlines.  We
    detect this situation and rebuild the tabular structure so that each
    adherent corresponds to its own row again.
    """

    if len(df) != 1:
        return df

    row = df.iloc[0]
    expanded: dict[str, list[object]] = {}
    max_length = 0
    multiline_detected = False

    for column in df.columns:
        value = row[column]
        if isinstance(value, str):
            normalized_text = value.replace("\r\n", "\n").replace("\r", "\n")
            parts = [part.strip() for part in normalized_text.split("\n")]
            if len(parts) > 1:
                multiline_detected = True
            normalized_column = _normalize_header_label(column)
            expected_headers = {normalized_column}
            target_column = _VALIDATION_HEADER_MAP.get(normalized_column)
            if target_column:
                expected_headers.add(
                    _normalize_header_label(str(target_column).replace("_", " "))
                )
            if parts and _normalize_header_label(parts[0]) in expected_headers:
                parts = parts[1:]
            expanded[column] = parts if parts else [""]
        else:
            expanded[column] = [value]
        max_length = max(max_length, len(expanded[column]))

    if not multiline_detected or max_length <= 1:
        return df

    for column, values in expanded.items():
        if len(values) < max_length:
            values = values + ["" for _ in range(max_length - len(values))]
        expanded[column] = values

    return pd.DataFrame(expanded)


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

    df = _flatten_dataframe_columns(df)
    df = _drop_image_like_columns(df)
    df = _expand_single_row_dataframe(df)

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

    def _normalize_erreur_valide_flag(value: object) -> str:
        if isinstance(value, bool):
            return "true" if value else "false"
        text = str(value or "").strip().lower()
        if text in {"true", "yes", "1", "y", "oui"}:
            return "true"
        if text in {"false", "no", "0", "n", "non"}:
            return "false"
        return ""

    for update in updates or []:
        nom = _norm(update.get("Nom"))
        prenom = _norm(update.get("Prénom"))
        if not nom and not prenom:
            continue

        ddn_update = str(update.get("Date_de_naissance") or "").strip()
        expire_update = str(update.get("Expire_le") or "").strip()
        montant_update = str(update.get("Montant") or "").strip()

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

        expire_existing = str(row.get("Expire_le") or "").strip()

        if row.get("Nom") != nom:
            row["Nom"] = nom
            changed = True
        if row.get("Prénom") != prenom:
            row["Prénom"] = prenom
            changed = True

        for key in ("Date_de_naissance", "Expire_le", "Email"):
            value = update.get(key)
            if value is None:
                continue
            value_str = str(value).strip()
            if not value_str:
                continue
            if str(row.get(key) or "").strip() != value_str:
                row[key] = value_str
                changed = True

        if montant_update:
            if str(row.get("Montant") or "").strip() != montant_update:
                row["Montant"] = montant_update
                changed = True

        if expire_update:
            if expire_existing and expire_existing != expire_update:
                desired_flag = "false"
            else:
                desired_flag = "true"
            current_flag = _normalize_erreur_valide_flag(row.get("ErreurValide"))
            if current_flag != desired_flag:
                row["ErreurValide"] = desired_flag
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
    "parse_validation_three_line_file",
    "parse_validation_workbook",
    "apply_validation_updates",
]
