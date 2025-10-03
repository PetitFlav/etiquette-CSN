from __future__ import annotations

import csv
import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

from .config import DB_PATH, LAST_IMPORT_DIR, LAST_IMPORT_METADATA
from .db import connect, record_print
from .io_utils import lire_tableau, normalize_name


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
]
