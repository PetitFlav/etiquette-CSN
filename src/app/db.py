from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import sqlite3
from typing import Iterable, Optional
import hashlib
from datetime import datetime

DB_PATH = Path("data/app.db")

SCHEMA = """
PRAGMA journal_mode=WAL;

CREATE TABLE IF NOT EXISTS prints (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  nom TEXT NOT NULL,
  prenom TEXT NOT NULL,
  ddn TEXT NOT NULL,
  expire TEXT NOT NULL,
  email TEXT,
  montant TEXT,
  zpl_checksum TEXT,
  status TEXT NOT NULL DEFAULT 'printed', -- 'printed' ou 'simulated'
  printed_at TEXT NOT NULL                -- ISO 8601 via datetime.utcnow().isoformat()
);

-- ❌ on supprime l'unicité: on veut compter TOUTES les impressions
DROP INDEX IF EXISTS ux_prints_person_expire;

-- ✅ index utiles
CREATE INDEX IF NOT EXISTS ix_prints_person_expire ON prints(nom, prenom, ddn, expire);
CREATE INDEX IF NOT EXISTS ix_prints_person        ON prints(nom, prenom, ddn);
CREATE INDEX IF NOT EXISTS ix_prints_status_time   ON prints(status, printed_at);

CREATE TABLE IF NOT EXISTS attestation_emails (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  nom TEXT NOT NULL,
  prenom TEXT NOT NULL,
  ddn TEXT DEFAULT '',
  expire TEXT,
  email TEXT,
  montant TEXT,
  sent_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS ix_attestation_person ON attestation_emails(nom, prenom, ddn);
CREATE INDEX IF NOT EXISTS ix_attestation_sent   ON attestation_emails(sent_at);

-- ✅ vue agrégée par personne (Nom+Prénom+DDN), TOUTES expirations
--    cnt = nb d'impressions réelles (status='printed')
--    last_print = dernière date d'impression réelle
DROP VIEW IF EXISTS v_person_stats;
CREATE VIEW v_person_stats AS
SELECT
  nom,
  prenom,
  ddn,
  MAX(CASE WHEN status='printed' THEN printed_at END) AS last_print,
  SUM(CASE WHEN status='printed' THEN 1 ELSE 0 END)   AS cnt
FROM prints
GROUP BY nom, prenom, ddn;
"""

def connect(db_path: Path = DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    return conn

def init_db(db_path: Path = DB_PATH) -> None:
    with connect(db_path) as cn:
        cn.executescript(SCHEMA)
        _ensure_email_column(cn)
        _ensure_montant_column(cn)
        _ensure_attestation_table(cn)

def sha1(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def _ensure_columns(conn: sqlite3.Connection) -> set[str]:
    try:
        cur = conn.execute("PRAGMA table_info(prints)")
    except sqlite3.OperationalError:
        return set()

    return {
        (row["name"] if isinstance(row, sqlite3.Row) else row[1])
        for row in cur.fetchall()
    }

def _ensure_email_column(conn: sqlite3.Connection) -> None:
    """Add the ``email`` column to ``prints`` if it is missing."""

    columns = _ensure_columns(conn)
    if columns and "email" not in columns:
        conn.execute("ALTER TABLE prints ADD COLUMN email TEXT")


def _ensure_montant_column(conn: sqlite3.Connection) -> None:
    """Add the ``montant`` column to ``prints`` if it is missing."""

    columns = _ensure_columns(conn)
    if columns and "montant" not in columns:
        conn.execute("ALTER TABLE prints ADD COLUMN montant TEXT")


def _ensure_attestation_table(conn: sqlite3.Connection) -> None:
    """Make sure the attestation email log table exists."""

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS attestation_emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT NOT NULL,
            prenom TEXT NOT NULL,
            ddn TEXT DEFAULT '',
            expire TEXT,
            email TEXT,
            montant TEXT,
            sent_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS ix_attestation_person ON attestation_emails(nom, prenom, ddn)"
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS ix_attestation_sent ON attestation_emails(sent_at)"
    )

def already_printed(conn: sqlite3.Connection, nom: str, prenom: str, ddn: str, expire: str) -> bool:
    cur = conn.execute(
        "SELECT 1 FROM prints WHERE nom=? AND prenom=? AND ddn=? AND expire=? LIMIT 1",
        (nom.strip(), prenom.strip(), ddn.strip(), expire.strip()),
    )
    return cur.fetchone() is not None

def record_print(
    conn: sqlite3.Connection,
    nom: str,
    prenom: str,
    ddn: str,
    expire: str,
    email: str | None = None,
    montant: str | None = None,
    zpl: str | None = None,
    status: str = "printed",
) -> bool:
    _ensure_email_column(conn)
    _ensure_montant_column(conn)
    checksum = sha1(zpl) if zpl else None
    email_value = (email or "").strip()
    montant_value = (montant or "").strip()
    conn.execute(
        """
        INSERT INTO prints(nom, prenom, ddn, expire, email, montant, zpl_checksum, status, printed_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            nom.strip(),
            prenom.strip(),
            ddn.strip(),
            expire.strip(),
            email_value,
            montant_value,
            checksum,
            status,
            datetime.utcnow().isoformat(),
        ),
    )
    return True


def record_attestation_email(
    conn: sqlite3.Connection,
    nom: str,
    prenom: str,
    ddn: str | None,
    expire: str | None,
    email: str | None,
    montant: str | None,
    sent_at: datetime | None = None,
) -> bool:
    """Persist an attestation email send event."""

    _ensure_attestation_table(conn)
    payload_sent_at = (sent_at or datetime.utcnow()).isoformat()
    conn.execute(
        """
        INSERT INTO attestation_emails(nom, prenom, ddn, expire, email, montant, sent_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            (nom or "").strip(),
            (prenom or "").strip(),
            (ddn or "").strip(),
            (expire or "").strip(),
            (email or "").strip(),
            (montant or "").strip(),
            payload_sent_at,
        ),
    )
    return True


def update_person_montant(
    conn: sqlite3.Connection,
    nom: str,
    prenom: str,
    montant: str,
) -> int:
    """Update the ``montant`` value for the given person.

    Returns the number of rows that were updated.
    """

    if not nom or not prenom:
        return 0

    _ensure_montant_column(conn)
    montant_value = (montant or "").strip()
    cur = conn.execute(
        "UPDATE prints SET montant=? WHERE nom=? AND prenom=?",
        (montant_value, nom.strip(), prenom.strip()),
    )
    return cur.rowcount

def list_prints(conn: sqlite3.Connection, expire: Optional[str] = None) -> Iterable[sqlite3.Row]:
    if expire:
        cur = conn.execute("SELECT * FROM prints WHERE expire=? ORDER BY printed_at DESC", (expire,))
    else:
        cur = conn.execute("SELECT * FROM prints ORDER BY printed_at DESC")
    return cur.fetchall()

@dataclass
class PersonContact:
    nom: str
    prenom: str
    email: str
    montant: str = ""


def load_last_attestation_by_person(conn: sqlite3.Connection) -> dict[tuple[str, str, str], str]:
    """Return the latest attestation send timestamp grouped by person."""

    _ensure_attestation_table(conn)
    cur = conn.execute(
        """
        SELECT
            nom,
            prenom,
            COALESCE(ddn, '') AS ddn,
            MAX(sent_at) AS last_sent
        FROM attestation_emails
        GROUP BY nom, prenom, COALESCE(ddn, '')
        """
    )
    result: dict[tuple[str, str, str], str] = {}
    for row in cur.fetchall():
        key = (
            (row["nom"] or "").strip().lower(),
            (row["prenom"] or "").strip().lower(),
            (row["ddn"] or "").strip(),
        )
        result[key] = (row["last_sent"] or "").strip()
    return result


def person_stats(conn: sqlite3.Connection, nom: Optional[str] = None, prenom: Optional[str] = None) -> Iterable[sqlite3.Row]:
    sql = "SELECT nom, prenom, ddn, last_print, cnt FROM v_person_stats"
    params, conds = [], []
    if nom:
        conds.append("nom = ?"); params.append(nom.strip())
    if prenom:
        conds.append("prenom = ?"); params.append(prenom.strip())
    if conds:
        sql += " WHERE " + " AND ".join(conds)
    sql += " ORDER BY nom, prenom, ddn"
    return conn.execute(sql, params).fetchall()


def fetch_latest_contact(
    conn: sqlite3.Connection,
    nom: str,
    prenom: str,
    ddn: Optional[str] = None,
) -> PersonContact | None:
    try:
        params: list[str] = [nom.strip(), prenom.strip()]
        base_sql = "SELECT nom, prenom, email, montant FROM prints WHERE nom=? AND prenom=?"
        if ddn and ddn.strip():
            base_sql += " AND ddn=?"
            params.append(ddn.strip())
        params_tuple = tuple(params)

        with_email_sql = (
            base_sql + " AND email IS NOT NULL AND TRIM(email) != '' ORDER BY printed_at DESC LIMIT 1"
        )
        row_with_email = conn.execute(with_email_sql, params_tuple).fetchone()

        latest_sql = base_sql + " ORDER BY printed_at DESC LIMIT 1"
        latest_row = conn.execute(latest_sql, params_tuple).fetchone()
    except sqlite3.OperationalError:
        return None

    row_source = latest_row or row_with_email
    if not row_source:
        return None

    email_row = row_with_email or row_source
    montant_row = latest_row or row_with_email

    email = (email_row["email"] or "").strip()
    montant = (montant_row["montant"] or "").strip()

    return PersonContact(
        nom=(row_source["nom"] or "").strip() or nom.strip(),
        prenom=(row_source["prenom"] or "").strip() or prenom.strip(),
        email=email,
        montant=montant,
    )
