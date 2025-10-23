from __future__ import annotations
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

def sha1(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def _ensure_email_column(conn: sqlite3.Connection) -> None:
    """Add the ``email`` column to ``prints`` if it is missing."""

    try:
        cur = conn.execute("PRAGMA table_info(prints)")
    except sqlite3.OperationalError:
        return

    columns = {
        (row["name"] if isinstance(row, sqlite3.Row) else row[1])
        for row in cur.fetchall()
    }

    if "email" not in columns:
        conn.execute("ALTER TABLE prints ADD COLUMN email TEXT")

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
    zpl: str | None = None,
    status: str = "printed",
) -> bool:
    _ensure_email_column(conn)
    checksum = sha1(zpl) if zpl else None
    email_value = (email or "").strip()
    conn.execute(
        """
        INSERT INTO prints(nom, prenom, ddn, expire, email, zpl_checksum, status, printed_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            nom.strip(),
            prenom.strip(),
            ddn.strip(),
            expire.strip(),
            email_value,
            checksum,
            status,
            datetime.utcnow().isoformat(),
        ),
    )
    return True

def list_prints(conn: sqlite3.Connection, expire: Optional[str] = None) -> Iterable[sqlite3.Row]:
    if expire:
        cur = conn.execute("SELECT * FROM prints WHERE expire=? ORDER BY printed_at DESC", (expire,))
    else:
        cur = conn.execute("SELECT * FROM prints ORDER BY printed_at DESC")
    return cur.fetchall()

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
