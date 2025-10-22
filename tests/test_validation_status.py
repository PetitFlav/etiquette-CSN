from __future__ import annotations

import os
import sqlite3
from datetime import datetime, timedelta

import pandas as pd

from src.app.validation import (
    DbExpiration,
    build_validation_lookup,
    compute_validation_status,
    find_latest_validation_export,
    load_latest_expiration_by_person,
    load_validation_export,
    parse_validator_names,
)


def test_find_latest_validation_export(tmp_path):
    directory = tmp_path / "Validation"
    directory.mkdir()
    first = directory / "alpha_validation.csv"
    second = directory / "beta_validation.csv"
    first.write_text("nom;prenom;valide_par\n", encoding="utf-8")
    second.write_text("nom;prenom;valide_par\n", encoding="utf-8")
    past = datetime.now() - timedelta(days=1)
    os.utime(first, (past.timestamp(), past.timestamp()))
    latest = find_latest_validation_export(directory)
    assert latest == second


def test_load_validation_export_returns_rows(tmp_path):
    path = tmp_path / "sample_validation.csv"
    df = pd.DataFrame(
        [
            {"nom": "DUPONT", "prenom": "ALICE", "valide_par": "Jean"},
            {"nom": "DURAND", "prenom": "Bob", "valide_par": "Marie"},
        ]
    )
    df.to_csv(path, sep=";", index=False, encoding="utf-8")
    rows = load_validation_export(path)
    assert rows == [
        {"nom": "DUPONT", "prenom": "ALICE", "valide_par": "Jean"},
        {"nom": "DURAND", "prenom": "Bob", "valide_par": "Marie"},
    ]


def test_compute_validation_status_applies_rules():
    db_lookup = {
        ("DUPONT", "ALICE"): DbExpiration(expire="31/12/2026", printed_at="2024-01-01"),
        ("DURAND", "BOB"): DbExpiration(expire="01/01/2025", printed_at="2024-01-02"),
        ("MARTIN", "CLARA"): DbExpiration(expire="31/12/2026", printed_at="2024-01-03"),
        ("RIVIERE", "EMMA"): DbExpiration(expire="31/12/2026", printed_at="2024-01-04"),
    }
    validation_lookup = build_validation_lookup(
        [
            {"nom": "dupont", "prenom": "alice", "valide_par": "Validé par Jean"},
            {"nom": "durand", "prenom": "bob", "valide_par": "Validé par Marie"},
            {"nom": "martin", "prenom": "clara", "valide_par": "Validé par Paul"},
        ]
    )
    validators = parse_validator_names("Jean;Marie")
    default_expire = "31/12/2026"

    status_none = compute_validation_status(("NOG", "BODY"), db_lookup, validation_lookup, default_expire, validators)
    status_question = compute_validation_status(("RIVIERE", "EMMA"), db_lookup, validation_lookup, default_expire, validators)
    status_red = compute_validation_status(("DURAND", "BOB"), db_lookup, validation_lookup, "31/12/2027", validators)
    status_green = compute_validation_status(("DUPONT", "ALICE"), db_lookup, validation_lookup, default_expire, validators)
    status_orange = compute_validation_status(("MARTIN", "CLARA"), db_lookup, validation_lookup, default_expire, validators)

    assert status_none == ""
    assert status_question == "question"
    assert status_red == "red"
    assert status_green == "green"
    assert status_orange == "orange"


def test_load_latest_expiration_by_person_returns_latest(tmp_path):
    db_path = tmp_path / "test.db"
    conn = sqlite3.connect(db_path)
    conn.execute(
        """
        CREATE TABLE prints (
            nom TEXT,
            prenom TEXT,
            ddn TEXT,
            expire TEXT,
            zpl_checksum TEXT,
            status TEXT,
            printed_at TEXT
        )
        """
    )
    rows = [
        ("Dupont", "Alice", "", "31/12/2025", None, "printed", "2024-01-01T10:00:00"),
        ("Dupont", "Alice", "", "31/12/2026", None, "printed", "2024-02-01T10:00:00"),
        ("Durand", "Bob", "", "31/12/2024", None, "printed", "2023-12-01T10:00:00"),
    ]
    conn.executemany(
        "INSERT INTO prints(nom, prenom, ddn, expire, zpl_checksum, status, printed_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
        rows,
    )
    conn.commit()
    lookup = load_latest_expiration_by_person(conn)
    conn.close()

    assert lookup[("DUPONT", "ALICE")].expire == "31/12/2026"
    assert lookup[("DURAND", "BOB")].expire == "31/12/2024"
