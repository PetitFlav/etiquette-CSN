from __future__ import annotations

from pathlib import Path
import sqlite3

from src.app.attestations import AttestationData, generate_attestation_pdf, load_attestation_settings
from src.app.db import fetch_latest_contact


def test_generate_attestation_pdf(tmp_path: Path) -> None:
    data = AttestationData(
        nom="DUPONT",
        prenom="ALICE",
        email="alice@example.com",
        montant="45 €",
        expire="31/12/2026",
        date_de_naissance="01/01/1990",
    )

    pdf_path = generate_attestation_pdf(tmp_path, data)

    assert pdf_path.exists()
    assert pdf_path.read_bytes().startswith(b"%PDF")


def test_load_attestation_settings_defaults() -> None:
    cfg = {"smtp_host": "smtp.example.com", "smtp_sender": "noreply@example.com"}
    settings = load_attestation_settings(cfg)

    assert settings.host == "smtp.example.com"
    assert settings.sender == "noreply@example.com"
    assert settings.port == 587
    assert settings.use_tls is True


def test_fetch_latest_contact_prefers_most_recent_email() -> None:
    cn = sqlite3.connect(":memory:")
    cn.row_factory = sqlite3.Row
    cn.executescript(
        """
        CREATE TABLE prints (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT,
            prenom TEXT,
            ddn TEXT,
            expire TEXT,
            email TEXT,
            zpl_checksum TEXT,
            status TEXT,
            printed_at TEXT
        );
        """
    )

    cn.execute(
        "INSERT INTO prints (nom, prenom, ddn, expire, email, status, printed_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
        ("DUPONT", "ALICE", "01/01/1990", "31/12/2025", "alice@example.com", "printed", "2024-01-01T10:00:00"),
    )
    cn.execute(
        "INSERT INTO prints (nom, prenom, ddn, expire, email, status, printed_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (
            "DUPONT",
            "ALICE",
            "01/01/1990",
            "31/12/2026",
            "alice.new@example.com",
            "printed",
            "2024-02-01T10:00:00",
        ),
    )

    contact = fetch_latest_contact(cn, "DUPONT", "ALICE", "01/01/1990")

    assert contact is not None
    assert contact.email == "alice.new@example.com"
