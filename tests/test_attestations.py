from __future__ import annotations

from datetime import datetime
from pathlib import Path
import sqlite3
from zipfile import ZipFile

import src.app.attestations as attestations_module
from src.app.attestations import AttestationData, generate_attestation_pdf, load_attestation_settings
from src.app.db import fetch_latest_contact


def _write_template(path: Path) -> None:
    xml = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:body>
    <w:p><w:r><w:t>Attestation de paiement</w:t></w:r></w:p>
    <w:p><w:r><w:t>Mr, Mme, Melle .......</w:t></w:r></w:p>
    <w:p><w:r><w:t>pour la somme de .....€</w:t></w:r></w:p>
    <w:p><w:r><w:t>Fait à Nantes, le </w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    with ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", xml)


def test_generate_attestation_pdf(tmp_path: Path) -> None:
    template_path = tmp_path / "modele_attestation.docx"
    _write_template(template_path)

    previous_template = attestations_module.ATTESTATION_TEMPLATE_PATH
    attestations_module.ATTESTATION_TEMPLATE_PATH = template_path
    data = AttestationData(
        nom="DUPONT",
        prenom="ALICE",
        email="alice@example.com",
        montant="45",
        expire="31/12/2026",
        date_de_naissance="01/01/1990",
        generated_at=datetime(2024, 1, 15, 12, 0, 0),
    )

    try:
        pdf_path = generate_attestation_pdf(tmp_path, data)
    finally:
        attestations_module.ATTESTATION_TEMPLATE_PATH = previous_template

    assert pdf_path.exists()
    assert pdf_path.read_bytes().startswith(b"%PDF")
    pdf_bytes = pdf_path.read_bytes()
    assert b"Mr, Mme, Melle ALICE DUPONT" in pdf_bytes
    assert b"45 \x80" in pdf_bytes  # "45 €" en encodage CP1252
    assert b"Fait \xe0 Nantes, le 15/01/2024" in pdf_bytes


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
            montant TEXT,
            zpl_checksum TEXT,
            status TEXT,
            printed_at TEXT
        );
        """
    )

    cn.execute(
        "INSERT INTO prints (nom, prenom, ddn, expire, email, montant, status, printed_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            "DUPONT",
            "ALICE",
            "01/01/1990",
            "31/12/2025",
            "alice@example.com",
            "35",
            "printed",
            "2024-01-01T10:00:00",
        ),
    )
    cn.execute(
        "INSERT INTO prints (nom, prenom, ddn, expire, email, montant, status, printed_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            "DUPONT",
            "ALICE",
            "01/01/1990",
            "31/12/2026",
            "alice.new@example.com",
            "45",
            "printed",
            "2024-02-01T10:00:00",
        ),
    )

    contact = fetch_latest_contact(cn, "DUPONT", "ALICE", "01/01/1990")

    assert contact is not None
    assert contact.email == "alice.new@example.com"
    assert contact.montant == "45"
