from __future__ import annotations

from dataclasses import replace
from datetime import datetime
from pathlib import Path
import sqlite3
from zipfile import ZipFile

import pytest

import src.app.attestations as attestations_module
from src.app.attestations import (
    AttestationData,
    SMTPSettings,
    _smtp_connection,
    generate_attestation_pdf,
    load_attestation_settings,
)
from src.app.crypto_utils import encrypt_secret
from src.app.db import fetch_latest_contact


def _write_template(path: Path) -> None:
    xml = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:body>
    <w:p><w:r><w:t>Attestation de paiement</w:t></w:r></w:p>
    <w:p><w:r><w:t>Nom : &lt;nom&gt;</w:t></w:r></w:p>
    <w:p><w:r><w:t>Prénom : &lt;prenom&gt;</w:t></w:r></w:p>
    <w:p><w:r><w:t>Montant : &lt;montant&gt;</w:t></w:r></w:p>
    <w:p><w:r><w:t>Saison : &lt;saison&gt;</w:t></w:r></w:p>
    <w:p><w:r><w:t>Date : &lt;DateDuJour&gt;</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
    with ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", xml)


def test_generate_attestation_pdf(tmp_path: Path) -> None:
    template_path = tmp_path / "modele_attestation.docx"
    _write_template(template_path)

    data = AttestationData(
        nom="DUPONT",
        prenom="ALICE",
        email="alice@example.com",
        montant="45",
        expire="31/12/2026",
        date_de_naissance="01/01/1990",
        generated_at=datetime(2024, 1, 15, 12, 0, 0),
    )

    def dummy_converter(docx_path: Path, pdf_path: Path) -> None:
        with ZipFile(docx_path) as archive:
            xml = archive.read("word/document.xml").decode("utf-8")
        assert "Nom : DUPONT" in xml
        assert "Prénom : DUPONT" not in xml  # ensure order preserved
        assert "Prénom : ALICE" in xml
        assert "Montant : 45 €" in xml
        assert "Saison : 2025/2026" in xml
        assert "Date : 15/01/2024" in xml
        pdf_payload = (
            "%PDF-1.4\nAttestation ALICE DUPONT\nMontant 45 €\n"
        ).encode("cp1252", "ignore")
        pdf_path.write_bytes(pdf_payload)

    pdf_path = generate_attestation_pdf(
        tmp_path,
        data,
        template_path=template_path,
        converter=dummy_converter,
    )

    expected_directory = tmp_path / "envoyees"
    assert pdf_path.parent == expected_directory
    assert pdf_path.name == "DUPONT_ALICE_attestation_2025_2026.pdf"
    assert pdf_path.exists()
    assert pdf_path.read_bytes().startswith(b"%PDF")
    pdf_bytes = pdf_path.read_bytes()
    assert b"Attestation ALICE DUPONT" in pdf_bytes
    assert b"45 \x80" in pdf_bytes  # "45 €" en encodage CP1252


def test_generate_attestation_pdf_overwrites_existing_file(tmp_path: Path) -> None:
    template_path = tmp_path / "modele_attestation.docx"
    _write_template(template_path)

    data = AttestationData(
        nom="Dupont",
        prenom="Alice",
        email="alice@example.com",
        montant="45",
        expire="31/12/2026",
        date_de_naissance="01/01/1990",
        generated_at=datetime(2024, 1, 15, 12, 0, 0),
    )

    def converter(docx_path: Path, pdf_path: Path) -> None:
        with ZipFile(docx_path) as archive:
            xml = archive.read("word/document.xml")
        pdf_path.write_bytes(b"%PDF-1.4\n" + xml)

    first_path = generate_attestation_pdf(
        tmp_path,
        data,
        template_path=template_path,
        converter=converter,
    )
    first_bytes = first_path.read_bytes()

    updated_data = replace(
        data,
        montant="50",
        generated_at=datetime(2024, 2, 1, 9, 30, 0),
    )

    second_path = generate_attestation_pdf(
        tmp_path,
        updated_data,
        template_path=template_path,
        converter=converter,
    )

    assert second_path == first_path
    assert second_path.read_bytes() != first_bytes


def test_load_attestation_settings_defaults() -> None:
    cfg = {"smtp_host": "smtp.example.com", "smtp_sender": "noreply@example.com"}
    settings = load_attestation_settings(cfg)

    assert settings.host == "smtp.example.com"
    assert settings.sender == "noreply@example.com"
    assert settings.port == 587
    assert settings.use_tls is True


def test_load_attestation_settings_decrypts_encrypted_password() -> None:
    encrypted = encrypt_secret("monsecret")
    cfg = {
        "smtp_host": "smtp.example.com",
        "smtp_sender": "noreply@example.com",
        "smtp_user": "user",
        "smtp_password": encrypted,
    }

    settings = load_attestation_settings(cfg)

    assert settings.password == "monsecret"


def test_smtp_connection_decrypts_password_before_login(monkeypatch) -> None:
    encrypted = encrypt_secret("monsecret")
    settings = SMTPSettings(
        host="smtp.example.com",
        port=587,
        sender="noreply@example.com",
        username="user",
        password=encrypted,
        use_tls=False,
        use_ssl=False,
    )

    login_calls: list[tuple[str, str]] = []

    class DummySMTP:
        def __init__(self, host, port, timeout):
            assert host == "smtp.example.com"
            assert port == 587
            assert timeout == settings.timeout

        def login(self, username, password):
            login_calls.append((username, password))

        def quit(self):
            pass

    monkeypatch.setattr(attestations_module.smtplib, "SMTP", DummySMTP)

    with _smtp_connection(settings):
        pass

    assert login_calls == [("user", "monsecret")]


def test_smtp_connection_raises_for_invalid_encrypted_password(monkeypatch) -> None:
    settings = SMTPSettings(
        host="smtp.example.com",
        port=587,
        sender="noreply@example.com",
        username="user",
        password="enc:###",  # invalide
        use_tls=False,
        use_ssl=False,
    )

    class DummySMTP:
        def __init__(self, host, port, timeout):
            pass

        def login(self, username, password):  # pragma: no cover - should not be called
            raise AssertionError("login should not be called when password is invalid")

        def quit(self):
            pass

    monkeypatch.setattr(attestations_module.smtplib, "SMTP", DummySMTP)

    with pytest.raises(RuntimeError, match="Mot de passe SMTP chiffré invalide"):
        with _smtp_connection(settings):
            pass


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


def test_fetch_latest_contact_uses_latest_montant_even_without_email() -> None:
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
            "",
            "50",
            "printed",
            "2024-03-01T10:00:00",
        ),
    )

    contact = fetch_latest_contact(cn, "DUPONT", "ALICE", "01/01/1990")

    assert contact is not None
    assert contact.email == "alice@example.com"
    assert contact.montant == "50"
