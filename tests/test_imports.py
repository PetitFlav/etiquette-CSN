from __future__ import annotations

from pathlib import Path

from src.app.db import connect, init_db, record_print
from src.app.imports import (
    build_ddn_lookup_from_rows,
    import_already_printed_csv,
    load_last_import,
    persist_last_import,
)
from src.app.io_utils import lire_tableau


def test_build_ddn_lookup_handles_conflicts():
    rows = [
        {"Nom": "Dup", "Prénom": "Test", "Date_de_naissance": "01/01/2000"},
        {"Nom": "Dup", "Prénom": "Test", "Date_de_naissance": "02/02/2000"},
        {"Nom": "Unique", "Prénom": "One", "Date_de_naissance": "03/03/2000"},
        {"Nom": "dùpônt", "Prénom": "álix", "Date_de_naissance": "04/04/2000"},
        {"Nom": "", "Prénom": "", "Date_de_naissance": ""},
    ]

    lookup = build_ddn_lookup_from_rows(rows)
    assert lookup[("unique", "one")] == "03/03/2000"
    assert lookup[("dup", "test")] is None
    assert lookup[("dupont", "alix")] == "04/04/2000"


def test_import_already_printed_csv_resolves_ddn(tmp_path):
    db_path = tmp_path / "app.db"
    init_db(db_path)

    with connect(db_path) as cn:
        # Unique DDN for Beta
        record_print(cn, "Beta", "User", "10/10/2010", "01/01/2024", zpl=None, status="printed")
        # Multiple DDN for Gamma to trigger the "ambiguous" path
        record_print(cn, "Gamma", "User", "11/11/2011", "01/01/2024", zpl=None, status="printed")
        record_print(cn, "Gamma", "User", "22/02/2012", "02/02/2024", zpl=None, status="printed")

    csv_path = tmp_path / "import.csv"
    csv_path.write_text("Alpha;Test\nBeta;User\nGamma;User\n;\n", encoding="utf-8")

    rows_lookup = build_ddn_lookup_from_rows(
        [{"Nom": "Alpha", "Prénom": "Test", "Date_de_naissance": "05/05/2005"}]
    )

    imported, skipped = import_already_printed_csv(
        csv_path,
        "31/12/2025",
        rows_ddn_lookup=rows_lookup,
        db_path=db_path,
    )

    assert imported == 3
    assert skipped == 1

    with connect(db_path) as cn:
        rows = cn.execute(
            "SELECT nom, prenom, ddn FROM prints WHERE expire=? ORDER BY nom",
            ("31/12/2025",),
        ).fetchall()

    assert [tuple(row) for row in rows] == [
        ("ALPHA", "TEST", "05/05/2005"),
        ("BETA", "USER", "10/10/2010"),
        ("GAMMA", "USER", ""),
    ]


def test_import_normalizes_names(tmp_path):
    db_path = tmp_path / "app.db"
    init_db(db_path)

    csv_path = tmp_path / "import.csv"
    csv_path.write_text("dùpônt;álix\n", encoding="utf-8")

    rows_lookup = build_ddn_lookup_from_rows(
        [{"Nom": "dùpônt", "Prénom": "álix", "Date_de_naissance": "01/01/2000"}]
    )

    imported, skipped = import_already_printed_csv(
        csv_path,
        "31/12/2025",
        rows_ddn_lookup=rows_lookup,
        db_path=db_path,
    )

    assert imported == 1
    assert skipped == 0

    with connect(db_path) as cn:
        row = cn.execute(
            "SELECT nom, prenom, ddn FROM prints WHERE expire=?",
            ("31/12/2025",),
        ).fetchone()

    assert tuple(row) == ("DUPONT", "ALIX", "01/01/2000")


def _write_sample_csv(path: Path) -> None:
    path.write_text(
        "Ignoré\nIgnoré\nIgnoré\n"
        "Nom,Prénom,Date_de_naissance,Expire_le,Email,Montant,ErreurValide\n"
        "Alpha,Test,01/01/2000,31/12/2025,alpha@example.com,123.45,Yes\n",
        encoding="utf-8",
    )


def _write_csv_without_optional_columns(path: Path) -> None:
    path.write_text(
        "Ignoré\nIgnoré\nIgnoré\n"
        "Nom,Prénom,Date_de_naissance,Expire_le,Email\n"
        "Beta,User,02/02/2002,31/12/2025,beta@example.com\n",
        encoding="utf-8",
    )


def test_lire_tableau_adds_missing_optional_columns(tmp_path):
    sample = tmp_path / "import.csv"
    _write_csv_without_optional_columns(sample)

    df = lire_tableau(sample)

    assert list(df.columns) == [
        "Nom",
        "Prénom",
        "Date_de_naissance",
        "Expire_le",
        "Email",
        "Montant",
        "ErreurValide",
    ]
    assert df.iloc[0]["Nom"] == "BETA"
    assert df.iloc[0]["Montant"] == ""
    assert df.iloc[0]["ErreurValide"] == ""


def test_persist_and_load_last_import(tmp_path, monkeypatch):
    from src.app import imports as imports_mod

    source = tmp_path / "import.csv"
    _write_sample_csv(source)

    cache_dir = tmp_path / "cache"
    metadata_file = tmp_path / "meta.json"

    monkeypatch.setattr(imports_mod, "LAST_IMPORT_DIR", cache_dir)
    monkeypatch.setattr(imports_mod, "LAST_IMPORT_METADATA", metadata_file)

    metadata = persist_last_import(source)

    assert metadata_file.exists()
    assert Path(metadata["cached_path"]).exists()
    assert metadata["source_name"] == source.name
    assert metadata["cached_name"] == source.name
    assert "stored_at" in metadata

    rows, loaded_metadata = load_last_import()

    assert len(rows) == 1
    assert rows[0]["Nom"] == "ALPHA"
    assert rows[0]["Email"] == "alpha@example.com"
    assert rows[0]["Montant"] == "123.45"
    assert rows[0]["ErreurValide"] == "Yes"
    assert loaded_metadata["cached_name"] == source.name
