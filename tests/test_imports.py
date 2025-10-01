from __future__ import annotations

from pathlib import Path

from src.app.db import connect, init_db, record_print
from src.app.imports import (
    build_ddn_lookup_from_rows,
    import_already_printed_csv,
    load_last_import,
    persist_last_import,
)


def test_build_ddn_lookup_handles_conflicts():
    rows = [
        {"Nom": "Dup", "Prénom": "Test", "Date_de_naissance": "01/01/2000"},
        {"Nom": "Dup", "Prénom": "Test", "Date_de_naissance": "02/02/2000"},
        {"Nom": "Unique", "Prénom": "One", "Date_de_naissance": "03/03/2000"},
        {"Nom": "", "Prénom": "", "Date_de_naissance": ""},
    ]

    lookup = build_ddn_lookup_from_rows(rows)
    assert lookup[("unique", "one")] == "03/03/2000"
    assert lookup[("dup", "test")] is None


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
        ("Alpha", "Test", "05/05/2005"),
        ("Beta", "User", "10/10/2010"),
        ("Gamma", "User", ""),
    ]


def _write_sample_csv(path: Path) -> None:
    path.write_text(
        "Ignoré\nIgnoré\nIgnoré\n"
        "Nom,Prénom,Date_de_naissance,Expire_le\n"
        "Alpha,Test,01/01/2000,31/12/2025\n",
        encoding="utf-8",
    )


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
    assert rows[0]["Nom"] == "Alpha"
    assert loaded_metadata["cached_name"] == source.name
