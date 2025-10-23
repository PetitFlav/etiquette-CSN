from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from src.app.db import connect, init_db, record_print
from src.app.imports import (
    apply_validation_updates,
    build_ddn_lookup_from_rows,
    import_already_printed_csv,
    load_last_import,
    parse_validation_three_line_file,
    parse_validation_workbook,
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


def test_import_handles_header_with_amount_column(tmp_path):
    db_path = tmp_path / "app.db"
    init_db(db_path)

    csv_path = tmp_path / "import.csv"
    csv_path.write_text(
        "Nom;Prénom;Montant\nAlpha;Test;15\nBeta;User;30\n",
        encoding="utf-8",
    )

    imported, skipped = import_already_printed_csv(
        csv_path,
        "31/12/2025",
        rows_ddn_lookup=None,
        db_path=db_path,
    )

    assert imported == 2
    assert skipped == 0

    with connect(db_path) as cn:
        rows = cn.execute(
            "SELECT nom, prenom FROM prints WHERE expire=? ORDER BY nom",
            ("31/12/2025",),
        ).fetchall()

    assert [tuple(row) for row in rows] == [("ALPHA", "TEST"), ("BETA", "USER")]


def test_record_print_stores_email(tmp_path):
    db_path = tmp_path / "app.db"
    init_db(db_path)

    with connect(db_path) as cn:
        record_print(
            cn,
            "Alpha",
            "User",
            "01/01/2000",
            "31/12/2025",
            "alpha@example.com",
            zpl=None,
            status="printed",
        )
        row = cn.execute("SELECT email FROM prints LIMIT 1").fetchone()

    assert row["email"] == "alpha@example.com"


def test_parse_validation_workbook_reformats(tmp_path):
    data = pd.DataFrame(
        {
            "Nom de famille": ["Dùpont", ""],
            "Prenom": ["Élise", ""],
            "Date naissance": [pd.Timestamp(1990, 5, 1), None],
            "Date fin": ["2026-12-31", None],
            "Mail": ["elise@example.com", None],
            "Validation": ["Oui", None],
            "Montant payé": [12.5, None],
        }
    )
    path = tmp_path / "validation.xlsx"
    data.to_excel(path, index=False)

    rows = parse_validation_workbook(path)

    assert len(rows) == 1
    row = rows[0]
    assert row["Nom"] == "DUPONT"
    assert row["Prénom"] == "ELISE"
    assert row["Date_de_naissance"] == "01/05/1990"
    assert row["Expire_le"] == "31/12/2026"
    assert row["Email"] == "elise@example.com"
    assert row["Montant"] == "12.5"
    assert row["ErreurValide"] == "Oui"


def test_parse_validation_workbook_handles_libreoffice_single_row(tmp_path):
    data = pd.DataFrame(
        {
            "Nom": ["Nom\nDùpont\nAlpha"],
            "Prénom": ["Prénom\nÉlise\nBob"],
            "Date naissance": ["Date de naissance\n01/05/1990\n02/06/1992"],
            "Date fin": ["Date fin\n2026-12-31\n2026-11-30"],
            "Adresse mail": ["Adresse mail\nelise@example.com\nbob@example.com"],
            "Validation": ["Validation\nOui\nNon"],
            "Montant payé": ["Montant payé\n12,5\n30,75"],
            "Image 1": ["Picture 1"],
        }
    )
    path = tmp_path / "validation_libreoffice.xlsx"
    data.to_excel(path, index=False)

    rows = parse_validation_workbook(path)

    assert len(rows) == 2
    first, second = rows
    assert first["Nom"] == "DUPONT"
    assert first["Prénom"] == "ELISE"
    assert first["Date_de_naissance"] == "01/05/1990"
    assert first["Expire_le"] == "31/12/2026"
    assert first["Email"] == "elise@example.com"
    assert first["Montant"] == "12,5"
    assert first["ErreurValide"] == "Oui"

    assert second["Nom"] == "ALPHA"
    assert second["Prénom"] == "BOB"
    assert second["Date_de_naissance"] == "02/06/1992"
    assert second["Expire_le"] == "30/11/2026"
    assert second["Email"] == "bob@example.com"
    assert second["Montant"] == "30,75"
    assert second["ErreurValide"] == "Non"


def test_apply_validation_updates_updates_and_adds():
    existing = [
        {
            "Nom": "ALPHA",
            "Prénom": "TEST",
            "Date_de_naissance": "01/01/2000",
            "Expire_le": "31/12/2025",
            "Email": "",
            "Montant": "",
            "ErreurValide": "",
            "Derniere": "2024-01-01T12:00:00",
            "Compteur": 2,
        }
    ]
    updates = [
        {
            "Nom": "Alpha",
            "Prénom": "Test",
            "Date_de_naissance": "01/01/2000",
            "Email": "alpha@example.com",
            "Montant": "42",
            "Expire_le": "31/12/2025",
            "ErreurValide": "Yes",
        },
        {
            "Nom": "Beta",
            "Prénom": "User",
            "Date_de_naissance": "",
            "Expire_le": "31/12/2026",
            "Montant": "42",
            "ErreurValide": "No",
        },
    ]

    merged, updated, added = apply_validation_updates(existing, updates)

    assert updated == 1
    assert added == 1
    assert merged[0]["Email"] == "alpha@example.com"
    assert merged[0]["Montant"] == "42"
    assert merged[0]["ErreurValide"] == "true"
    assert any(r["Nom"] == "BETA" and r["Prénom"] == "USER" for r in merged)
    new_row = next(r for r in merged if r["Nom"] == "BETA")
    assert new_row["Compteur"] == 0
    assert new_row["Derniere"] == ""


def test_apply_validation_updates_marks_error_when_expire_differs():
    existing = [
        {
            "Nom": "ALPHA",
            "Prénom": "TEST",
            "Date_de_naissance": "01/01/2000",
            "Expire_le": "31/12/2025",
            "Email": "",
            "Montant": "",
            "ErreurValide": "",
            "Derniere": "",
            "Compteur": 0,
        }
    ]
    updates = [
        {
            "Nom": "Alpha",
            "Prénom": "Test",
            "Date_de_naissance": "01/01/2000",
            "Expire_le": "30/11/2025",
            "Montant": "15",
        }
    ]

    merged, updated, added = apply_validation_updates(existing, updates)

    assert updated == 1
    assert added == 0
    assert merged[0]["Montant"] == "15"
    assert merged[0]["ErreurValide"] == "false"


def test_parse_validation_three_line_file_from_csv(tmp_path):
    source = tmp_path / "adhesions.csv"
    source.write_text(
        "DÛPÔNT Élise\n"
        "Payé par AUBER Yann\n"
        "Montant : 45,5 €\n\n"
        "MARTIN Paul\n"
        "Payé\n"
        "Montant : 30 €\n",
        encoding="utf-8",
    )

    export_dir = tmp_path / "export"
    result = parse_validation_three_line_file(source, output_dir=export_dir)

    assert result.export_path is not None
    assert result.export_path.parent == export_dir
    assert len(result.rows) == 2

    first, second = result.rows
    assert first == {
        "nom": "DUPONT",
        "prenom": "ELISE",
        "valide_par": "Yann AUBER",
        "montant": "45.50",
    }
    assert second == {
        "nom": "MARTIN",
        "prenom": "PAUL",
        "valide_par": "",
        "montant": "30.00",
    }

    content = result.export_path.read_text(encoding="utf-8").strip().splitlines()
    assert content[0] == "nom;prenom;valide_par;montant"
    assert "DUPONT;ELISE;Yann AUBER;45.50" in content[1:]
    assert "MARTIN;PAUL;;30.00" in content[1:]


def test_parse_validation_three_line_file_from_excel(tmp_path):
    data = pd.DataFrame(
        {
            0: [
                "DUPONT Élise",
                "Payé par AUBER Yann",
                "Montant : 45 €",
                "",
                "MARTIN Paul",
                "Payé",
                "Montant : 30 €",
            ]
        }
    )
    source = tmp_path / "adhesions.xlsx"
    data.to_excel(source, index=False, header=False)

    export_dir = tmp_path / "export"
    result = parse_validation_three_line_file(source, output_dir=export_dir)

    assert result.export_path is not None
    assert result.export_path.parent == export_dir
    assert len(result.rows) == 2
    assert any(row["nom"] == "DUPONT" and row["prenom"] == "ELISE" for row in result.rows)


def test_parse_validation_three_line_file_handles_hyphenated_first_name(tmp_path):
    source = tmp_path / "hyphen.csv"
    source.write_text(
        "DUPONT JEAN-PAUL\n"
        "Payé par MARTIN Alice\n"
        "Montant : 12 €\n",
        encoding="utf-8",
    )

    result = parse_validation_three_line_file(source, output_dir=tmp_path)

    assert result.rows == [
        {
            "nom": "DUPONT",
            "prenom": "JEAN-PAUL",
            "valide_par": "Alice MARTIN",
            "montant": "12.00",
        }
    ]


def test_parse_validation_three_line_file_collapses_spaced_hyphen(tmp_path):
    source = tmp_path / "spaced-hyphen.csv"
    source.write_text(
        "DU - PONT JEAN - PAUL\n"
        "Payé par MARTIN Alice\n"
        "Montant : 12 €\n",
        encoding="utf-8",
    )

    result = parse_validation_three_line_file(source, output_dir=tmp_path)

    assert result.rows == [
        {
            "nom": "DU-PONT",
            "prenom": "JEAN-PAUL",
            "valide_par": "Alice MARTIN",
            "montant": "12.00",
        }
    ]


def test_parse_validation_three_line_file_skips_lines_until_next_name(tmp_path):
    source = tmp_path / "extra_lines.csv"
    source.write_text(
        "DUPONT Elise\n"
        "Payé par CLUB TEST\n"
        "Montant : 45 €\n"
        "Informations diverses\n"
        "Autre ligne\n"
        "MARTIN Paul\n"
        "Payé\n"
        "Montant : 30 €\n",
        encoding="utf-8",
    )

    result = parse_validation_three_line_file(source, output_dir=tmp_path)

    assert [row["nom"] for row in result.rows] == ["DUPONT", "MARTIN"]


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
