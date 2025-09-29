from __future__ import annotations

from pathlib import Path
import typer

from .io_utils import lire_tableau
from .zpl import genere_zpl, ecrire_sorties
from .db import init_db, connect, already_printed, record_print, list_prints as db_list_prints

app = typer.Typer(add_completion=False)


@app.command()
def initdb(
    db: Path = typer.Option(Path("data/app.db"), help="Chemin de la base SQLite"),
):
    """
    Crée/initialise la base SQLite (unicité par personne + expiration).
    """
    # init_db crée le schéma si besoin
    init_db(db)
    typer.echo(f"Base initialisée : {db}")


@app.command()
def simulate(
    fichier: Path = typer.Argument(..., help="Chemin du CSV ou Excel"),
    sorties: Path = typer.Option(Path("data/sorties"), help="Dossier de sortie des .zpl"),
    expiration: str = typer.Option(
        "31/12/2026", help="Date d’expiration à garder (JJ/MM/AAAA)"
    ),
    db: Path = typer.Option(Path("data/app.db"), help="Base SQLite pour le suivi"),
    force: bool = typer.Option(
        False, help="Ignorer l'unicité (regénère même si déjà imprimé)"
    ),
    log: bool = typer.Option(
        True, help="Journaliser chaque génération dans la base SQLite"
    ),
):
    """
    Lit le fichier, filtre par date d'expiration, génère les ZPL et
    enregistre le suivi/unicité dans SQLite.
    """
    # Lecture du tableau
    df = lire_tableau(fichier)

    # Filtre : on ne garde que les lignes à imprimer (réinscrits)
    df = df[df["Expire_le"].str.strip() == expiration]

    if df.empty:
        typer.echo("Aucune ligne à imprimer pour cette date d’expiration.")
        raise typer.Exit(code=0)

    # Connexion DB (création fichier si absent)
    cn = connect(db)

    # Filtrer ce qui n'a pas encore été imprimé (sauf --force)
    a_generer = []
    for r in df.to_dict(orient="records"):
        nom = r["Nom"].strip()
        prenom = r["Prénom"].strip()
        ddn = r["Date_de_naissance"].strip()
        expire = r["Expire_le"].strip()

        if force or not already_printed(cn, nom, prenom, ddn, expire):
            a_generer.append(r)
        else:
            typer.echo(f"SKIP (déjà imprimé) : {nom} {prenom} [{ddn}] {expire}")

    if not a_generer:
        typer.echo(
            "Rien à générer (tout déjà imprimé). Utilise --force pour regénérer."
        )
        raise typer.Exit(code=0)

    # Génération des fichiers .zpl (simulation) + écriture disque
    fichiers = genere_zpl(a_generer)
    ecrire_sorties(sorties, fichiers)

    # Journalisation (status = simulated)
    if log:
        with cn:
            for (fname, contenu), r in zip(fichiers, a_generer):
                record_print(
                    cn,
                    r["Nom"],
                    r["Prénom"],
                    r["Date_de_naissance"],
                    r["Expire_le"],
                    contenu,
                    status="simulated",
                )

    typer.echo(f"OK – {len(fichiers)} fichiers .zpl écrits dans {sorties}.")


@app.command()
def listprints(
    db: Path = typer.Option(Path("data/app.db"), help="Base SQLite"),
    expiration: str | None = typer.Option(
        None, help="Filtrer par date d'expiration (JJ/MM/AAAA)"
    ),
):
    """
    Liste l’historique des impressions enregistrées dans SQLite.
    """
    cn = connect(db)
    rows = db_list_prints(cn, expiration)

    if not rows:
        typer.echo("(aucune impression enregistrée)")
        raise typer.Exit(code=0)

    for row in rows:
        typer.echo(
            f"{row['printed_at']} | "
            f"{row['nom']} {row['prenom']} [{row['ddn']}] {row['expire']} "
            f"| status={row['status']}"
        )


if __name__ == "__main__":
    app()

