from __future__ import annotations

from pathlib import Path
from typing import Iterable, List, Tuple

from jinja2 import Environment, FileSystemLoader, select_autoescape


# Dossier des templates Jinja2
TEMPLATES_DIR = Path(__file__).parent / "templates"
TEMPLATE_NAME = "label.zpl.j2"


def _env() -> Environment:
    """
    Prépare l'environnement Jinja2 (sans auto-escape, car ZPL = texte brut).
    """
    return Environment(
        loader=FileSystemLoader(TEMPLATES_DIR),
        autoescape=select_autoescape(enabled_extensions=()),
        trim_blocks=True,
        lstrip_blocks=True,
    )


def _sanitize_filename(name: str) -> str:
    """
    Nettoie un nom de fichier (évite espaces et caractères spéciaux gênants).
    """
    bad = r'<>:"/\\|?*'
    for ch in bad:
        name = name.replace(ch, "_")
    name = name.replace(" ", "_")
    return name


def genere_zpl(
    records: Iterable[dict],
    template_name: str = TEMPLATE_NAME,
    filename_fmt: str = "{Nom}_{Prénom}.zpl",
) -> List[Tuple[str, str]]:
    """
    Rend le template ZPL pour chaque record et retourne une liste de
    (nom_fichier, contenu_zpl).

    :param records: itérable de dicts avec clés: Nom, Prénom, Date_de_naissance, Expire_le
    :param template_name: nom du template Jinja dans src/app/templates/
    :param filename_fmt: format du nom de fichier (accède aux clés du record)
    """
    env = _env()
    template = env.get_template(template_name)

    sorties: List[Tuple[str, str]] = []
    for r in records:
        contenu = template.render(
            nom=(r.get("Nom", "") or "").strip(),
            prenom=(r.get("Prénom", "") or "").strip(),
            ddn=(r.get("Date_de_naissance", "") or "").strip(),
            expire=(r.get("Expire_le", "") or "").strip(),
        )
        raw_name = filename_fmt.format(**r)
        fname = _sanitize_filename(raw_name)
        sorties.append((fname, contenu))
    return sorties


def ecrire_sorties(sorties_dir: Path, fichiers: List[Tuple[str, str]]) -> None:
    """
    Écrit les fichiers ZPL dans le dossier de sortie.
    """
    sorties_dir = Path(sorties_dir)
    sorties_dir.mkdir(parents=True, exist_ok=True)
    for fname, contenu in fichiers:
        (sorties_dir / fname).write_text(contenu, encoding="utf-8")