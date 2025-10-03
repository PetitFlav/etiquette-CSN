#!/usr/bin/env python3
"""
Impression Brother QL-570 (Linux/Windows) – simple et robuste.

Dépendances (déjà dans ton requirements) :
  pip install brother_ql Pillow pyusb

Linux (si USB) :
  sudo apt install -y libusb-1.0-0
  # (Optionnel) règle udev 04f9 pour éviter sudo.

Usage examples :
  # 1) Test simple
  python print_ql570.py --nom Dupont --prenom Marie --ddn 14/02/1990 --expire 31/12/2026

  # 2) Largeur différente (29 mm rouleau DK-22210)
  python print_ql570.py --label 29 --nom Martin --prenom Luc --ddn 03/07/1988 --expire 31/12/2025

  # 3) Backend forcé (Linux: linux_kernel si exposée /dev/usb/lp0)
  python print_ql570.py --backend linux_kernel --device /dev/usb/lp0 --nom Test --prenom Alice --ddn 01/01/1990 --expire 31/12/2026

  # 4) Windows : backend win32, device = nom d’imprimante
  python print_ql570.py --backend win32 --device "Brother QL-570" --nom Bob --prenom Eva --ddn 02/02/1992 --expire 31/12/2026
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional
from pathlib import Path
from datetime import datetime
import sys

from brother_ql.backends import backend_factory
from brother_ql.conversion import convert
from PIL import Image, ImageDraw, ImageFont
import typer

app = typer.Typer(add_completion=False)

# ----------------- Config par défaut -----------------
MODEL = "QL-570"   # ton modèle
# Largeur tête thermique QL en pixels à 300dpi pour ruban 62mm : ~696 px
# Pour 29mm, on peut rester sur 306 px de hauteur utile.
DEFAULT_CANVAS = {
    62: (696, 300),   # (largeur px, hauteur px)
    38: (413, 300),
    29: (306, 300),
    12: (118, 300),
}
# ------------------------------------------------------

@dataclass
class PrintParams:
    model: str = MODEL
    label: str = "62"              # "62"=DK-22205 (continu 62mm), "29"=DK-22210, etc.
    backend: str = "pyusb"         # "pyusb" | "linux_kernel" | "win32"
    device: Optional[str] = None   # None -> auto; USB path (/dev/usb/lp0) ; "Brother QL-570" (Windows)
    rotate: str = "auto"           # auto | 0 | 90 | 180 | 270
    threshold: int = 70            # noir/blanc
    dpi_600: bool = False          # QL-570 = 300dpi → False
    red: bool = False              # QL-570 n’imprime pas en rouge (False)

_FONT_CANDIDATES = (
    (
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ),
    (
        Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
        Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    ),
    (
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/arialbd.ttf"),
    ),
)


def _find_font(*, size: int = 36, bold: bool = False) -> Optional[ImageFont.FreeTypeFont]:
    """
    Essaie de charger une police TrueType lisible si présente, sinon utilise la police PIL par défaut.
    """
    for regular_path, bold_path in _FONT_CANDIDATES:
        target = bold_path if bold and bold_path else regular_path
        if not target:
            continue
        if target.exists():
            try:
                return ImageFont.truetype(str(target), size)
            except Exception:
                continue
    if bold:
        return _find_font(size=size, bold=False)
    return None  # PIL default bitmap font sera utilisée


def _line_height(font: Optional[ImageFont.ImageFont]) -> int:
    if font is None:
        return 40
    try:
        ascent, descent = font.getmetrics()
        return ascent + descent
    except Exception:
        try:
            bbox = font.getbbox("Ag")
            return bbox[3] - bbox[1]
        except Exception:
            return getattr(font, "size", 40)

def saison_from_expire(expire: str) -> str:
    # expire attendu "JJ/MM/AAAA"
    try:
        dt = datetime.strptime(expire, "%d/%m/%Y")
        return f"{dt.year-1} / {dt.year}"
    except Exception:
        return ""


def make_label_image(nom: str, prenom: str, ddn: str, expire: str, label_mm: int) -> Image.Image:
    """
    Étiquette simplifiée : Nom, Prénom, Saison.
    """
    canvas = DEFAULT_CANVAS.get(label_mm, DEFAULT_CANVAS[62])
    W, H = canvas
    img = Image.new("1", (W, H), 1)  # blanc
    d = ImageDraw.Draw(img)

    # Police
    f_big = _find_font()
    f_small = _find_font(size=28)
    saison_font_size = max(getattr(f_big, "size", 36), getattr(f_small, "size", 28))
    f_saison = _find_font(size=saison_font_size, bold=True) or f_big or f_small

    # Texte
    saison = saison_from_expire(expire)
    line1 = nom.upper()
    line2 = prenom.capitalize()
    line3 = f"Saison : {saison}" if saison else f"Saison : {expire}"

    # Placement
    line_height_big = _line_height(f_big)

    y = 10
    d.text((10, y), line1, fill=0, font=f_big)
    y += line_height_big + 10
    d.text((10, y), line2, fill=0, font=f_big)
    y += line_height_big + 10
    y += line_height_big
    d.text((10, y), line3, fill=0, font=f_saison)

    return img

def _open_printer(params: PrintParams):
    backend = backend_factory(params.backend)
    if params.device:
        # Ouvre explicitement la cible (Windows: nom d’imprimante ; Linux: /dev/usb/lpX)
        return backend.open(params.device)
    # Auto-détection USB (pyusb)
    devices = backend.enumerate()
    if not devices:
        raise RuntimeError("Aucune imprimante Brother QL détectée. "
                           "Spécifie --device et/ou vérifie le backend (--backend pyusb|linux_kernel|win32).")
    return backend.open(devices[0])

def _render_to_bin(img: Image.Image, params: PrintParams) -> bytes:
    render = convert(
        model=params.model,
        images=[img],
        label=params.label,           # ex. "62", "29", "38"…
        rotate=params.rotate,
        threshold=params.threshold,
        dpi_600=params.dpi_600,
        red=params.red,
    )
    return render.output

def print_one(nom: str, prenom: str, ddn: str, expire: str, params: PrintParams) -> None:
    """
    Génère l'image et l'envoie à l'imprimante.
    """
    img = make_label_image(nom, prenom, ddn, expire, int(params.label))
    printer = _open_printer(params)
    data = _render_to_bin(img, params)
    printer.write(data)

@app.command()
def print_label(
    nom: str = typer.Option(..., help="Nom"),
    prenom: str = typer.Option(..., help="Prénom"),
    ddn: str = typer.Option(..., help="Date de naissance JJ/MM/AAAA"),
    expire: str = typer.Option(..., help="Date d'expiration JJ/MM/AAAA"),
    label: str = typer.Option("62", help="Largeur du ruban (mm): 62, 38, 29, 12…"),
    backend: str = typer.Option("pyusb", help="Backend: pyusb | linux_kernel | win32"),
    device: Optional[str] = typer.Option(None, help="Cible explicite: /dev/usb/lp0 (Linux) ou nom imprimante (Windows)"),
    rotate: str = typer.Option("auto", help="Rotation: auto | 0 | 90 | 180 | 270"),
):
    """
    Imprime UNE étiquette sur Brother QL-570.
    """
    params = PrintParams(
        model=MODEL,
        label=label,
        backend=backend,
        device=device,
        rotate=rotate,
        threshold=70,
        dpi_600=False,
        red=False,
    )
    try:
        print_one(nom, prenom, ddn, expire, params)
        typer.echo("OK – Étiquette envoyée à l'imprimante.")
    except Exception as e:
        typer.secho(f"ERREUR: {e}", fg=typer.colors.RED)
        raise typer.Exit(code=1)

if __name__ == "__main__":
    app()
