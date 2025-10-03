from __future__ import annotations

from pathlib import Path
from typing import Protocol

from brother_ql.backends import backend_factory
from brother_ql.conversion import convert
from brother_ql.raster import BrotherQLRaster
from PIL import Image, ImageDraw, ImageFont

MODEL = "QL-570"
DEFAULT_CANVAS = {
    62: (696, 300),
    38: (413, 300),
    29: (306, 300),
    12: (118, 300),
}


class BrotherHandle(Protocol):
    def write(self, data: bytes) -> object:  # pragma: no cover - protocol definition
        ...


def saison_from_expire(expire: str) -> str:
    """Return ``"AAAA / AAAA+1"`` computed from a ``JJ/MM/AAAA`` date."""
    from datetime import datetime

    try:
        dt = datetime.strptime(expire, "%d/%m/%Y")
        return f"{dt.year-1} / {dt.year}"
    except Exception:  # pragma: no cover - defensive fallback
        return ""


_FONT_CANDIDATES: tuple[tuple[Path, Path | None], ...] = (
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


def _find_font(*, size: int = 36, bold: bool = False) -> ImageFont.ImageFont | None:
    """Attempt to load a readable TTF font, otherwise fallback to PIL's default."""

    for regular_path, bold_path in _FONT_CANDIDATES:
        target = bold_path if bold and bold_path else regular_path
        if not target:
            continue
        try:
            if target.exists():
                return ImageFont.truetype(str(target), size)
        except Exception:  # pragma: no cover - font loading depends on OS
            continue

    if bold:
        return _find_font(size=size, bold=False)
    return None


def _line_height(font: ImageFont.ImageFont | None) -> int:
    if font is None:
        return 40
    try:
        ascent, descent = font.getmetrics()
        return ascent + descent
    except Exception:  # pragma: no cover - depends on font backend
        try:
            bbox = font.getbbox("Ag")
            return bbox[3] - bbox[1]
        except Exception:  # pragma: no cover - depends on font backend
            return getattr(font, "size", 40)


def make_label_image_simple(nom: str, prenom: str, expire: str, label_mm: int = 62) -> Image.Image:
    """Generate a simple label containing ``Nom``, ``Prénom`` and the season."""
    width, height = DEFAULT_CANVAS.get(label_mm, DEFAULT_CANVAS[62])
    img = Image.new("1", (width, height), 1)
    draw = ImageDraw.Draw(img)
    font_big = _find_font()
    font_small = _find_font(size=28)
    saison_font_size = max(getattr(font_big, "size", 36), getattr(font_small, "size", 28))
    font_saison = _find_font(size=saison_font_size, bold=True)
    if font_saison is None:
        font_saison = font_big or font_small

    line1 = (nom or "").upper()
    line2 = (prenom or "").capitalize()
    saison = saison_from_expire(expire) or expire
    line3 = f"Saison : {saison}"

    y = 10
    draw.text((10, y), line1, fill=0, font=font_big)
    y += _line_height(font_big) + 10
    draw.text((10, y), line2, fill=0, font=font_big)
    y += _line_height(font_big) + 10
    draw.text((10, y), line3, fill=0, font=font_saison)
    return img


def _open_bql_handle(backend_name: str, device: str | None) -> BrotherHandle:
    """Open the Brother QL handle depending on the selected backend."""
    backend = backend_factory(backend_name)
    normalized = backend_name.lower().strip()
    needs_device = {"linux_kernel", "network", "file"}
    if normalized in needs_device:
        if not device:
            raise RuntimeError(
                f"Le backend '{normalized}' requiert 'device' dans config.ini :\n"
                " - linux_kernel : /dev/usb/lp0\n"
                " - network      : 192.168.1.50:9100\n"
                " - file         : /chemin/sortie.bin"
            )
        return backend.open(device)

    if hasattr(backend, "enumerate"):
        targets = backend.enumerate()
        if device:
            return backend.open(device)
        if not targets:
            raise RuntimeError(
                "Aucun périphérique Brother détecté pour backend='pyusb'.\n"
                "Vérifie le câble/driver/permissions (libusb)."
            )
        return backend.open(targets[0])

    if not device:
        raise RuntimeError(
            f"Le backend '{normalized}' ne fournit pas enumerate() et aucun 'device' n'a été donné."
        )
    return backend.open(device)


def _render_label_bytes(
    nom: str,
    prenom: str,
    expire: str,
    label_mm: int,
    rotate_val: str | int = 0,
) -> tuple[bytes, Image.Image]:
    img = make_label_image_simple(nom, prenom, expire, label_mm)
    raster = BrotherQLRaster(MODEL)
    raster.exception_on_warning = True

    try:
        rotation = int(rotate_val)
    except Exception:  # pragma: no cover - fallback path
        rotation = 0

    result = convert(
        qlr=raster,
        images=[img],
        label=str(label_mm),
        rotate=rotation,
        threshold=70,
        dither=False,
        compress=False,
        red=False,
        dpi_600=False,
        hq=True,
        cut=False,
    )

    payload = result.output if hasattr(result, "output") else result
    return payload, img


def _print_via_brotherql(backend_name: str, device: str | None, payload: bytes) -> None:
    handle = _open_bql_handle(backend_name, device)
    handle.write(payload)


def _print_via_win32_driver(device_name: str, pil_image: Image.Image) -> None:
    import win32print  # type: ignore[import-not-found]
    import win32ui  # type: ignore[import-not-found]
    from PIL import ImageWin

    if not device_name:
        device_name = win32print.GetDefaultPrinter()
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(device_name)
    hdc.StartDoc("Etiquette CSN")
    hdc.StartPage()
    dib = ImageWin.Dib(pil_image.convert("RGB"))
    width, height = pil_image.size
    dib.draw(hdc.GetHandleOutput(), (0, 0, width, height))
    hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()


def print_ql570_direct(
    nom: str,
    prenom: str,
    ddn: str,
    expire: str,
    *,
    label: str = "62",
    backend_name: str = "pyusb",
    device: str | None = None,
    rotate: str = "0",
) -> None:
    label_mm = int(label or "62")
    payload, img = _render_label_bytes(nom, prenom, expire, label_mm, rotate_val=rotate)

    backend_key = backend_name.lower()
    if backend_key in {"pyusb", "linux_kernel", "network", "file", "dummy"}:
        _print_via_brotherql(backend_name, device, payload)
    elif backend_key == "win32print":
        _print_via_win32_driver(device or "", img)
    else:  # pragma: no cover - guardrail
        raise RuntimeError(f"Backend inconnu : {backend_name}")


__all__ = [
    "MODEL",
    "DEFAULT_CANVAS",
    "saison_from_expire",
    "make_label_image_simple",
    "print_ql570_direct",
]
