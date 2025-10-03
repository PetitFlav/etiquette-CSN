from __future__ import annotations

from pathlib import Path

from PIL import Image

from src.app import printing


def test_saison_from_expire_handles_valid_date():
    assert printing.saison_from_expire("15/09/2024") == "2023 / 2024"


def test_find_font_uses_bold_variant_when_available(monkeypatch):
    fake_regular = Path("/tmp/fake-regular.ttf")
    fake_bold = Path("/tmp/fake-bold.ttf")
    monkeypatch.setattr(printing, "_FONT_CANDIDATES", ((fake_regular, fake_bold),))

    def fake_exists(self: Path) -> bool:  # noqa: D401 - simple stub
        return self in {fake_regular, fake_bold}

    def fake_truetype(path: str, size: int):  # noqa: ANN001
        class DummyFont:
            def __init__(self, path: str, size: int) -> None:
                self.path = path
                self.size = size

            def getmetrics(self) -> tuple[int, int]:
                return self.size, self.size // 4

        return DummyFont(path, size)

    monkeypatch.setattr(Path, "exists", fake_exists, raising=False)
    monkeypatch.setattr(printing.ImageFont, "truetype", fake_truetype)

    font = printing._find_font(size=20, bold=True)
    assert font is not None
    assert str(font.path).endswith("fake-bold.ttf")
    assert font.size == 20


def test_make_label_image_simple_requests_bold_font(monkeypatch):
    calls: list[tuple[int, bool]] = []
    original_find_font = printing._find_font

    def tracked_find_font(*, size: int = 36, bold: bool = False):  # type: ignore[override]
        calls.append((size, bold))
        return original_find_font(size=size, bold=bold)

    monkeypatch.setattr(printing, "_find_font", tracked_find_font)

    printing.make_label_image_simple("Nom", "Prenom", "01/01/2025")

    bold_calls = [size for size, is_bold in calls if is_bold]
    assert bold_calls, "La police bold doit être recherchée pour la saison."
    regular_sizes = [size for size, is_bold in calls if not is_bold]
    assert regular_sizes
    assert all(size >= max(regular_sizes) for size in bold_calls)


def test_make_label_image_simple_inserts_blank_line(monkeypatch):
    positions: list[tuple[tuple[int, int], str]] = []

    class DummyFont:
        def __init__(self, size: int) -> None:
            self.size = size

        def getmetrics(self) -> tuple[int, int]:
            return self.size, self.size // 4

    def fake_find_font(*, size: int = 36, bold: bool = False):  # type: ignore[override]
        return DummyFont(size)

    class DummyDraw:
        def __init__(self, _img: Image.Image) -> None:
            pass

        def text(self, xy: tuple[int, int], text: str, *, fill: int, font: DummyFont) -> None:
            positions.append((xy, text))

    monkeypatch.setattr(printing, "_find_font", fake_find_font)
    monkeypatch.setattr(printing.ImageDraw, "Draw", lambda img: DummyDraw(img))

    printing.make_label_image_simple("Nom", "Prenom", "01/01/2025")

    assert [label for (_, label) in positions] == ["NOM", "Prenom", "Saison : 2024 / 2025"]

    line_height = printing._line_height(fake_find_font())
    first_gap = positions[1][0][1] - positions[0][0][1]
    second_gap = positions[2][0][1] - positions[1][0][1]

    assert first_gap == line_height + 10
    assert second_gap == line_height + 10 + line_height


def test_print_ql570_direct_builds_payload(monkeypatch):
    called = {}

    def fake_make_label_image(nom, prenom, expire, label_mm):  # noqa: ANN001
        called["image"] = (nom, prenom, expire, label_mm)
        return Image.new("1", (10, 10), 1)

    def fake_convert(**kwargs):  # noqa: ANN001
        called["convert"] = kwargs

        class Result:
            output = b"payload"

        return Result()

    def fake_print_via_brotherql(backend_name, device, payload):  # noqa: ANN001
        called["print"] = (backend_name, device, payload)

    monkeypatch.setattr(printing, "make_label_image_simple", fake_make_label_image)
    monkeypatch.setattr(printing, "convert", fake_convert)
    monkeypatch.setattr(printing, "_print_via_brotherql", fake_print_via_brotherql)

    printing.print_ql570_direct(
        "Nom",
        "Prenom",
        "",
        "01/01/2025",
        backend_name="pyusb",
        device="usb",
        label="38",
        rotate="90",
    )

    assert called["image"] == ("Nom", "Prenom", "01/01/2025", 38)
    assert called["convert"]["label"] == "38"
    assert called["print"] == ("pyusb", "usb", b"payload")
