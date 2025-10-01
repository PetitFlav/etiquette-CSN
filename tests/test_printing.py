from __future__ import annotations

from PIL import Image

from src.app import printing


def test_saison_from_expire_handles_valid_date():
    assert printing.saison_from_expire("15/09/2024") == "2023 / 2024"


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
