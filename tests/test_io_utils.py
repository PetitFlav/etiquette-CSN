import pytest

from src.app.io_utils import normalize_name


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("Jean - Baptiste", "JEAN-BAPTISTE"),
        ("Francois- xavier", "FRANCOIS-XAVIER"),
        ("  Marie   Louise  ", "MARIE LOUISE"),
        ("Élodie", "ELODIE"),
    ],
)
def test_normalize_name_harmonizes_spacing_and_accents(raw, expected):
    assert normalize_name(raw) == expected
