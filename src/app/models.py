from __future__ import annotations

from pydantic import BaseModel, field_validator
from datetime import datetime


class Ligne(BaseModel):
    """
    Représente une ligne issue du CSV/Excel.
    Les colonnes doivent être normalisées avant (via io_utils).
    """

    Nom: str
    Prénom: str
    Date_de_naissance: str
    Expire_le: str

    @field_validator("Expire_le")
    @classmethod
    def check_expire(cls, v: str) -> str:
        """
        Vérifie que la date est au format JJ/MM/AAAA.
        """
        try:
            datetime.strptime(v.strip(), "%d/%m/%Y")
        except ValueError as e:
            raise ValueError(f"Format de date invalide pour Expire_le: {v}") from e
        return v

    @field_validator("Date_de_naissance")
    @classmethod
    def check_ddn(cls, v: str) -> str:
        """
        Vérifie que la date de naissance est aussi valide (JJ/MM/AAAA).
        """
        if not v.strip():
            return v  # tolère vide
        try:
            datetime.strptime(v.strip(), "%d/%m/%Y")
        except ValueError as e:
            raise ValueError(
                f"Format de date invalide pour Date_de_naissance: {v}"
            ) from e
        return v