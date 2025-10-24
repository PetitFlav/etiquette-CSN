from __future__ import annotations

import base64
import binascii
import hashlib


_SECRET = hashlib.sha256(b"etiquette-app-smtp").digest()


def _xor_bytes(data: bytes, key: bytes) -> bytes:
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))


def encrypt_secret(value: str) -> str:
    """Return an obfuscated representation of ``value``.

    The output is prefixed with ``enc:`` so that it can be detected later on.
    Empty values are returned as-is to avoid storing meaningless markers.
    """

    if not value:
        return ""

    payload = value.encode("utf-8")
    encrypted = _xor_bytes(payload, _SECRET)
    token = base64.urlsafe_b64encode(encrypted).decode("ascii")
    return f"enc:{token}"


def decrypt_secret(value: str) -> str:
    """Reverse :func:`encrypt_secret`.

    ``ValueError`` is raised when the input does not look like an encrypted
    token. This allows callers to distinguish between already encrypted and
    plaintext passwords.
    """

    if not is_encrypted_secret(value):
        raise ValueError("Secret is not encrypted")

    token = value.split(":", 1)[1]
    try:
        payload = base64.b64decode(
            token.encode("ascii"), altchars=b"-_", validate=True
        )
    except (binascii.Error, ValueError) as exc:  # pragma: no cover - depends on token
        raise ValueError("Encrypted secret is malformed") from exc

    decrypted = _xor_bytes(payload, _SECRET)
    return decrypted.decode("utf-8")


def is_encrypted_secret(value: str) -> bool:
    return value.startswith("enc:")


__all__ = ["encrypt_secret", "decrypt_secret", "is_encrypted_secret"]
