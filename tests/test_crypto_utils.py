from src.app.crypto_utils import decrypt_secret, encrypt_secret, is_encrypted_secret


def test_encrypt_and_decrypt_roundtrip():
    encrypted = encrypt_secret("monsecret")
    assert encrypted.startswith("enc:")
    assert decrypt_secret(encrypted) == "monsecret"


def test_is_encrypted_secret():
    token = encrypt_secret("abc")
    assert is_encrypted_secret(token)
    assert not is_encrypted_secret("plain-text")


def test_decrypt_secret_with_plain_value_raises():
    try:
        decrypt_secret("plain")
    except ValueError:
        pass
    else:
        raise AssertionError("Expected decrypt_secret to raise for plaintext value")
