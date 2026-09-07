"""Do not persist, expose or bypass a rejected workbook opening password."""
from fastapi import HTTPException
from msoffcrypto.exceptions import InvalidKeyError
import pytest

from backend.routers import projects


def test_wrong_workbook_password_returns_actionable_error_without_echoing_it(monkeypatch) -> None:
    password = "synthetic-test-only"
    attempted = []

    class EncryptedBook:
        def load_key(self, *, password: str, verify_password: bool) -> None:
            attempted.append(verify_password)
            raise InvalidKeyError(password)

        def decrypt(self, *args, **kwargs):
            raise AssertionError("Do not decrypt after a rejected password")

    monkeypatch.setattr(projects, "OfficeFile", lambda content: EncryptedBook())
    with pytest.raises(HTTPException) as error:
        projects._decrypt_encrypted_ooxml(b"test encrypted bytes", password)
    assert error.value.status_code == 400
    assert "打开密码不正确" in error.value.detail
    assert password not in error.value.detail
    assert attempted == [True]
