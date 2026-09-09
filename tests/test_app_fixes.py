"""Тесты трёх падений в app.py (запускать из docx_converter/)."""
import app


class _FakeUpload:
    """Минимальный стенд вместо streamlit UploadedFile."""

    def __init__(self, name: str, data: bytes = b"fake"):
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data


def test_page_range_ignored_for_non_pdf(monkeypatch):
    seen = []

    def fake_convert(path, page_range=None):
        seen.append(page_range)
        return "# Документ"

    monkeypatch.setattr(app, "convert_with_markitdown", fake_convert)

    result = app._convert_uploaded_file(_FakeUpload("отчёт.docx"), "1-3")

    assert seen == [None]
    assert result["error"] is None
    assert result["markdown"] == "# Документ"
    assert result["page_range"] == "all"


class _RaisingSecrets:
    """st.secrets на машине без .streamlit/secrets.toml."""

    def __contains__(self, key):
        raise FileNotFoundError("No secrets files found")


def test_drive_unavailable_without_secrets_file(monkeypatch):
    monkeypatch.setattr(app.st, "secrets", _RaisingSecrets())

    assert app._drive_secrets_available() is False


def test_drive_available_with_service_account(monkeypatch):
    monkeypatch.setattr(app.st, "secrets", {"gcp_service_account": {}})

    assert app._drive_secrets_available() is True
