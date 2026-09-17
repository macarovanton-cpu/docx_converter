"""Тесты трёх падений в app.py (запускать из docx_converter/)."""
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor

import app
from convert import convert_md_to_docx


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


BOM_FIXTURE = Path(__file__).resolve().parents[1] / "test_formatting_bom.md"


def test_decode_md_upload_strips_bom():
    md_text = app._decode_md_upload(BOM_FIXTURE.read_bytes())

    assert "﻿" not in md_text
    assert md_text.startswith("# ")


def test_decode_md_upload_falls_back_to_cp1251():
    assert app._decode_md_upload("Тест кириллицы".encode("cp1251")) == "Тест кириллицы"


def test_bom_fixture_first_block_renders_as_h1(tmp_path):
    """BOM не должен превращать H1 в обычный абзац."""
    md_text = app._decode_md_upload(BOM_FIXTURE.read_bytes())
    out = tmp_path / "bom.docx"
    convert_md_to_docx(md_text, str(out))

    first = Document(str(out)).paragraphs[0]
    run = first.runs[0]
    assert not first.text.startswith("#")
    assert run.font.name == "PT Sans Narrow"
    assert run.font.size == Pt(18)
    assert run.font.color.rgb == RGBColor.from_string("015198")


def test_doc_types_point_at_distinct_templates():
    """ПРАВКА #54: раньше оба типа грузили один и тот же file_id."""
    drive_ids = [cfg["drive_id"] for cfg in app.DOC_TYPES.values()]

    assert len(set(drive_ids)) == len(drive_ids)
    assert all(drive_ids)


def test_get_template_downloads_selected_drive_id(monkeypatch):
    seen = []

    def fake_download(file_id):
        seen.append(file_id)
        return "/tmp/template.docx"

    monkeypatch.setattr(app, "download_template_from_drive", fake_download)

    path = app.get_template(use_drive=True,
                            drive_file_id="1_E7eI5PgMD50MEI8RNl8xoiWmhUsOUap",
                            local_path=None)

    assert seen == ["1_E7eI5PgMD50MEI8RNl8xoiWmhUsOUap"]
    assert path == "/tmp/template.docx"
