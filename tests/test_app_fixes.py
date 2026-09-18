"""Тесты трёх падений в app.py (запускать из docx_converter/)."""
import io
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Pt, RGBColor
from pypdf import PdfReader, PdfWriter

import app
from convert import STYLES, convert_md_to_docx
from ocr.mineru_provider import MineruAuthError, MineruError
from pdf_core import PageInfo


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


def test_every_doc_type_has_a_known_style():
    """ПРАВКА #59: опечатка в "style" отдала бы клиенту письмо в стиле ПЗ —
    ловим её здесь, а не в готовом документе."""
    for name, cfg in app.DOC_TYPES.items():
        assert "style" in cfg, f"у типа {name} нет ключа style"
        assert cfg["style"] in STYLES, f"у типа {name} неизвестный style: {cfg['style']}"


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


# --- ПРАВКА #73: MinerU в режиме «Файлы -> Markdown» -------------------------

def _make_zip(full_md: str) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("full.md", full_md)
    return buffer.getvalue()


def _fake_mineru(monkeypatch, tmp_path, **texts):
    """Фейковый провайдер вместо сети + кэш в tmp_path. Возвращает журнал вызовов."""
    calls: list[str] = []

    class FakeProvider:
        def __init__(self, mv):
            self.mv = mv

        def page_infos(self, pdf_bytes, page_range=None):
            return [PageInfo(1, False, True)]

        def fetch_raw_zip(self, pdf_bytes, page_range=None):
            calls.append(self.mv)
            return _make_zip(texts[self.mv])

    monkeypatch.setattr(app, "_mineru_provider_factory",
                        lambda api_key, status: FakeProvider)
    monkeypatch.setattr(app, "_OCR_CACHE_ROOT", tmp_path / "cache")
    return calls


def test_ocr_mode_options_hide_ocrmypdf_without_binaries():
    """ПРАВКА #73: на Streamlit Cloud бинарников нет — OCRmyPDF не предлагается."""
    assert app._ocr_mode_options(False) == ["off", "mineru"]
    assert app._ocr_mode_options(True) == ["off", "auto", "mineru"]


def test_mineru_api_key_prefers_secrets(monkeypatch):
    monkeypatch.setattr(app.st, "secrets", {"MINERU_API_KEY": "from-secrets"})
    monkeypatch.setenv("MINERU_API_KEY", "from-env")

    assert app._mineru_api_key() == "from-secrets"


class _NoSecretsFile:
    def __getitem__(self, key):
        raise FileNotFoundError("No secrets files found")


def test_mineru_api_key_falls_back_to_env(monkeypatch):
    """ПРАВКА #73: нет secrets.toml или нет ключа в нём — берём переменную окружения."""
    monkeypatch.setenv("MINERU_API_KEY", "from-env")

    monkeypatch.setattr(app.st, "secrets", _NoSecretsFile())
    assert app._mineru_api_key() == "from-env"

    monkeypatch.setattr(app.st, "secrets", {"gemini": {}})
    assert app._mineru_api_key() == "from-env"


def test_mineru_api_key_absent(monkeypatch):
    monkeypatch.setattr(app.st, "secrets", {})
    monkeypatch.delenv("MINERU_API_KEY", raising=False)

    assert app._mineru_api_key() is None


def test_mineru_mode_runs_pipeline_with_fake_provider(monkeypatch, tmp_path):
    """ПРАВКА #73: UI зовёт run_pipeline, в результате — markdown и отчёт."""
    calls = _fake_mineru(monkeypatch, tmp_path, vlm="# Договор\n\nТекст договора.")

    result = app._convert_uploaded_file(_FakeUpload("скан.pdf"), None,
                                        ocr_mode="mineru")

    assert result["error"] is None
    assert "Текст договора." in result["markdown"]
    assert result["report"]["provider"] == "mineru"
    assert result["report"]["verified"] is False
    assert calls == ["vlm"]


def test_mineru_mode_verify_runs_second_model(monkeypatch, tmp_path):
    calls = _fake_mineru(monkeypatch, tmp_path,
                         vlm="Цена 100 рублей за штуку товара.",
                         pipeline="Цена 700 рублей за штуку товара.")

    result = app._convert_uploaded_file(_FakeUpload("скан.pdf"), None,
                                        ocr_mode="mineru", verify=True)

    assert result["error"] is None
    assert calls == ["vlm", "pipeline"]
    assert result["report"]["verified"] is True
    _, low_conf = app._split_findings(result["report"])
    assert low_conf, "расхождение 100/700 должно стать находкой low_confidence"


def test_mineru_auth_error_gets_key_hint(monkeypatch, tmp_path):
    """ПРАВКА #73: ошибка ключа — текстом с подсказкой, не трейсбеком."""
    def factory(api_key, status):
        def build(model_version):
            raise MineruAuthError("A0202: token invalid")
        return build

    monkeypatch.setattr(app, "_mineru_provider_factory", factory)
    monkeypatch.setattr(app, "_OCR_CACHE_ROOT", tmp_path / "cache")

    result = app._convert_uploaded_file(_FakeUpload("скан.pdf"), None,
                                        ocr_mode="mineru")

    assert "A0202" in result["error"]
    assert "MINERU_API_KEY" in result["error"]
    assert result["report"] is None


def test_mineru_error_shown_as_text_without_hint(monkeypatch, tmp_path):
    def factory(api_key, status):
        def build(model_version):
            raise MineruError("распознавание не удалось: boom")
        return build

    monkeypatch.setattr(app, "_mineru_provider_factory", factory)
    monkeypatch.setattr(app, "_OCR_CACHE_ROOT", tmp_path / "cache")

    result = app._convert_uploaded_file(_FakeUpload("скан.pdf"), None,
                                        ocr_mode="mineru")

    assert result["error"] == "распознавание не удалось: boom"


def _blank_pdf(pages: int) -> bytes:
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=200, height=200)
    buffer = io.BytesIO()
    writer.write(buffer)
    return buffer.getvalue()


def test_pdf_page_subset_cuts_selected_pages():
    """ПРАВКА #73: run_pipeline диапазона не знает — страницы вырезает app.py."""
    pdf = _blank_pdf(3)

    subset = app._pdf_page_subset(pdf, "2-3")

    assert len(PdfReader(io.BytesIO(subset)).pages) == 2
    assert app._pdf_page_subset(pdf, None) is pdf


def test_pdf_page_subset_rejects_page_outside_pdf():
    with pytest.raises(ValueError, match="всего 3"):
        app._pdf_page_subset(_blank_pdf(3), "2-5")


def test_split_findings_separates_low_confidence():
    report = {"findings": [{"rule": "inn_checksum"}, {"rule": "low_confidence"},
                           {"rule": "gost_format"}]}

    main, low_conf = app._split_findings(report)

    assert [f["rule"] for f in main] == ["inn_checksum", "gost_format"]
    assert [f["rule"] for f in low_conf] == ["low_confidence"]
