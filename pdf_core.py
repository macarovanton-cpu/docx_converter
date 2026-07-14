"""Провайдер-агностичное ядро PDF -> Markdown.

Единая точка входа для извлечения текста из PDF (bytes -> markdown),
пригодная для вызова из любого контекста: без Streamlit и без кэшей.
"""

import os
import tempfile
from typing import Any, Protocol

from file_converter import analyze_pdf_pages, convert_with_markitdown
from ocr_auto_mode import convert_pdf_with_optional_ocr, pdf_pages_without_text_layer
from ocr_converter import ocr_pdf_to_searchable_pdf


class OcrProvider(Protocol):
    """Интерфейс OCR-провайдера: скан-PDF (bytes) -> markdown."""

    def ocr_pdf_to_markdown(self, pdf_bytes: bytes,
                            page_range: str | None = None) -> str: ...


class OcrmypdfProvider:
    """Текущий движок: ocrmypdf -> searchable PDF -> markitdown."""

    def __init__(self, ocr_func=ocr_pdf_to_searchable_pdf,
                 convert_func=convert_with_markitdown):
        self._ocr_func = ocr_func
        self._convert_func = convert_func

    def ocr_pdf_to_markdown(self, pdf_bytes: bytes,
                            page_range: str | None = None) -> str:
        src_path = _write_temp_pdf(pdf_bytes)
        ocr_path = _make_temp_pdf_path()
        try:
            self._ocr_func(src_path, ocr_path)
            return self._convert_func(ocr_path, page_range=page_range)
        finally:
            _unlink_quiet(src_path)
            _unlink_quiet(ocr_path)


def pdf_to_markdown_with_status(
    pdf_bytes: bytes,
    *,
    page_range: str | None = None,
    mode: str = "auto",
    provider: OcrProvider | None = None,
) -> tuple[str, dict[str, Any] | None]:
    tmp_path = _write_temp_pdf(pdf_bytes)
    try:
        if mode != "auto":
            return convert_with_markitdown(tmp_path, page_range=page_range), None
        pages = analyze_pdf_pages(tmp_path)
        if provider is None:
            # дефолт = текущее поведение: существующий оркестратор
            return convert_pdf_with_optional_ocr(
                tmp_path, page_range=page_range, pages=pages)
        # явный провайдер — задел под второго (облачный vision-OCR)
        pages_without_text = pdf_pages_without_text_layer(pages, page_range)
        if not pages_without_text:
            # ponytail: статусы продублированы из ocr_auto_mode;
            # при втором провайдере вынести в общее место
            return convert_with_markitdown(tmp_path, page_range=page_range), {
                "mode": "auto",
                "status": "not_needed",
                "message": "OCR auto: текстовый слой найден, OCR не нужен.",
                "pages_without_text_layer": [],
            }
        markdown = provider.ocr_pdf_to_markdown(pdf_bytes, page_range)
        return markdown, {
            "mode": "auto",
            "status": "applied",
            "message": (
                "OCR auto: OCR применён "
                f"(страницы без текстового слоя: {', '.join(map(str, pages_without_text))})."
            ),
            "pages_without_text_layer": pages_without_text,
        }
    finally:
        _unlink_quiet(tmp_path)


def pdf_to_markdown(
    pdf_bytes: bytes,
    *,
    page_range: str | None = None,
    mode: str = "auto",
    provider: OcrProvider | None = None,
) -> str:
    return pdf_to_markdown_with_status(
        pdf_bytes, page_range=page_range, mode=mode, provider=provider)[0]


def _write_temp_pdf(pdf_bytes: bytes) -> str:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    try:
        tmp.write(pdf_bytes)
        return tmp.name
    finally:
        tmp.close()


def _make_temp_pdf_path() -> str:
    fd, path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    _unlink_quiet(path)
    return path


def _unlink_quiet(path: str) -> None:
    try:
        os.unlink(path)
    except OSError:
        pass
