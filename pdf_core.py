"""Провайдер-агностичное ядро PDF -> Markdown.

Единая точка входа для извлечения текста из PDF (bytes -> markdown),
пригодная для вызова из любого контекста: без Streamlit и без кэшей.
"""

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol

from file_converter import analyze_pdf_pages, convert_with_markitdown
from ocr_auto_mode import (
    convert_pdf_with_optional_ocr,
    pdf_pages_without_text_layer,
    selected_pdf_pages,
)
from ocr_converter import ocr_pdf_to_searchable_pdf


# ПРАВКА #60: провайдер отдаёт не голый markdown, а результат со сведениями о страницах.
@dataclass(frozen=True)
class PageInfo:
    number: int                      # 1-based
    has_text_layer: bool
    ocr_applied: bool
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class OcrResult:
    markdown: str
    content_list: list | None        # блоки MinerU; None — провайдер блоков не даёт
    pages: list[PageInfo]
    provider: str                    # "ocrmypdf" | "mineru"
    model_version: str | None        # "vlm" | "pipeline" | None
    raw_dir: Path | None             # папка с распакованным сырьём; None — сырья нет


class OcrProvider(Protocol):
    """Интерфейс OCR-провайдера: скан-PDF (bytes) -> OcrResult."""

    def ocr_pdf(self, pdf_bytes: bytes,
                page_range: str | None = None) -> OcrResult: ...


# ПРАВКА #86: сверка фрагмента текста с изображением страницы (этап B).
VERDICTS = ("agree", "fix", "unreadable")


@dataclass(frozen=True)
class VerifyResult:
    verdict: str                 # одно из VERDICTS
    correction: str | None       # весь фрагмент в исправленном виде; только при verdict == "fix"
    confidence: float            # 0–1, самооценка модели (см. PLACEHOLDER 4)
    raw: str                     # сырой текст ответа модели


# ПРАВКА #88: модель только переписывает вырезки; вердикт считает ocr.measure.judge по тексту тракта.
# GeminiVerifier.verify (#86) в коде остаётся, но этому протоколу не соответствует: в замере не участвует.
class Verifier(Protocol):
    """Вырезки одной страницы + вопрос -> дословный текст каждой, в том же порядке ("" — текста нет)."""

    def transcribe(self, images: list[bytes], question: str) -> list[str]: ...


class OcrmypdfProvider:
    """Текущий движок: ocrmypdf -> searchable PDF -> markitdown."""

    def __init__(self, ocr_func=ocr_pdf_to_searchable_pdf,
                 convert_func=convert_with_markitdown,
                 analyze_func=analyze_pdf_pages):
        self._ocr_func = ocr_func
        self._convert_func = convert_func
        self._analyze_func = analyze_func

    # ПРАВКА #60: метод переименован и отдаёт OcrResult; страницы — из analyze_func.
    def ocr_pdf(self, pdf_bytes: bytes,
                page_range: str | None = None) -> OcrResult:
        src_path = _write_temp_pdf(pdf_bytes)
        ocr_path = _make_temp_pdf_path()
        try:
            pages = selected_pdf_pages(self._analyze_func(src_path), page_range)
            self._ocr_func(src_path, ocr_path)
            markdown = self._convert_func(ocr_path, page_range=page_range)
        finally:
            _unlink_quiet(src_path)
            _unlink_quiet(ocr_path)
        return OcrResult(
            markdown=markdown,
            content_list=None,
            # движок запускается с --skip-text: OCR получают страницы без текстового слоя
            pages=[PageInfo(number=page["page_number"],
                            has_text_layer=bool(page["has_text_layer"]),
                            ocr_applied=not page["has_text_layer"])
                   for page in pages],
            provider="ocrmypdf",
            model_version=None,
            raw_dir=None,
        )


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
        # ПРАВКА #60: провайдер отдаёт OcrResult; статусу нужен только markdown.
        markdown = provider.ocr_pdf(pdf_bytes, page_range).markdown
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
