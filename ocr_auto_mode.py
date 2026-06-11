import os
import tempfile
from collections.abc import Callable
from typing import Any

from file_converter import convert_with_markitdown, parse_page_range
from ocr_converter import ocr_pdf_to_searchable_pdf


def selected_pdf_pages(pages: list[dict[str, Any]],
                       page_range: str | None) -> list[dict[str, Any]]:
    page_indexes = parse_page_range(page_range) if page_range else None
    if page_indexes is None:
        return pages

    selected_page_numbers = {page_index + 1 for page_index in page_indexes}
    return [
        page for page in pages
        if page.get("page_number") in selected_page_numbers
    ]


def pdf_pages_without_text_layer(
    pages: list[dict[str, Any]],
    page_range: str | None = None,
) -> list[int]:
    pages = selected_pdf_pages(pages, page_range)
    return [
        page["page_number"] for page in pages
        if not page.get("has_text_layer")
    ]


def convert_pdf_with_optional_ocr(
    pdf_path: str,
    page_range: str | None,
    pages: list[dict[str, Any]],
    convert_func: Callable[..., str] = convert_with_markitdown,
    ocr_func: Callable[[str, str], None] = ocr_pdf_to_searchable_pdf,
) -> tuple[str, dict[str, Any]]:
    pages_without_text = pdf_pages_without_text_layer(pages, page_range)
    if not pages_without_text:
        markdown = convert_func(pdf_path, page_range=page_range)
        return markdown, {
            "mode": "auto",
            "status": "not_needed",
            "message": "OCR auto: текстовый слой найден, OCR не нужен.",
            "pages_without_text_layer": [],
        }

    tmp_ocr_path = _make_temp_pdf_path()
    try:
        try:
            ocr_func(pdf_path, tmp_ocr_path)
        except Exception as e:
            raise RuntimeError(f"OCR auto: ошибка OCR: {e}") from e
        markdown = convert_func(tmp_ocr_path, page_range=page_range)
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
        try:
            os.unlink(tmp_ocr_path)
        except OSError:
            pass


def _make_temp_pdf_path() -> str:
    fd, path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    try:
        os.unlink(path)
    except OSError:
        pass
    return path
