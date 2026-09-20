"""ПРАВКА #81: единый вход тракта: PDF/DOCX/XLSX -> (markdown, report)."""

import hashlib
import os
import tempfile
from pathlib import Path

from file_converter import analyze_pdf_pages, convert_with_markitdown
from ocr.cache import CacheBackend
from ocr.cli import run_pipeline
from ocr.postprocess import postprocess
from ocr.validate import annotate as annotate_md
from ocr.validate import build_report, validate
from ocr_auto_mode import pdf_pages_without_text_layer
from pdf_core import pdf_to_markdown_with_status

ROUTES = ("scan", "text_tables", "text", "office")
OFFICE_EXTS = (".docx", ".xlsx")
MIN_TABLE_ROWS, MIN_TABLE_COLS = 2, 2       # PLACEHOLDER: пороги стартовые, на одной фикстуре


def _write_temp(data: bytes, suffix: str) -> str:
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as handle:
        handle.write(data)
    return path


def pdf_has_tables(pdf_path: str) -> bool:
    """True, если хоть на одной странице page.find_tables() нашёл таблицу
    не меньше MIN_TABLE_ROWS x MIN_TABLE_COLS. Выход на первой найденной."""
    # ponytail: стратегия find_tables по умолчанию (линии) — безрамочную таблицу не
    # увидит, такой PDF уйдёт в text (MarkItDown, плоский текст). Стратегию "text"
    # наугад не подбирать: решает человек, на фикстуре с безрамочной таблицей.
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.find_tables():
                rows = table.rows
                if (len(rows) >= MIN_TABLE_ROWS
                        and max(len(row.cells) for row in rows) >= MIN_TABLE_COLS):
                    return True
    return False


def detect_route(data: bytes, filename: str) -> str:
    ext = Path(filename).suffix.lower()
    if ext in OFFICE_EXTS:
        return "office"
    if ext != ".pdf":
        raise ValueError(f"Неподдерживаемый формат: {ext or filename!r} "
                         f"(допустимы .pdf, {', '.join(OFFICE_EXTS)})")
    tmp_path = _write_temp(data, ".pdf")
    try:
        if pdf_pages_without_text_layer(analyze_pdf_pages(tmp_path)):
            return "scan"                   # смешанный PDF — тоже scan
        return "text_tables" if pdf_has_tables(tmp_path) else "text"
    finally:
        os.unlink(tmp_path)


def ingest(data: bytes, *, source_name: str, work_dir: Path,
           engine: str = "mineru", mode: str = "vlm",
           verify: bool = False, annotate: bool = False,
           annotate_all: bool = False,
           cache: "CacheBackend | None" = None,
           provider_factory=None) -> tuple[str, dict]:
    """Какой бы конвертер ни отработал — postprocess + validate + тот же report.json."""
    pipeline = dict(source_name=source_name, work_dir=work_dir, engine=engine, mode=mode,
                    verify=verify, annotate=annotate, annotate_all=annotate_all,
                    cache=cache, provider_factory=provider_factory)
    ext = Path(source_name).suffix.lower()
    if engine == "ocrmypdf" and ext == ".pdf":
        return run_pipeline(data, **pipeline)     # явный выбор человека: без детектора
    route = detect_route(data, source_name)
    if route in ("scan", "text_tables"):
        # один вызов на оба маршрута: is_ocr провайдер считает сам по текстовому слою
        return run_pipeline(data, **pipeline)
    if verify:
        raise ValueError("--verify доступен только для маршрутов MinerU")

    if route == "text":
        md = pdf_to_markdown_with_status(data, mode="auto")[0]
    else:
        tmp_path = _write_temp(data, ext)         # по расширению выбирается конвертер
        try:
            md = convert_with_markitdown(tmp_path)
        finally:
            os.unlink(tmp_path)

    md, findings = postprocess(md, None)
    findings = findings + validate(md, None)
    report = build_report(source=source_name, sha256=hashlib.sha256(data).hexdigest(),
                          provider="markitdown", model_version=None, cache_hit=False,
                          verified=False, findings=findings, content_list=None)
    if annotate or annotate_all:                  # как в run_pipeline (ПРАВКА #70)
        md = annotate_md(md, report, include_low_confidence=annotate_all)
    return md, report
