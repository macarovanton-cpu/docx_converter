"""ПРАВКА #66: CLI OCR-тракта: PDF -> out.md + report.json + JSON-итог в stdout."""

import argparse
import hashlib
import json
import sys
import traceback
from pathlib import Path

from ocr.cache import CacheBackend, build_meta, cache_key, make_cache
from ocr.diff import diff_findings
from ocr.mineru_provider import MineruProvider, result_from_zip
from ocr.postprocess import postprocess
from ocr.validate import annotate as annotate_md
from ocr.validate import build_report, validate
from pdf_core import OcrmypdfProvider

EXIT_OK, EXIT_ERROR, EXIT_CRITICAL = 0, 1, 2

ENGINES = ("mineru", "ocrmypdf")
MODES = ("vlm", "pipeline")
OTHER_MODE = {"vlm": "pipeline", "pipeline": "vlm"}


class _ArgumentError(Exception):
    """Ошибка разбора аргументов; argparse-овский SystemExit(2) наружу не летит."""


class _Parser(argparse.ArgumentParser):
    def error(self, message):            # код 2 занят «есть критичные находки»
        raise _ArgumentError(message)


def _default_provider_factory(model_version: str):
    return MineruProvider(model_version=model_version)


def _mineru_result(pdf_bytes: bytes, model_version: str, *, work_dir: Path,
                   cache: "CacheBackend | None", provider_factory):
    """(OcrResult, попадание в кэш). При попадании провайдер не создаётся, сети нет."""
    key = cache_key(pdf_bytes, "mineru", model_version)
    cached = cache.get(key) if cache is not None else None
    if cached is not None:
        # страниц нет: провайдера не спрашивали. В отчёт страницы идут из content_list
        zip_bytes, pages = cached[0], []
    else:
        provider = (provider_factory or _default_provider_factory)(model_version)
        zip_bytes = provider.fetch_raw_zip(pdf_bytes)
        pages = provider.page_infos(pdf_bytes)
        if cache is not None:
            cache.put(key, zip_bytes,
                      build_meta(key, pdf_bytes, zip_bytes, provider="mineru",
                                 model_version=model_version, page_range=None,
                                 pages=len(pages)))
    result = result_from_zip(zip_bytes, Path(work_dir) / "raw" / model_version,
                             model_version=model_version, pages=pages)
    return result, cached is not None


def run_pipeline(pdf_bytes: bytes, *, source_name: str, work_dir: Path,
                 engine: str = "mineru", mode: str = "vlm",
                 verify: bool = False, annotate: bool = False,
                 annotate_all: bool = False,
                 cache: "CacheBackend | None" = None,
                 provider_factory=None) -> tuple[str, dict]:
    """Весь тракт: OCR -> постпроцессор -> валидатор -> (сверка) -> (markdown, report).

    Сборка живёт здесь, а не в main: UI этапа 8 зовёт ту же функцию без изменений.
    """
    if engine not in ENGINES:
        raise ValueError(f"engine={engine!r}: допустимы {ENGINES}")

    if engine == "ocrmypdf":
        if verify:
            raise ValueError("--verify доступен только для --engine mineru")
        result, cache_hit, model_version = (
            OcrmypdfProvider().ocr_pdf(pdf_bytes), False, None)
    else:
        if mode not in MODES:
            raise ValueError(f"mode={mode!r}: допустимы {MODES}")
        result, cache_hit = _mineru_result(
            pdf_bytes, mode, work_dir=work_dir, cache=cache,
            provider_factory=provider_factory)
        model_version = mode

    # ПРАВКА #74: content_list идёт и в постпроцессор — он возвращает из него
    # блоки, которых MinerU не положил в full.md
    md, findings = postprocess(result.markdown, result.content_list)
    findings = findings + validate(md, result.content_list)
    if verify:
        second, _ = _mineru_result(
            pdf_bytes, OTHER_MODE[mode], work_dir=work_dir, cache=cache,
            provider_factory=provider_factory)
        # второму прогону — свой content_list: иначе сверка приняла бы
        # возвращённые блоки за расхождение прогонов
        second_md = postprocess(second.markdown, second.content_list)[0]
        findings = findings + diff_findings(md, second_md, result.content_list)

    report = build_report(source=source_name,
                          sha256=hashlib.sha256(pdf_bytes).hexdigest(),
                          provider=engine, model_version=model_version,
                          cache_hit=cache_hit, verified=verify,
                          findings=findings, content_list=result.content_list)
    if annotate or annotate_all:          # ПРАВКА #70: --annotate-all включает и пометки
        md = annotate_md(md, report, include_low_confidence=annotate_all)
    return md, report


def _parse_args(argv: "list[str] | None") -> argparse.Namespace:
    parser = _Parser(
        prog="python -m ocr.cli",
        description="Тендерный PDF -> Markdown со структурой + отчёт о сомнительных местах.")
    parser.add_argument("input", help="входной PDF/DOCX/XLSX")     # ПРАВКА #81
    parser.add_argument("--out", required=True, help="папка для out.md и report.json")
    parser.add_argument("--engine", default="mineru", choices=ENGINES,
                        help="движок OCR (по умолчанию mineru)")
    parser.add_argument("--mode", default="vlm", choices=MODES,
                        help="модель MinerU (по умолчанию vlm; для ocrmypdf игнорируется)")
    parser.add_argument("--verify", action="store_true",
                        help="второй прогон другим движком MinerU, расхождения -> находки")
    parser.add_argument("--annotate", action="store_true",
                        help="вставить находки в out.md как «!! ПРОВЕРИТЬ: … !!»")
    parser.add_argument("--annotate-all", action="store_true",
                        help="то же, но вместе с low_confidence (их десятки)")
    parser.add_argument("--cache", default="local", choices=("local", "drive"),
                        help="бэкенд кэша сырого ответа (по умолчанию local)")
    return parser.parse_args(argv)


def main(argv: "list[str] | None" = None, *, provider_factory=None) -> int:
    """Ровно одна строка JSON в stdout; всё остальное — в stderr. Коды 0 / 2 / 1."""
    summary = {"status": "error", "out_md": None, "report": None,
               "cache_hit": False, "findings": None, "error": None}
    try:
        # ПРАВКА #81: вход .pdf/.docx/.xlsx идёт через ingest; импорт здесь —
        # ocr.ingest сам импортирует run_pipeline из этого модуля
        from ocr.ingest import ingest

        args = _parse_args(argv)
        source = Path(args.input)
        data = source.read_bytes()
        out_dir = Path(args.out).resolve()
        cache = make_cache(args.cache)
        md, report = ingest(data, source_name=source.name,
                            work_dir=out_dir, engine=args.engine,
                            mode=args.mode, verify=args.verify,
                            annotate=args.annotate,
                            annotate_all=args.annotate_all, cache=cache,
                            provider_factory=provider_factory)
        # пишем только после успешного тракта: недописанный out.md хуже отсутствующего
        out_dir.mkdir(parents=True, exist_ok=True)
        out_md, report_path = out_dir / "out.md", out_dir / "report.json"
        out_md.write_text(md, encoding="utf-8")
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2),
                               encoding="utf-8")
        critical = report["summary"]["critical"]
        summary.update(status="findings" if critical else "ok",
                       out_md=str(out_md), report=str(report_path),
                       cache_hit=report["cache_hit"], findings=report["summary"])
        code = EXIT_CRITICAL if critical else EXIT_OK
    except Exception as exc:                 # ключ API в текст исключений не попадает
        traceback.print_exc(file=sys.stderr)
        summary["error"] = str(exc) or type(exc).__name__
        code = EXIT_ERROR
    print(json.dumps(summary, ensure_ascii=False))
    return code


if __name__ == "__main__":
    raise SystemExit(main())
