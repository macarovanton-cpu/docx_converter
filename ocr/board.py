"""ПРАВКА #82: табло качества по фикстурам. Строго оффлайн."""

import json
import sys
import tempfile
import traceback
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from file_converter import get_pdf_page_count
from ocr.cache import LocalCache, build_meta, cache_key
from ocr.ingest import detect_route, ingest
from ocr.postprocess import parse_pipe_tables

BOARD_SCHEMA_VERSION = 1
FIXTURES = Path(__file__).resolve().parents[1] / "_test" / "fixtures" / "ocr"
BOARD_JSON = FIXTURES.parents[1] / "quality_board.json"      # _test/quality_board.json
DRAFTS = FIXTURES.parents[1] / "board"                       # _test/board/<stem>/
ERRORS_HEADER = "# страница | было | надо\n"

BOARD = (("bakeoff.pdf", "vlm_raw.zip"), ("bakeoff2.pdf", "vlm_raw2.zip"),
         ("bakeoff3.pdf", "vlm_raw3.zip"), ("textpdf1.pdf", "vlm_raw_textpdf1.zip"),
         ("docx1.docx", None), ("xlsx1.xlsx", None))

NO_ERRORS_FILE = ("bakeoff.pdf",)           # у неё golden.md
BROKEN_TABLE_RULES = ("html_table_unparsed", "table_merge_failed")
CONSOLE_COLUMNS = ("fixture", "route", "critical", "warning", "info",
                   "tables", "table_broken", "chars", "count_diffs")

ERRORS_TEMPLATE = ERRORS_HEADER + """\
# «было» — дословный фрагмент _test/board/{stem}/out.md, встречается в нём ровно один раз
#   (не уникален — расширить контекстом). «надо» — чем заменить; пусто = удалить.
# Вставка потерянного: «было» = соседний текст, «надо» = он же со вставкой.
# Черта внутри текста — \\| . Страница — номер в PDF; для DOCX/XLSX — «-» или имя листа.
# Строки с # и пустые игнорируются. Ошибок нет — оставить файл как есть.
# DOCX/XLSX: только потери и искажения (объединённые ячейки, нумерация, колонтитулы,
#   потерянные строки). Стиль не правим.
"""


def _offline(model_version):
    raise RuntimeError(f"табло оффлайн: нет сырого zip для {model_version}")


def _require(path: Path) -> Path:
    if not path.is_file():
        raise FileNotFoundError(str(path))
    return path


def board_row(name: str, md: str, report: dict, route: str) -> dict:
    tables = parse_pipe_tables(md)
    by_rule = Counter(item["rule"] for item in report["findings"])
    broken = sum(len(row) != len(table[0]) for table in tables for row in table)
    return {
        "fixture": name,
        "route": route,
        "provider": report["provider"],
        "critical": report["summary"]["critical"],
        "warning": report["summary"]["warning"],
        "info": report["summary"]["info"],
        "by_rule": dict(sorted(by_rule.items())),
        "tables": len(tables),
        "table_rows": sum(len(table) for table in tables),
        "table_broken": broken + sum(by_rule[rule] for rule in BROKEN_TABLE_RULES),
        "chars": len(md),
        "tokens": len(md.split()),
        "count_diffs": None,                # заполняет спека 10
        "threshold": None,
    }


def build_board(work_dir: Path, *, fixtures: Path = FIXTURES) -> tuple[dict, dict[str, tuple[str, dict]]]:
    # временный кэш: рабочий .cache/ocr/ табло не читает и не пишет
    cache = LocalCache(work_dir / "cache")
    rows, outputs = [], {}
    for name, zip_name in BOARD:
        path = _require(fixtures / name)
        data = path.read_bytes()
        if zip_name is not None and (fixtures / zip_name).is_file():
            # нет zip — не FileNotFoundError: промах кэша сам упадёт в _offline
            zip_bytes = (fixtures / zip_name).read_bytes()
            key = cache_key(data, "mineru", "vlm")
            cache.put(key, zip_bytes, build_meta(
                key, data, zip_bytes, provider="mineru", model_version="vlm",
                page_range=None, pages=get_pdf_page_count(str(path))))
        md, report = ingest(data, source_name=name, work_dir=work_dir / path.stem,
                            cache=cache, provider_factory=_offline)
        outputs[name] = (md, report)
        rows.append(board_row(name, md, report, detect_route(data, name)))
    board = {"schema_version": BOARD_SCHEMA_VERSION,
             "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
             "rows": rows}
    return board, outputs


def _print_table(rows: list[dict]) -> None:
    lines = [CONSOLE_COLUMNS] + [tuple("-" if row[col] is None else str(row[col])
                                       for col in CONSOLE_COLUMNS) for row in rows]
    widths = [max(len(line[i]) for line in lines) for i in range(len(CONSOLE_COLUMNS))]
    for line in lines:
        print("  ".join(cell.ljust(width) for cell, width in zip(line, widths)).rstrip())


def main(argv: "list[str] | None" = None) -> int:
    try:
        # глобалы читаются в момент вызова: тесты подменяют пути на tmp_path
        with tempfile.TemporaryDirectory() as work:
            board, outputs = build_board(Path(work), fixtures=FIXTURES)
        _print_table(board["rows"])
        BOARD_JSON.parent.mkdir(parents=True, exist_ok=True)
        BOARD_JSON.write_text(json.dumps(board, ensure_ascii=False, indent=2), encoding="utf-8")
        for name, (md, report) in outputs.items():
            stem = Path(name).stem
            draft = DRAFTS / stem
            draft.mkdir(parents=True, exist_ok=True)
            (draft / "out.md").write_text(md, encoding="utf-8")
            (draft / "report.json").write_text(
                json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            errors = FIXTURES / f"{stem}.errors.txt"
            if name not in NO_ERRORS_FILE and not errors.exists():   # ручной труд не трогать
                errors.write_text(ERRORS_TEMPLATE.format(stem=stem), encoding="utf-8")
        return 0
    except Exception:
        traceback.print_exc(file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
