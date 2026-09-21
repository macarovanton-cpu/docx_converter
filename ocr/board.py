"""ПРАВКА #82: табло качества по фикстурам. Строго оффлайн.

ПРАВКА #83: эталоны по фикстурам (--golden), count_diffs / порог / ‰ в табло.
"""

import argparse
import difflib
import json
import re
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
                   "tables", "table_broken", "chars", "count_diffs", "threshold", "per_1000_tokens")
CONSOLE_TITLES = {"count_diffs": "diffs", "threshold": "thr", "per_1000_tokens": "‰"}   # ПРАВКА #83

# ПРАВКА #83: эталон = черновик тракта + закрытый список правок из <stem>.errors.txt
GOLDEN = {"bakeoff.pdf": "golden.md", "bakeoff2.pdf": "bakeoff2.golden.md",
          "bakeoff3.pdf": "bakeoff3.golden.md", "textpdf1.pdf": "textpdf1.golden.md",
          "docx1.docx": "docx1.golden.md", "xlsx1.xlsx": "xlsx1.golden.md"}

# остаток тракта до эталона по каждой фикстуре отдельно; факт на день сборки (21.09.2026), без запаса.
# bakeoff.pdf: 13 — сырой путь (vlm_raw.zip) против golden.md с правкой 11;
# это НЕ REMAINING_DIFFS из test_ocr_postprocess (там вход vlm.md)
THRESHOLDS = {"bakeoff.pdf": 13, "bakeoff2.pdf": 20, "bakeoff3.pdf": 8,
              "textpdf1.pdf": 13, "docx1.docx": 8, "xlsx1.xlsx": 1}

_SEPARATOR_RE = re.compile(r"^\s*\|?[\s:|-]+\|?\s*$")

ERRORS_TEMPLATE = ERRORS_HEADER + """\
# «было» — дословный фрагмент _test/board/{stem}/out.md, встречается в нём ровно один раз
#   (не уникален — расширить контекстом). «надо» — чем заменить; пусто = удалить.
# Вставка потерянного: «было» = соседний текст, «надо» = он же со вставкой.
# Черта внутри текста — \\| . Страница — номер в PDF; для DOCX/XLSX — «-» или имя листа.
# Строки с # и пустые игнорируются. Ошибок нет — оставить файл как есть.
# DOCX/XLSX: только потери и искажения (объединённые ячейки, нумерация, колонтитулы,
#   потерянные строки). Стиль не правим.
"""


def text_tokens(md: str) -> list[str]:
    r"""Токены текста без разметки.

    1. строки-разделители pipe-таблиц (^\s*\|?[\s:|-]+\|?\s*$ с хотя бы одним '-') удалить;
    2. HTML-теги <[^>]+> заменить пробелом;
    3. символ '|' заменить пробелом;
    4. str.split().
    """
    kept = [line for line in md.split("\n")
            if not ("-" in line and _SEPARATOR_RE.match(line))]
    text = re.sub(r"<[^>]+>", " ", "\n".join(kept))
    return text.replace("|", " ").split()


def count_diffs(a: str, b: str) -> int:
    """Число не-'equal' опкодов
    difflib.SequenceMatcher(None, text_tokens(a), text_tokens(b), autojunk=False)."""
    matcher = difflib.SequenceMatcher(None, text_tokens(a), text_tokens(b), autojunk=False)
    return sum(1 for tag, *_ in matcher.get_opcodes() if tag != "equal")


def parse_errors_numbered(text: str) -> list[tuple[int, tuple[str, str, str]]]:
    r"""ПРАВКА #83: строки «страница | было | надо»; черта внутри поля — \|.

    ПРАВКА #84: (номер строки файла, (страница, было, надо)) — номер как в редакторе.
    """
    errors = []
    for number, line in enumerate(text.splitlines(), 1):
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        fields = [part.strip().replace("\\|", "|") for part in re.split(r"(?<!\\)\|", line)]
        if len(fields) != 3 or not fields[1]:
            raise ValueError(f"строка {number}: нужно «страница | было | надо» с непустым «было»: {line!r}")
        errors.append((number, tuple(fields)))
    return errors


def parse_errors(text: str) -> list[tuple[str, str, str]]:
    return [error for _, error in parse_errors_numbered(text)]


def apply_errors(md: str, errors: list[tuple[str, str, str]], *,
                 lines: "list[int] | None" = None) -> str:
    """ПРАВКА #83: правки по порядку, каждая — ровно по одному вхождению, без догадок.

    ПРАВКА #84: lines — номера строк errors.txt, попадают в текст ошибки.
    """
    if lines is not None and len(lines) != len(errors):
        raise ValueError(f"lines: {len(lines)} номеров на {len(errors)} правок")
    for index, (_page, was, need) in enumerate(errors):
        n = md.count(was)
        if n != 1:
            where = "" if lines is None else f"строка {lines[index]} errors.txt: "
            raise ValueError(f"{where}«{was}»: вхождений {n}, нужно ровно 1")
        md = md.replace(was, need)
    return md


def _offline(model_version):
    raise RuntimeError(f"табло оффлайн: нет сырого zip для {model_version}")


def _require(path: Path) -> Path:
    if not path.is_file():
        raise FileNotFoundError(str(path))
    return path


def board_row(name: str, md: str, report: dict, route: str, golden: "str | None" = None) -> dict:
    tables = parse_pipe_tables(md)
    by_rule = Counter(item["rule"] for item in report["findings"])
    broken = sum(len(row) != len(table[0]) for table in tables for row in table)
    diffs = None if golden is None else count_diffs(md, golden)
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
        # ПРАВКА #83: нет файла эталона — все три None (поведение спеки 09)
        "count_diffs": None if golden is None else diffs,
        "threshold": None if golden is None else THRESHOLDS[name],
        "per_1000_tokens": None if golden is None else round(1000 * diffs / len(text_tokens(golden)), 1),
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
        golden = fixtures / GOLDEN[name]
        rows.append(board_row(name, md, report, detect_route(data, name),
                              golden.read_text(encoding="utf-8") if golden.is_file() else None))
    board = {"schema_version": BOARD_SCHEMA_VERSION,
             "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
             "rows": rows}
    return board, outputs


def _print_table(rows: list[dict]) -> None:
    lines = [tuple(CONSOLE_TITLES.get(col, col) for col in CONSOLE_COLUMNS)] + [tuple("-" if row[col] is None else str(row[col])
                                       for col in CONSOLE_COLUMNS) for row in rows]
    widths = [max(len(line[i]) for line in lines) for i in range(len(CONSOLE_COLUMNS))]
    for line in lines:
        print("  ".join(cell.ljust(width) for cell, width in zip(line, widths)).rstrip())


def write_goldens(outputs: dict, force: bool) -> None:
    """ПРАВКА #83: пять эталонов из черновиков и errors.txt; golden.md бейкоффа не трогается никогда."""
    targets = [(name, FIXTURES / golden) for name, golden in GOLDEN.items() if name not in NO_ERRORS_FILE]
    existing = [path.name for _, path in targets if path.exists()]
    if existing and not force:
        raise FileExistsError(f"эталон уже есть (нужен --force): {', '.join(existing)}")
    built = {}
    for name, path in targets:          # сначала собрать всё, потом писать: отказ не оставляет половину
        errors = _require(FIXTURES / f"{Path(name).stem}.errors.txt")   # пустой список подтверждается файлом
        try:        # ПРАВКА #84: ошибка называет файл и номер строки
            numbered = parse_errors_numbered(errors.read_text(encoding="utf-8"))
            built[path] = apply_errors(outputs[name][0], [error for _, error in numbered],
                                       lines=[number for number, _ in numbered])
        except ValueError as exc:
            raise ValueError(f"{errors.name}: {exc}") from exc
    for path, text in built.items():
        path.write_text(text, encoding="utf-8", newline="")
        print(path.name)


def main(argv: "list[str] | None" = None) -> int:
    try:
        parser = argparse.ArgumentParser(prog="python -m ocr.board")
        parser.add_argument("--golden", action="store_true", help="собрать <stem>.golden.md из черновиков и errors.txt")
        parser.add_argument("--force", action="store_true", help="перезаписать существующие эталоны")
        args = parser.parse_args(argv)
        # глобалы читаются в момент вызова: тесты подменяют пути на tmp_path
        with tempfile.TemporaryDirectory() as work:
            board, outputs = build_board(Path(work), fixtures=FIXTURES)
            if args.golden:
                write_goldens(outputs, args.force)
                board, outputs = build_board(Path(work) / "again", fixtures=FIXTURES)   # табло уже с эталонами
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
