"""ПРАВКА #87: измерение vision-сверки на остатке эталонов. В тракт не подключено.
ПРАВКА #88: модель переписывает полосы вырезки, вердикт — локальный diff по text_tokens.
ПРАВКА #89: бэкенд — Claude Code (claude -p); --verifier / --model / --fixture."""

import argparse
import hashlib
import json
import os
import shutil
import sys
import tempfile
import traceback
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path

import PIL.Image

from ocr.board import BOARD, FIXTURES, GOLDEN, build_board, count_diffs, text_tokens
from ocr.claude_code_verifier import CLAUDE_MODEL, ClaudeCodeVerifier
from ocr.gemini_verifier import (CONTEXT_TOKENS, VerifierAuthError, VerifierConfigError, VerifierError,
                                 VerifierQuotaError, crop_block, locate_block)
from ocr.postprocess import HOMOGLYPHS_CYR_TO_LAT
from pdf_core import Verifier

MEASURE_SCHEMA_VERSION = 2
MEASURE_DIR = FIXTURES.parents[1]                  # _test/: verify_measure.<бэкенд>.<модель>.json, verify_review.….md
CROPS_DIR = MEASURE_DIR / "verify_crops"          # <stem>/pNN-MM.png — полоса MM страницы NN
ERROR_OUTCOMES = ("found", "neighbor", "not_found", "wrong_fix")
CONTROL_OUTCOMES = ("agree", "false_alarm", "unreadable")
SERVICE_OUTCOMES = ("unlocated", "no_image", "error")
ERROR_KINDS = ("merge", "homoglyph", "chars")
REVIEW_OUTCOMES = ("false_alarm", "wrong_fix", "neighbor")   # в отчёт для сверки человеком, в этом порядке
TILE_HEIGHT = 700            # PLACEHOLDER: высота полосы, px при CROP_RESOLUTION = 200 (≈ 8.9 см)
TILE_OVERLAP = 200           # PLACEHOLDER: перекрытие соседних полос, px (≈ 2.5 см, шесть строк 12 pt)
WINDOW_SLACK = 3             # окно транскрипции длиннее или короче фрагмента не больше чем на столько токенов
MIN_RATIO = 0.5              # PLACEHOLDER: SequenceMatcher.ratio лучшего окна ниже — место не найдено
MARKERS = frozenset("-–—•*·")   # токен только из этих знаков — маркер списка, в сравнении не участвует
STOP_ERRORS = (VerifierQuotaError, VerifierAuthError, VerifierConfigError)
VERIFIERS = ("claude-code", "gemini")         # ПРАВКА #89
NO_RULE = "—"

TRANSCRIBE_QUESTION = (
    "Тебе даны картинки — вырезки одной страницы документа на русском языке (скан или печать).\n"
    "Перепиши текст каждой картинки дословно, как он напечатан: буквы, цифры, знаки, латиница и\n"
    "кириллица, пробелы между словами. Ничего не исправляй, не дополняй и не перефразируй: опечатку\n"
    "на картинке переписывай как опечатку.\n"
    "Каждую строку текста начинай с новой строки. Слово, перенесённое со знаком переноса в конце\n"
    "строки, пиши целиком без знака переноса; дефис внутри слова (технико-экономический) сохраняй.\n"
    "Строки, обрезанные верхним или нижним краем картинки, пропускай.\n"
    "Таблицу переписывай ячейка за ячейкой: строки таблицы сверху вниз, ячейки строки слева направо,\n"
    "строки текста внутри ячейки — по порядку, каждая с новой строки. Без разметки таблиц.\n"
    "Ответь одним JSON-объектом без пояснений: ключ — имя файла картинки, значение — её текст;\n"
    "картинка без читаемого текста — пустая строка:\n"
    '{"01.png": "<текст>", "02.png": "<текст>"}'
)

REVIEW_PROCEDURE = (
    "> Каждый случай сверить с картинкой. Модель права — это пропущенная ошибка тракта: строка в `<stem>.errors.txt`,\n"
    "> потом `--golden --force`, новые sha и порог — отдельной правкой; у `bakeoff.pdf` эталон `golden.md` не трогается\n"
    "> никогда — только запись в docs_sync. Модель ошиблась — случай остаётся ложной тревогой. Итог — в docs_sync: сколько\n"
    "> `false_alarm` подтвердилось как ложные."
)

_TO_CYRILLIC = str.maketrans({lat: cyr for cyr, lat in HOMOGLYPHS_CYR_TO_LAT.items()})


@dataclass(frozen=True)
class Case:
    id: str                          # "<stem>-e07" / "<stem>-c07"
    fixture: str
    kind: str                        # "error" | "control"
    tag: str | None
    fragment: str                    # окно текста тракта, токены через пробел
    expected: str                    # ПРАВКА #88: то же окно в эталоне, исправлены все опкоды окна; у control == fragment
    span: tuple[int, int] | None     # ПРАВКА #88: опкод — токены fragment.split()[s:e]; у control — None
    gold_span: tuple[int, int] | None   # ПРАВКА #88: его эталонная сторона — expected.split()[s:e]
    error_kind: str | None           # ПРАВКА #88: одно из ERROR_KINDS; у control — None
    rules: tuple[str, ...]
    page: int | None                 # 1-based
    block_type: str | None
    ambiguous: bool
    tiles: tuple[bytes, ...]         # ПРАВКА #88: полосы вырезки блока (PNG); () -> unlocated / no_image


def split_tiles(png: bytes, *, height: int = TILE_HEIGHT, overlap: int = TILE_OVERLAP) -> tuple[bytes, ...]:
    """Горизонтальные полосы с перекрытием, оттенки серого — ради размера (Read в Claude Code, порог 500 КБ)."""
    image = PIL.Image.open(BytesIO(png)).convert("L")
    w, h = image.size
    starts = [0] if h <= height else list(range(0, h - height, height - overlap)) + [h - height]
    tiles = []
    for s in starts:
        buffer = BytesIO()
        (image if h <= height else image.crop((0, s, w, s + height))).save(buffer, "PNG", optimize=True)
        tiles.append(buffer.getvalue())
    return tuple(tiles)


def error_kind(a_side: list[str], b_side: list[str]) -> str:
    a, b = "".join(a_side), "".join(b_side)
    if a == b:
        return "merge"
    if a.translate(_TO_CYRILLIC) == b.translate(_TO_CYRILLIC):
        return "homoglyph"
    return "chars"


def norm(tokens: list[str]) -> list[str]:
    return [token for token in tokens if not set(token) <= MARKERS]


def best_window(fragment: list[str], transcript: list[str]) -> tuple[int, int, float]:
    if not transcript:
        return 0, 0, 0.0
    n = len(fragment)
    matcher = SequenceMatcher(None, autojunk=False)
    matcher.set_seq2(fragment)          # фрагмент задаётся один раз: b2j строится по seq2
    best = (0, 0, -1.0)
    for length in range(max(1, n - WINDOW_SLACK), n + WINDOW_SLACK + 1):
        for start in range(0, max(1, len(transcript) - length + 1)):
            matcher.set_seq1(transcript[start:start + length])
            ratio = matcher.ratio()
            if ratio > best[2]:
                best = (start, start + length, ratio)
    return best


def _split(tokens: list[str], span: tuple[int, int]) -> tuple[list[str], tuple[int, int]]:
    """Токены без маркеров и пролёт опкода в них."""
    before, mid, after = norm(tokens[:span[0]]), norm(tokens[span[0]:span[1]]), norm(tokens[span[1]:])
    return before + mid + after, (len(before), len(before) + len(mid))


def _ops(a: list[str], b: list[str]) -> tuple[list[tuple], list[tuple]]:
    """(живые, срезанные краем) не-equal опкоды a↔b."""
    ops = [op for op in SequenceMatcher(None, a, b, autojunk=False).get_opcodes() if op[0] != "equal"]
    clipped = []
    if ops and ops[0][0] == "delete" and ops[0][1] == 0:
        clipped.append(ops.pop(0))
    if ops and ops[-1][0] == "delete" and ops[-1][2] == len(a):
        clipped.append(ops.pop())
    return ops, clipped


def _touches(i1: int, i2: int, s: int, e: int) -> bool:
    if i1 == i2 or s == e:
        return s <= i2 and i1 <= e
    return i1 < e and s < i2


def judge(case: Case, transcript: str) -> dict:
    if case.kind == "error":
        f, span = _split(case.fragment.split(), case.span)
        expected, gold_span = _split(case.expected.split(), case.gold_span)
    else:
        f = norm(case.fragment.split())
    t = norm(text_tokens(transcript))
    i, j, ratio = best_window(f, t)
    window = t[i:j]
    live, clipped = _ops(f, window)
    result = {"verdict": "fix" if live else "agree", "outcome": None, "reading": " ".join(window),
              "ratio": ratio, "clipped": sum(i2 - i1 for _, i1, i2, _, _ in clipped),
              "ops": [[tag, " ".join(f[i1:i2]), " ".join(window[j1:j2]), error_kind(f[i1:i2], window[j1:j2])]
                      for tag, i1, i2, j1, j2 in live]}
    if ratio < MIN_RATIO:
        result.update(verdict="unreadable", outcome="not_found" if case.kind == "error" else "unreadable")
    elif case.kind == "control":
        result["outcome"] = "false_alarm" if live else "agree"
    elif any(_touches(i1, i2, *span) for _, i1, i2, _, _ in clipped):
        result.update(outcome="not_found", reading="clipped")
    elif not live:
        result["outcome"] = "not_found"
    elif not any(_touches(i1, i2, *span) for _, i1, i2, _, _ in live):
        result["outcome"] = "neighbor"
    else:
        gold_live, _ = _ops(expected, window)
        result["outcome"] = ("wrong_fix" if any(_touches(i1, i2, *gold_span) for _, i1, i2, _, _ in gold_live)
                             else "found")
    return result


def _content_list(zip_path: Path) -> list:
    with zipfile.ZipFile(zip_path) as archive:
        names = [n for n in archive.namelist() if n.endswith("content_list.json")]   # _v2 под суффикс не подходит
        if len(names) != 1:
            raise ValueError(f"{zip_path.name}: content_list.json — {names}")
        return json.loads(archive.read(names[0]).decode("utf-8"))


def _squash(text: str) -> str:
    return "".join(text.split())


def _rules(a_side: list[str], findings: list[dict]) -> tuple[str, ...]:
    needle = _squash("".join(a_side))
    if not needle:
        return ()
    hit = set()
    for finding in findings:
        snippet = _squash(finding['snippet'])
        if snippet and (snippet in needle or needle in snippet):
            hit.add(finding['rule'])
    return tuple(sorted(hit))


def build_cases(outputs: dict, *, fixtures: Path = FIXTURES) -> list[Case]:
    cases = []
    for name, zip_name in BOARD:
        md, report = outputs[name]
        golden = (fixtures / GOLDEN[name]).read_text(encoding="utf-8")
        a, b = text_tokens(md), text_tokens(golden)
        opcodes = SequenceMatcher(None, a, b, autojunk=False).get_opcodes()
        diffs = [op for op in opcodes if op[0] != "equal"]
        assert len(diffs) == count_diffs(md, golden), (name, len(diffs))    # замер и табло считают одно и то же
        stem = Path(name).stem
        pdf = content_list = None
        crops = {}
        if zip_name is not None:
            pdf = (fixtures / name).read_bytes()
            content_list = _content_list(fixtures / zip_name)

        def to_b(i):
            """ПРАВКА #88: позиция в эталоне, соответствующая позиции i текста тракта."""
            if i == len(a):
                return len(b)
            if i == 0:          # вставка в начало документа (docx1-e01) — внутри окна, как вставка в конец
                return 0
            tag, i1, _, j1, _ = next(op for op in opcodes if op[1] <= i < op[2])
            return j1 + (i - i1) if tag == "equal" else j1

        def place(needle, before, after):
            """(page, block_type, ambiguous, tiles) или четыре пустышки."""
            if content_list is None:
                return None, None, False, ()
            block, ambiguous = locate_block(needle, before, after, content_list)
            if block is None:
                return None, None, False, ()
            key = (block['page_idx'], tuple(block['bbox']))     # большинство опкодов сидит в одних и тех же таблицах
            if key not in crops:
                crops[key] = split_tiles(crop_block(pdf, block['page_idx'], block['bbox']))
            return block['page_idx'] + 1, block['type'], ambiguous, crops[key]

        located = []
        for number, (tag, i1, i2, j1, j2) in enumerate(diffs, 1):
            lo, hi = max(0, i1 - CONTEXT_TOKENS), i2 + CONTEXT_TOKENS
            b_lo = to_b(lo)
            page, block_type, ambiguous, tiles = place(a[i1:i2], a[lo:i1], a[i2:hi])
            case = Case(id=f"{stem}-e{number:02d}", fixture=name, kind="error", tag=tag,
                        fragment=" ".join(a[lo:hi]), expected=" ".join(b[b_lo:to_b(min(hi, len(a)))]),
                        span=(i1 - lo, i2 - lo), gold_span=(j1 - b_lo, j2 - b_lo),
                        error_kind=error_kind(a[i1:i2], b[j1:j2]),
                        rules=_rules(a[i1:i2], report['findings']), page=page, block_type=block_type,
                        ambiguous=ambiguous, tiles=tiles)
            cases.append(case)
            if tiles:
                located.append(case)

        # контроль: окна внутри equal-участков, не ближе CONTEXT_TOKENS к любому опкоду; шаг равномерный, без random
        used = set()
        for index, error in enumerate(located):
            length = len(error.fragment.split())
            starts = [s for tag, i1, i2, _, _ in opcodes if tag == "equal"
                      for s in range(i1 + (CONTEXT_TOKENS if i1 else 0),
                                     i2 - (CONTEXT_TOKENS if i2 < len(a) else 0) - length + 1)]
            if not starts:
                continue
            step = len(starts) / len(located)
            for start in starts[int(index * step + step / 2):]:     # не привязался — следующий кандидат
                if used & set(range(start, start + length)):
                    continue
                window = a[start:start + length]
                page, block_type, ambiguous, tiles = place(window, [], [])
                if not tiles:
                    continue
                used.update(range(start, start + length))
                text = " ".join(window)
                cases.append(Case(id=f"{stem}-c{error.id[-2:]}", fixture=name, kind="control", tag=None,
                                  fragment=text, expected=text, span=None, gold_span=None, error_kind=None,
                                  rules=(), page=page, block_type=block_type, ambiguous=ambiguous, tiles=tiles))
                break
    return cases


def page_tiles(cases: list[Case]) -> dict[tuple[str, int], list[bytes]]:
    pages = {}
    for case in cases:
        if case.tiles:
            own = pages.setdefault((case.fixture, case.page), [])
            own.extend(tile for tile in case.tiles if tile not in own)
    return pages


def _tile_name(page: int, number: int) -> str:
    return f"p{page:02d}-{number:02d}.png"


def _count(cases: list[dict], **where) -> int:
    return sum(all(case[key] == value for key, value in where.items()) for case in cases)


def run_measure(cases: list[Case], verifier: Verifier) -> dict:
    pages = page_tiles(cases)
    answers, failures, stop = {}, {}, None
    for key, images in pages.items():           # одна страница — один вызов
        try:
            texts = verifier.transcribe(images, TRANSCRIBE_QUESTION)
            if len(texts) != len(images):
                raise VerifierError(f"ответов {len(texts)}, картинок {len(images)}")
        except STOP_ERRORS as exc:              # повтор доберёт остальное: отвеченное уже в кэше
            failures[key] = f"{type(exc).__name__}: {exc}"
            stop = key
            break
        except VerifierError as exc:
            failures[key] = f"{type(exc).__name__}: {exc}"
            continue
        answers[key] = (texts, list(getattr(verifier, "last_cache_hits", [False] * len(images))))

    rows = []
    for case in cases:
        row = {"id": case.id, "fixture": case.fixture, "kind": case.kind, "tag": case.tag,
               "error_kind": case.error_kind, "page": case.page, "block_type": case.block_type,
               "ambiguous": case.ambiguous, "rules": list(case.rules), "fragment": case.fragment,
               "expected": case.expected, "span": None if case.span is None else list(case.span),
               "crop": None, "verdict": None, "reading": None, "ratio": None, "ops": None, "clipped": None,
               "outcome": None, "cache_hit": False, "error": None}
        rows.append(row)
        key = (case.fixture, case.page)
        if not case.tiles:
            row["outcome"] = "no_image" if dict(BOARD)[case.fixture] is None else "unlocated"
        elif key in failures:
            row["error"] = failures[key]
            if key != stop:                                # страница остановки: исхода нет, только текст ошибки
                row["outcome"] = "error"
        elif key in answers:
            texts, hits = answers[key]
            best = None
            for tile in case.tiles:                        # своя полоса с наибольшим ratio, при равенстве — первая
                number = pages[key].index(tile)
                verdict = judge(case, texts[number])
                if best is None or verdict["ratio"] > best[1]["ratio"]:
                    best = (number, verdict)
            number, verdict = best
            row.update(verdict, crop=f"{CROPS_DIR.name}/{Path(case.fixture).stem}/{_tile_name(case.page, number + 1)}", cache_hit=bool(hits[number]))

    errors = [row for row in rows if row["kind"] == "error"]
    controls = [row for row in rows if row["kind"] == "control"]
    measured = [row for row in errors if row["page"] is not None]      # привязан к блоку <=> есть страница
    totals = {"opcodes": len(errors), "measured": len(measured),
              "no_image": _count(errors, outcome="no_image"), "unlocated": _count(errors, outcome="unlocated"),
              "control": len(controls)}
    totals.update({outcome: _count(rows, outcome=outcome)
                   for outcome in ERROR_OUTCOMES + CONTROL_OUTCOMES + ("error",)})

    def error_columns(subset):
        return {outcome: _count(subset, outcome=outcome) for outcome in ERROR_OUTCOMES}

    by_fixture = []
    for name, _ in BOARD:
        own = [row for row in rows if row["fixture"] == name]
        if not own:
            continue
        own_errors = [row for row in own if row["kind"] == "error"]
        by_fixture.append({"fixture": name, "opcodes": len(own_errors), **error_columns(own_errors),
                           "unlocated": _count(own_errors, outcome="unlocated"),
                           "no_image": _count(own_errors, outcome="no_image"),
                           "error": _count(own, outcome="error"),
                           "control": _count(own, kind="control"),
                           "false_alarm": _count(own, outcome="false_alarm"),
                           "unreadable": _count(own, outcome="unreadable")})
    rules = sorted({rule for row in measured for rule in row["rules"] or [NO_RULE]})
    by_rule = []
    for rule in rules:
        own = [row for row in measured if rule in (row["rules"] or [NO_RULE])]
        by_rule.append({"rule": rule, "opcodes": len(own), **error_columns(own)})
    by_block_type = []
    for block_type in sorted({row["block_type"] for row in measured + controls}):
        own = [row for row in measured if row["block_type"] == block_type]
        own_controls = [row for row in controls if row["block_type"] == block_type]
        by_block_type.append({"block_type": block_type, "opcodes": len(own), **error_columns(own),
                              "control": len(own_controls),
                              "false_alarm": _count(own_controls, outcome="false_alarm")})
    by_kind = []
    for kind in ERROR_KINDS:
        own = [row for row in measured if row["error_kind"] == kind]
        by_kind.append({"error_kind": kind, "opcodes": len(own), **error_columns(own)})
    return {"schema_version": MEASURE_SCHEMA_VERSION,
            "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "verifier": getattr(verifier, "name", type(verifier).__name__),
            "model": getattr(verifier, "model", None),
            "question_sha256": hashlib.sha256(TRANSCRIBE_QUESTION.encode()).hexdigest(),
            "complete": stop is None, "totals": totals, "by_fixture": by_fixture, "by_rule": by_rule,
            "by_block_type": by_block_type, "by_kind": by_kind,
            "calls": list(getattr(verifier, "calls_log", [])), "cases": rows}


def write_review(measure: dict, path: Path) -> None:
    """Markdown для сверки человеком: каждый случай из REVIEW_OUTCOMES — с полосой, фрагментом и прочитанным."""
    lines = [f"# Сверка vision-замера: {measure['verifier']} · {measure['model']} · {measure['created_at']}", "",
             "totals: " + json.dumps(measure["totals"], ensure_ascii=False), "", REVIEW_PROCEDURE, ""]
    for outcome in REVIEW_OUTCOMES:
        own = [row for row in measure["cases"] if row["outcome"] == outcome]
        lines += [f"## {outcome} — {len(own)}", ""]
        for row in own:
            lines += [f"### {row['id']} · {outcome} · стр. {row['page']} · {row['block_type']}",
                      f"![]({row['crop']})", f"- фрагмент: `{row['fragment']}`"]
            if row["kind"] == "error":
                lines.append(f"- эталон: `{row['expected']}`")
            ops = "; ".join(f"`{was or '∅'}` → `{read or '∅'}` ({kind})" for _, was, read, kind in row["ops"])
            lines += [f"- прочитано: `{row['reading']}`", f"- расхождения: {ops}",
                      "- [ ] модель права — ошибка в эталоне   [ ] модель ошиблась", ""]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def make_verifier(name: str, model: str | None) -> Verifier:
    """ПРАВКА #89: бэкенд транскрипции по имени; Gemini транскрипции не умеет — решение человека 21.09.2026."""
    if name == "gemini":
        raise VerifierConfigError("Gemini отвечает только на вопрос #87 (вердикт по фрагменту), транскрипции у него "
                                  "нет — в замер не подключён (решение человека 21.09.2026)")
    return ClaudeCodeVerifier(model=model or CLAUDE_MODEL)


def _print_table(rows: list[dict]) -> None:
    if not rows:
        return
    lines = [tuple(rows[0])] + [tuple(str(value) for value in row.values()) for row in rows]
    widths = [max(len(line[i]) for line in lines) for i in range(len(lines[0]))]
    for line in lines:
        print("  ".join(cell.ljust(width) for cell, width in zip(line, widths)).rstrip())
    print()


def main(argv: "list[str] | None" = None) -> int:
    parser = argparse.ArgumentParser(prog="python -X utf8 -m ocr.measure",
                                     description="Замер vision-сверки на остатке эталонов (в тракт не подключено)")
    parser.add_argument("--verifier", default=os.environ.get("VERIFIER", "claude-code"), help=" | ".join(VERIFIERS))
    parser.add_argument("--model", default=None, help="по умолчанию — модель бэкенда")
    parser.add_argument("--fixture", action="append", default=[], help="имя из BOARD; повторяемый")
    args = parser.parse_args(argv)
    board_names = [name for name, _ in BOARD]
    if args.verifier not in VERIFIERS:
        print(f"неизвестный бэкенд {args.verifier!r}: {', '.join(VERIFIERS)}", file=sys.stderr)
        return 1
    if unknown := [name for name in args.fixture if name not in board_names]:
        print(f"нет в BOARD: {unknown}; есть {board_names}", file=sys.stderr)
        return 1
    try:
        # глобалы читаются в момент вызова: тесты подменяют пути и make_verifier
        with tempfile.TemporaryDirectory() as work:
            _, outputs = build_board(Path(work), fixtures=FIXTURES)
        cases = build_cases(outputs, fixtures=FIXTURES)
        if args.fixture:
            cases = [case for case in cases if case.fixture in args.fixture]
        pages = page_tiles(cases)
        for stem in dict.fromkeys(Path(fixture).stem for fixture, _ in pages):   # полосы — до обращения к модели
            shutil.rmtree(CROPS_DIR / stem, ignore_errors=True)
        for (fixture, page), tiles in pages.items():
            for number, tile in enumerate(tiles, 1):
                path = CROPS_DIR / Path(fixture).stem / _tile_name(page, number)
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(tile)
        verifier = make_verifier(args.verifier, args.model)
        measure = run_measure(cases, verifier)
        for table in ("by_fixture", "by_rule", "by_block_type", "by_kind"):
            _print_table(measure[table])
        print("by_rule: опкод с несколькими правилами идёт в строку каждого — сумма строк может превышать measured")
        print("totals:", json.dumps(measure["totals"], ensure_ascii=False))
        calls = measure["calls"]
        print(f"вызовов {len(calls)}, картинок {sum(c['images'] for c in calls)}, входных токенов на вызов в среднем "
              f"{round(sum(c['input_tokens'] for c in calls) / len(calls)) if calls else 0}, "
              f"выходных всего {sum(c['output_tokens'] for c in calls)}")
        suffix = f"{measure['verifier']}.{measure['model']}"
        if args.fixture:
            suffix += "." + "+".join(Path(name).stem for name in board_names if name in args.fixture)
        MEASURE_DIR.mkdir(parents=True, exist_ok=True)
        (MEASURE_DIR / f"verify_measure.{suffix}.json").write_text(
            json.dumps(measure, ensure_ascii=False, indent=2), encoding="utf-8")
        write_review(measure, MEASURE_DIR / f"verify_review.{suffix}.md")
        if not measure["complete"]:
            stopped = next(row for row in measure["cases"] if row["outcome"] is None and row["error"])
            print(f"замер остановлен на {stopped['id']}: {stopped['error']}", file=sys.stderr)
            return 1
        return 0
    except Exception:
        traceback.print_exc(file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
