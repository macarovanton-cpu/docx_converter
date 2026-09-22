"""ПРАВКА #87: измерение vision-сверки на остатке эталонов. В тракт не подключено.
ПРАВКА #88: модель переписывает полосы вырезки, вердикт — локальный diff по text_tokens.
ПРАВКА #89: бэкенд — Claude Code (claude -p); --verifier / --model / --fixture.
ПРАВКА #90: страница случая — физическая (*_model.json), окно через стык страниц — по двум полосам, narrow,
приклеенный маркер, промах кэша не останавливает замер."""

import argparse
import hashlib
import html
import json
import os
import re
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
from ocr.claude_code_verifier import CLAUDE_MODEL, ClaudeCodeTimeoutError, ClaudeCodeVerifier
from ocr.gemini_verifier import (CONTEXT_TOKENS, VerifierAuthError, VerifierConfigError, VerifierError,
                                 VerifierQuotaError, crop_block, locate_block)
from ocr.postprocess import HOMOGLYPHS_CYR_TO_LAT
from pdf_core import Verifier

MEASURE_SCHEMA_VERSION = 3                         # ПРАВКА #90
MEASURE_DIR = FIXTURES.parents[1]                  # _test/: verify_measure.<бэкенд>.<модель>.json, verify_review.….md
CROPS_DIR = MEASURE_DIR / "verify_crops"          # <stem>/pNN-MM.png — полоса MM страницы NN
ERROR_OUTCOMES = ("found", "neighbor", "not_found", "wrong_fix")
CONTROL_OUTCOMES = ("agree", "false_alarm", "unreadable")
SERVICE_OUTCOMES = ("unlocated", "no_image", "narrow", "error")     # ПРАВКА #90: + narrow
PAGE_SOURCES = ("content_list", "model_json", "seam")              # ПРАВКА #90: откуда страница у случая с полосами
ERROR_REASONS = ("miss", "stop", "timeout", "cli", "parse", "count", "other")   # ПРАВКА #90
ERROR_KINDS = ("merge", "homoglyph", "chars")
REVIEW_OUTCOMES = ("false_alarm", "wrong_fix", "neighbor")   # в отчёт для сверки человеком, в этом порядке
TILE_HEIGHT = 700            # PLACEHOLDER: высота полосы, px при CROP_RESOLUTION = 200 (≈ 8.9 см)
TILE_OVERLAP = 200           # PLACEHOLDER: перекрытие соседних полос, px (≈ 2.5 см, шесть строк 12 pt)
WINDOW_SLACK = 3             # окно транскрипции длиннее или короче фрагмента не больше чем на столько токенов
MIN_RATIO = 0.5              # PLACEHOLDER: SequenceMatcher.ratio лучшего окна ниже — место не найдено
MARKERS = frozenset("-–—•*·")   # токен только из этих знаков — маркер списка, в сравнении не участвует
GLUED = frozenset(";:,.)")   # ПРАВКА #90: знак перед приклеенным маркером списка: "DBM14G;–", "связи).-", "°C:–"
STOP_ERRORS = (VerifierQuotaError, VerifierAuthError, VerifierConfigError)
VERIFIERS = ("claude-code", "gemini")         # ПРАВКА #89
NO_RULE = "—"
_TAG_RE = re.compile(r"<[^>]+>")   # как в ocr/gemini_verifier.py: файл заморожен, приватное имя не импортируем

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
    tiles: tuple[bytes, ...]         # ПРАВКА #88: полосы вырезки блока (PNG); () -> unlocated / no_image / narrow
    next_tiles: tuple[bytes, ...] = ()    # ПРАВКА #90: окно через стык — полосы блока следующей страницы; иначе ()
    page_source: str | None = None        # ПРАВКА #90: PAGE_SOURCES, "narrow", "unresolved"; None — блока нет
    list_page: int | None = None          # ПРАВКА #90: страница блока content_list (в v2 она и была page), 1-based


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
    # ПРАВКА #90: маркер, приклеенный после GLUED ("DBM14G;–"), отделяется — и выбрасывается как маркер
    tokens = [token[:-1] if len(token) > 1 and token[-1] in MARKERS and token[-2] in GLUED else token
              for token in tokens]
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


def model_pages(zip_path: Path) -> list[str]:
    """ПРАВКА #90: текст таблиц каждой страницы до склейки (*_model.json; индекс = page_idx), без пробелов."""
    with zipfile.ZipFile(zip_path) as archive:
        names = [n for n in archive.namelist() if n.endswith("_model.json")]
        if len(names) != 1:
            raise ValueError(f"{zip_path.name}: _model.json — {names}")
        pages = json.loads(archive.read(names[0]).decode("utf-8"))
    # unescape обязателен: кавычки там лежат как &quot; (без него bakeoff2-c13 не находится)
    return [_squash(html.unescape(_TAG_RE.sub(" ", " ".join(block['content'] for block in page
                                                            if block['type'] == "table"))))
            for page in pages]


def table_run(content_list: list, index: int) -> list[int]:
    """ПРАВКА #90: индексы блоков склеенной таблицы — блок и его продолжения (табличные, стр. +1, пустой table_body)."""
    run, cur = [index], index
    for i in range(index + 1, len(content_list)):
        block = content_list[i]
        if block['type'] != "table":
            continue
        if (block['page_idx'] != content_list[cur]['page_idx'] + 1
                or _squash(_TAG_RE.sub(" ", block.get("table_body") or ""))):   # у пустого продолжения поля может не быть
            break
        run.append(i)
        cur = i
    return run


def physical_place(needle: list[str], before: list[str], after: list[str], content_list: list, index: int,
                   pages: list[str]) -> tuple[str, list[int]]:
    """ПРАВКА #90: (источник, блоки) — физическая страница окна; index — блок locate_block, pages — model_pages."""
    block = content_list[index]
    if block['type'] != "table":
        text = _squash(_TAG_RE.sub(" ", block.get("text") or block.get("table_body") or ""))
        return ("content_list" if _squash(" ".join(before + needle + after)) in text else "narrow"), [index]
    run = table_run(content_list, index)
    if len(run) == 1:       # склейки нет, а сырой текст *_model.json может расходиться с content_list
        return "content_list", [index]
    page = [pages[content_list[j]['page_idx']] for j in run]
    for k in range(CONTEXT_TOKENS, -1, -1):
        nd = _squash("".join((before[-k:] if k else []) + needle + after[:k]))
        if not nd:
            continue
        on_page = [j for j, text in zip(run, page) if nd in text]
        if on_page:
            return ("content_list" if on_page[0] == index else "model_json"), [on_page[0]]
        for n in range(len(run) - 1):       # на одном k сначала страница, потом стык
            if nd in page[n] + page[n + 1]:
                return "seam", [run[n], run[n + 1]]
    return "unresolved", []


def error_reason(exc: Exception) -> str:
    """ПРАВКА #90: причина сбоя страницы, одна из ERROR_REASONS; первое совпадение сверху."""
    text = str(exc)
    if text.startswith("нет в кэше"):
        return "miss"
    if isinstance(exc, STOP_ERRORS):
        return "stop"
    if isinstance(exc, ClaudeCodeTimeoutError):
        return "timeout"
    if "claude -p:" in text:
        return "cli"
    if "ответов " in text:
        return "count"
    if "ответ модели не JSON" in text or "ответ не по файлам" in text:
        return "parse"
    return "other"


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
        pdf = content_list = pages = None
        crops = {}
        if zip_name is not None:
            pdf = (fixtures / name).read_bytes()
            content_list = _content_list(fixtures / zip_name)
            pages = model_pages(fixtures / zip_name)       # ПРАВКА #90

        def to_b(i):
            """ПРАВКА #88: позиция в эталоне, соответствующая позиции i текста тракта."""
            if i == len(a):
                return len(b)
            if i == 0:          # вставка в начало документа (docx1-e01) — внутри окна, как вставка в конец
                return 0
            tag, i1, _, j1, _ = next(op for op in opcodes if op[1] <= i < op[2])
            return j1 + (i - i1) if tag == "equal" else j1

        def tiles_of(block):
            key = (block['page_idx'], tuple(block['bbox']))     # большинство опкодов сидит в одних и тех же таблицах
            if key not in crops:
                crops[key] = split_tiles(crop_block(pdf, block['page_idx'], block['bbox']))
            return crops[key]

        def place(needle, before, after):
            """ПРАВКА #90: (page, block_type, ambiguous, tiles, next_tiles, page_source, list_page) или пустышки.
            Страница — физическая: у продолжения склеенной таблицы bbox — блока-продолжения content_list."""
            if content_list is None:
                return None, None, False, (), (), None, None
            block, ambiguous = locate_block(needle, before, after, content_list)
            if block is None:
                return None, None, False, (), (), None, None
            index = next(i for i, b in enumerate(content_list) if b is block)
            source, blocks = physical_place(needle, before, after, content_list, index, pages)
            list_page = block['page_idx'] + 1
            if source == "narrow":
                return list_page, block['type'], ambiguous, (), (), source, list_page
            if source == "unresolved":
                return None, block['type'], ambiguous, (), (), source, list_page
            first = content_list[blocks[0]]
            next_tiles = tiles_of(content_list[blocks[1]]) if source == "seam" else ()
            return first['page_idx'] + 1, block['type'], ambiguous, tiles_of(first), next_tiles, source, list_page

        located = []
        for number, (tag, i1, i2, j1, j2) in enumerate(diffs, 1):
            lo, hi = max(0, i1 - CONTEXT_TOKENS), i2 + CONTEXT_TOKENS
            b_lo = to_b(lo)
            page, block_type, ambiguous, tiles, next_tiles, page_source, list_page = place(a[i1:i2], a[lo:i1], a[i2:hi])
            case = Case(id=f"{stem}-e{number:02d}", fixture=name, kind="error", tag=tag,
                        fragment=" ".join(a[lo:hi]), expected=" ".join(b[b_lo:to_b(min(hi, len(a)))]),
                        span=(i1 - lo, i2 - lo), gold_span=(j1 - b_lo, j2 - b_lo),
                        error_kind=error_kind(a[i1:i2], b[j1:j2]),
                        rules=_rules(a[i1:i2], report['findings']), page=page, block_type=block_type,
                        ambiguous=ambiguous, tiles=tiles, next_tiles=next_tiles, page_source=page_source,
                        list_page=list_page)
            cases.append(case)
            if page_source is not None:     # ПРАВКА #90: есть блок, даже без полос — окна контроля остаются окнами v2
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
                page, block_type, ambiguous, tiles, next_tiles, page_source, list_page = place(window, [], [])
                if not tiles:           # ПРАВКА #90: и unresolved — следующий кандидат
                    continue
                used.update(range(start, start + length))
                text = " ".join(window)
                cases.append(Case(id=f"{stem}-c{error.id[-2:]}", fixture=name, kind="control", tag=None,
                                  fragment=text, expected=text, span=None, gold_span=None, error_kind=None,
                                  rules=(), page=page, block_type=block_type, ambiguous=ambiguous, tiles=tiles,
                                  next_tiles=next_tiles, page_source=page_source, list_page=list_page))
                break
    return cases


def page_tiles(cases: list[Case]) -> dict[tuple[str, int], list[bytes]]:
    pages = {}
    for case in cases:
        if case.tiles:
            own = pages.setdefault((case.fixture, case.page), [])
            own.extend(tile for tile in case.tiles if tile not in own)
            if case.next_tiles:     # ПРАВКА #90: стык — полосы следующей страницы идут в её вызов
                own = pages.setdefault((case.fixture, case.page + 1), [])
                own.extend(tile for tile in case.next_tiles if tile not in own)
    return pages


def _tile_name(page: int, number: int) -> str:
    return f"p{page:02d}-{number:02d}.png"


def _count(cases: list[dict], **where) -> int:
    return sum(all(case[key] == value for key, value in where.items()) for case in cases)


def run_measure(cases: list[Case], verifier: Verifier) -> dict:
    pages = page_tiles(cases)
    answers, failures, missing, stop = {}, {}, [], None
    for key, images in pages.items():           # одна страница — один вызов
        try:
            texts = verifier.transcribe(images, TRANSCRIBE_QUESTION)
            if len(texts) != len(images):
                raise VerifierError(f"ответов {len(texts)}, картинок {len(images)}")
        except VerifierError as exc:            # ПРАВКА #90: причина — в строку и в totals
            reason = error_reason(exc)
            failures[key] = (reason, f"{type(exc).__name__}: {exc}")
            if reason == "miss":                # промах кэша — не стоп: оплаченное дальше пересуживается бесплатно
                missing.append(key)
            elif reason == "stop":              # повтор доберёт остальное: отвеченное уже в кэше
                stop = key
                break
            continue
        answers[key] = (texts, list(getattr(verifier, "last_cache_hits", [False] * len(images))))

    rows = []
    for case in cases:
        row = {"id": case.id, "fixture": case.fixture, "kind": case.kind, "tag": case.tag,
               "error_kind": case.error_kind, "page": case.page, "block_type": case.block_type,
               "ambiguous": case.ambiguous, "rules": list(case.rules), "fragment": case.fragment,
               "expected": case.expected, "span": None if case.span is None else list(case.span),
               "crop": None, "crop_next": None, "verdict": None, "reading": None, "ratio": None, "ops": None,
               "clipped": None, "outcome": None, "cache_hit": False, "error": None, "error_reason": None,
               "page_source": case.page_source, "list_page": case.list_page}
        rows.append(row)
        key = (case.fixture, case.page)
        keys = [key, (case.fixture, case.page + 1)] if case.next_tiles else [key]     # ПРАВКА #90: стык — две страницы
        failed = next((k for k in keys if k in failures), None)
        if not case.tiles:
            row["outcome"] = ("no_image" if dict(BOARD)[case.fixture] is None
                              else "narrow" if case.page_source == "narrow" else "unlocated")
        elif failed is not None:
            row["error_reason"], row["error"] = failures[failed]
            if row["error_reason"] not in ("stop", "miss"):   # страница остановки или промаха: исхода нет
                row["outcome"] = "error"
        elif case.next_tiles and all(k in answers for k in keys):     # ПРАВКА #90: низ страницы + верх следующей
            (texts, hits), (texts_next, hits_next) = answers[key], answers[keys[1]]
            last, first = pages[key].index(case.tiles[-1]), pages[keys[1]].index(case.next_tiles[0])
            row.update(judge(case, texts[last] + "\n" + texts_next[first]),
                       crop=f"{CROPS_DIR.name}/{Path(case.fixture).stem}/{_tile_name(case.page, last + 1)}",
                       crop_next=f"{CROPS_DIR.name}/{Path(case.fixture).stem}/{_tile_name(case.page + 1, first + 1)}",
                       cache_hit=bool(hits[last] and hits_next[first]))
        elif not case.next_tiles and key in answers:
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
    # привязан к блоку <=> есть страница; narrow — блок есть, а полос нет (ПРАВКА #90)
    measured = [row for row in errors if row["page"] is not None and row["outcome"] != "narrow"]
    totals = {"opcodes": len(errors), "measured": len(measured),
              "no_image": _count(errors, outcome="no_image"), "unlocated": _count(errors, outcome="unlocated"),
              "narrow": _count(errors, outcome="narrow"), "control": len(controls)}
    totals.update({outcome: _count(rows, outcome=outcome)
                   for outcome in ERROR_OUTCOMES + CONTROL_OUTCOMES + ("error",)})
    totals["page_sources"] = {source: _count(rows, page_source=source) for source in PAGE_SOURCES}     # ПРАВКА #90
    totals["error_reasons"] = {reason: _count(rows, error_reason=reason) for reason in ERROR_REASONS}

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
                           "narrow": _count(own_errors, outcome="narrow"),        # ПРАВКА #90
                           "no_image": _count(own_errors, outcome="no_image"),
                           "error": _count(own, outcome="error"),
                           "control": _count(own, kind="control"),
                           "false_alarm": _count(own, outcome="false_alarm"),
                           "unreadable": _count(own, outcome="unreadable"),
                           "moved": sum(row["page_source"] in ("model_json", "seam") for row in own)})
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
            "complete": stop is None and not missing,
            "missing_pages": [list(key) for key in missing],     # ПРАВКА #90: точный план живого добора
            "totals": totals, "by_fixture": by_fixture, "by_rule": by_rule,
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
            page = f"{row['page']}|{row['page'] + 1}" if row["crop_next"] else row["page"]    # ПРАВКА #90: стык
            lines += [f"### {row['id']} · {outcome} · стр. {page} · {row['block_type']}", f"![]({row['crop']})"]
            if row["crop_next"]:
                lines.append(f"![]({row['crop_next']})")
            lines.append(f"- фрагмент: `{row['fragment']}`")
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
        suffix = f"v{MEASURE_SCHEMA_VERSION}.{measure['verifier']}.{measure['model']}"   # ПРАВКА #90: v2 — улики, не трогать
        if args.fixture:
            suffix += "." + "+".join(Path(name).stem for name in board_names if name in args.fixture)
        MEASURE_DIR.mkdir(parents=True, exist_ok=True)
        (MEASURE_DIR / f"verify_measure.{suffix}.json").write_text(
            json.dumps(measure, ensure_ascii=False, indent=2), encoding="utf-8")
        write_review(measure, MEASURE_DIR / f"verify_review.{suffix}.md")
        if not measure["complete"]:
            stopped = next((row for row in measure["cases"] if row["error_reason"] == "stop"), None)
            if stopped is not None:
                print(f"замер остановлен на {stopped['id']}: {stopped['error']}", file=sys.stderr)
            if missing := measure["missing_pages"]:     # ПРАВКА #90: промах кэша — не стоп, а план добора
                error = next((row["error"] for row in measure["cases"] if row["error_reason"] == "miss"),
                             "нет в кэше, вызов выключен: нужен CLAUDE_CODE_LIVE=1")
                pages = ", ".join(f"{Path(fixture).stem} p{page:02d}" for fixture, page in missing)
                print(f"нет в кэше {len(missing)} стр.: {pages} — {error}", file=sys.stderr)
            return 1
        return 0
    except Exception:
        traceback.print_exc(file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
