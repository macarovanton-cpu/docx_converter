"""ПРАВКА #91: сверка по картинке в тракте. Весь документ режется на окна, каждое окно — полосы своего блока
на физической странице; модель (claude -p) переписывает страницу, разница — локальный diff по токенам тем же
judge, что в замере (ocr.measure не меняется). Подсказка текст не меняет: только находка vision_diff.
Только локально: без claude, входа или лимита — находка vision_skipped, конвертация идёт дальше."""

import json
import re
import sys
import traceback
import zipfile
from collections import Counter
from dataclasses import dataclass
from difflib import SequenceMatcher
from io import BytesIO

from ocr import Finding
from ocr.board import text_tokens
from ocr.claude_code_verifier import ClaudeCodeVerifier
from ocr.gemini_verifier import CONTEXT_TOKENS, crop_block, locate_block
from ocr.measure import (TRANSCRIBE_QUESTION, Case, model_pages, norm, page_tiles, physical_place, run_measure,
                         split_tiles)
from ocr.postprocess import HOMOGLYPHS_CYR_TO_LAT
from pdf_core import Verifier

VISION_MODEL = "claude-opus-5"   # решение человека 22.09.2026: Opus по умолчанию, Sonnet — параметром
VISION_STRIDE = 3                # PLACEHOLDER: токенов в пролёте окна; окно — пролёт ± CONTEXT_TOKENS
PAGE_HINT_LIMIT = 15             # решение человека 23.09.2026: больше на странице — сбой нарезки или скана
SEVERITY = {"vision_diff": "warning", "vision_skipped": "warning"}
_MARKUP_RE = re.compile(r"#{1,6}|!?\[[^\]]*\]\([^)]*\)")    # заголовок, картинка, ссылка: на скане их нет
_TO_CYRILLIC = str.maketrans({lat: cyr for cyr, lat in HOMOGLYPHS_CYR_TO_LAT.items()})   # как в ocr/measure.py

@dataclass(frozen=True)
class Hint:
    start: int           # пролёт «нашей» стороны в plain-токенах, [start, end); у вставки — соседний токен
    end: int
    suggestion: str      # прочитанное по скану, токены через пробел; "" — у нас лишнее
    page: int            # физическая страница окна (у стыка — первая из двух)
    reading: str         # окно прочтения, по которому судилось окно


def make_verifier(model: str) -> Verifier:
    """live=True: флаг --vision (в UI — галочка) и есть согласие на расход подписки; кэш — умолчание."""
    return ClaudeCodeVerifier(model=model, live=True)


def plain_tokens(md: str) -> tuple[list[str], list[tuple[int, int]]]:
    """Токены окон и срез каждого в md: без #, картинок/ссылок и обратной косой (на скане их нет)."""
    tokens, spans, pos = [], [], 0
    for token in text_tokens(md):
        start = md.find(token, pos)
        if start == -1:
            raise ValueError(f"токен {token!r} не найден в markdown после позиции {pos}")
        pos = start + len(token)
        clean = token.replace("\\", "")
        if _MARKUP_RE.fullmatch(token) or not clean:
            continue
        tokens.append(clean)
        spans.append((start, pos))
    return tokens, spans


def _letters(tokens: list[str]) -> list[str]:
    return [t for t in (re.sub(r"[\W_]", "", x.translate(_TO_CYRILLIC)) for x in tokens) if t]


def is_cosmetic(ours: list[str], read: list[str]) -> bool:
    """Разница только в гомоглифах, пунктуации и тире; границы токенов остаются (решение 3)."""
    return _letters(ours) == _letters(read)


def _touches(i1: int, i2: int, s: int, e: int) -> bool:
    # копия ocr.measure._touches: файл заморожен, приватное имя не импортируем
    if i1 == i2 or s == e:
        return s <= i2 and i1 <= e
    return i1 < e and s < i2


def build_windows(tokens: list[str], pdf_bytes: bytes, content_list: list,
                  pages: list[str]) -> tuple[list[Case], Counter]:
    """Окна по всему документу: пролёт VISION_STRIDE ± CONTEXT_TOKENS; в cases — только окна с полосами."""
    crops, cases, skipped = {}, [], Counter(unlocated=0, unresolved=0)

    def tiles_of(block):
        key = (block['page_idx'], tuple(block['bbox']))     # как tiles_of в ocr.measure.build_cases
        if key not in crops:
            crops[key] = split_tiles(crop_block(pdf_bytes, block['page_idx'], block['bbox']))
        return crops[key]

    for i in range(0, len(tokens), VISION_STRIDE):
        e = min(len(tokens), i + VISION_STRIDE)
        lo, hi = max(0, i - CONTEXT_TOKENS), min(len(tokens), e + CONTEXT_TOKENS)
        needle, before, after = tokens[i:e], tokens[lo:i], tokens[e:hi]
        block, ambiguous = locate_block(needle, before, after, content_list)
        if block is None:
            skipped["unlocated"] += 1
            continue
        index = next(k for k, b in enumerate(content_list) if b is block)
        source, blocks = physical_place(needle, before, after, content_list, index, pages)
        if source == "unresolved":
            skipped["unresolved"] += 1
            continue
        if source == "narrow":      # в тракте судится по полосам своего блока (в замере — служебный исход)
            blocks = [index]
        first = content_list[blocks[0]]
        next_tiles = tiles_of(content_list[blocks[1]]) if source == "seam" else ()
        window = " ".join(tokens[lo:hi])
        cases.append(Case(id=f"w{i:05d}", fixture="", kind="error", tag=None, fragment=window, expected=window,
                          span=(i - lo, e - lo), gold_span=(i - lo, e - lo), error_kind=None, rules=(),
                          page=first['page_idx'] + 1, block_type=block['type'], ambiguous=ambiguous,
                          tiles=tiles_of(first), next_tiles=next_tiles, page_source=source,
                          list_page=block['page_idx'] + 1))
    return cases, skipped


def hints(rows: list[dict]) -> list[Hint]:
    """Подсказки из строк run_measure: опкоды «наше ↔ прочитанное», касающиеся пролёта, кроме косметических."""
    found = {}
    for row in rows:
        if row["outcome"] != "wrong_fix":
            continue
        raw = row["fragment"].split()
        s, e = row["span"]
        lo = int(row["id"][1:]) - s
        # как ocr.measure._split: norm поштучный, пролёт — в живых токенах
        kept = [k for k, t in enumerate(raw) if norm([t])]
        f = [norm([raw[k]])[0] for k in kept]
        fs, fe = sum(k < s for k in kept), sum(k < e for k in kept)
        w = row["reading"].split()
        for tag, i1, i2, j1, j2 in SequenceMatcher(None, f, w, autojunk=False).get_opcodes():
            if tag == "equal" or not _touches(i1, i2, fs, fe) or is_cosmetic(f[i1:i2], w[j1:j2]):
                continue
            if i2 > i1:
                start, end = lo + kept[i1], lo + kept[i2 - 1] + 1
            else:           # вставка: соседний токен слева, в начале окна — справа
                start = lo + kept[i1 - 1] if i1 else lo + kept[0]
                end = start + 1
            hint = Hint(start, end, " ".join(w[j1:j2]), row["page"], row["reading"])
            found.setdefault((start, end, hint.suggestion), hint)    # опкод на границе пролётов — в двух окнах
    return sorted(found.values(), key=lambda hint: (hint.start, hint.end))


def limit_pages(found: list[Hint]) -> tuple[list[Hint], dict[int, int]]:
    """Предохранитель: страница, где подсказок больше PAGE_HINT_LIMIT, выдаётся одной находкой, не подсказками."""
    counts = Counter(hint.page for hint in found)
    flooded = {page: n for page, n in counts.items() if n > PAGE_HINT_LIMIT}
    return [hint for hint in found if hint.page not in flooded], flooded


def skipped_finding(page: int | None, text: str, model: str) -> Finding:
    return Finding(rule="vision_skipped", severity=SEVERITY["vision_skipped"], page=page, snippet="",
                   suggestion=text, model=model)


class _Progress:
    """Обёртка верификатора для run_measure: прогресс перед каждой страницей, номер последней начатой."""

    def __init__(self, verifier, keys, progress):
        self._verifier, self._keys, self._progress = verifier, keys, progress
        self.name = getattr(verifier, "name", type(verifier).__name__)
        self.model = getattr(verifier, "model", None)
        self.done = 0               # сколько страниц начато; при стопе keys[done - 1] — страница остановки
        self.last_cache_hits: list[bool] = []

    def transcribe(self, images, question):
        page = self._keys[self.done][1]
        self.done += 1
        if self._progress is not None:
            self._progress(page, self.done, len(self._keys))
        texts = self._verifier.transcribe(images, question)
        self.last_cache_hits = list(getattr(self._verifier, "last_cache_hits", [False] * len(images)))
        return texts


def _content_list(zip_bytes: bytes, source_name: str) -> list:
    # копия ocr.measure._content_list, но по байтам: файл заморожен, приватное имя не импортируем
    with zipfile.ZipFile(BytesIO(zip_bytes)) as archive:
        names = [n for n in archive.namelist() if n.endswith("content_list.json")]   # _v2 под суффикс не подходит
        if len(names) != 1:
            raise ValueError(f"{source_name}: zip MinerU: content_list.json — {names}")
        return json.loads(archive.read(names[0]).decode("utf-8"))


def vision_findings(md: str, pdf_bytes: bytes, zip_bytes: bytes, *, source_name: str, model: str = VISION_MODEL,
                    progress=None) -> list[Finding]:
    """Находки vision_diff / vision_skipped. Исключение наружу не выходит никогда."""
    try:
        content_list = _content_list(zip_bytes, source_name)
        archive = BytesIO(zip_bytes)
        archive.name = f"{source_name}: zip MinerU"     # model_pages берёт .name только в текст ошибки
        pages = model_pages(archive)
        tokens, spans = plain_tokens(md)
        cases, skipped = build_windows(tokens, pdf_bytes, content_list, pages)
        verifier = make_verifier(model)                 # после content_list: без него верификатор не создаётся
        keys = list(page_tiles(cases))
        wrapper = _Progress(verifier, keys, progress)
        rows = run_measure(cases, wrapper)["cases"]
        found, flooded = limit_pages(hints(rows))

        findings = [Finding(rule="vision_diff", severity=SEVERITY["vision_diff"], page=hint.page,
                            snippet=md[spans[hint.start][0]:spans[hint.end - 1][1]], suggestion=hint.suggestion,
                            reading=hint.reading, model=model) for hint in found]
        stop = next((row for row in rows if row["error_reason"] == "stop"), None)
        if stop is not None:            # run_measure остановился на последней начатой странице
            rest = keys[wrapper.done - 1:]
            findings.append(skipped_finding(
                rest[0][1], f"сверка остановлена на стр. {rest[0][1]}: {stop['error']}; не сверены стр. "
                            + ", ".join(str(page) for _, page in rest), model))
        failed = dict.fromkeys((row["page"], row["error"]) for row in rows
                               if row["error_reason"] not in (None, "stop"))
        findings += [skipped_finding(page, f"стр. {page} не сверена: {error}", model) for page, error in failed]
        findings += [skipped_finding(page, f"стр. {page}: подсказок {n} — больше {PAGE_HINT_LIMIT}, похоже на сбой "
                                           f"нарезки или скана; подсказки страницы не выданы, сверить страницу "
                                           f"глазами", model) for page, n in flooded.items()]
        unlocated = skipped["unlocated"] + skipped["unresolved"]
        unreadable = sum(row["verdict"] == "unreadable" for row in rows)
        clipped = sum(row["reading"] == "clipped" for row in rows)
        total = len(range(0, len(tokens), VISION_STRIDE))
        if unlocated + unreadable + clipped:
            findings.append(skipped_finding(
                None, f"не сверено окон {unlocated + unreadable + clipped} из {total}: не привязаны к скану — "
                      f"{unlocated}, не прочитаны — {unreadable}, у края полосы — {clipped}", model))
        calls = getattr(verifier, "network_calls", 0)
        print(f"сверка по картинке ({model}): стр. {len(keys)}, вызовов claude {calls}, подсказок {len(found)}",
              file=sys.stderr)
        return findings
    except Exception as exc:
        traceback.print_exc(file=sys.stderr)
        return [skipped_finding(None, f"сверка не выполнена: {type(exc).__name__}: {exc}", model)]
