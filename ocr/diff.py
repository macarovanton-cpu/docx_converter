"""ПРАВКА #65: сверка двух OCR-прогонов, расхождения -> находки low_confidence."""

import difflib
import re
import unicodedata

from ocr import Finding
from ocr.postprocess import HOMOGLYPHS_CYR_TO_LAT
from ocr.validate import page_of

SNIPPET_MAX = 120
# ПРАВКА #70: расхождение короче — джиттер OCR на одну-две буквы, не находка
MIN_RUN = 3

# свёртка латиница → кириллица, безусловная: обе стороны сворачиваются одинаково,
# «IР54» и «IP54» обязаны совпасть. Цена — «РоЕ»/«PoE» неразличимы (починено в 04).
_TO_CYRILLIC = str.maketrans({lat: cyr for cyr, lat in HOMOGLYPHS_CYR_TO_LAT.items()})

_NUMERO_RE = re.compile(r"\bNo\.?\s*(?=\d|п/п)")
_NUMERO_GAP_RE = re.compile(r"№\s*")
_NOT_WORD_RE = re.compile(r"[^\w№]|_")
# токен — всё между пробелами и трубами разметки
_TOKEN_RE = re.compile(r"[^\s|]+")
# строка-разделитель pipe-таблицы: только пробелы, трубы, двоеточия и дефисы
_SEPARATOR_RE = re.compile(r"^[\s|:-]*-[\s|:-]*$")


def normalize(text: str) -> str:
    """NFKC, «No» → «№», гомоглифы → кириллица, регистр, пунктуация и пробелы — прочь.

    ПРАВКА #70: пробелы и переносы удаляются целиком. Разошлась только разбивка
    на слова («твердыми, сыпучими» против «твердыми,сыпучими») — значит прогоны
    прочли одно и то же, и сверять тут нечего.
    """
    text = unicodedata.normalize("NFKC", text)
    text = _NUMERO_GAP_RE.sub("№", _NUMERO_RE.sub("№", text))
    # свёртка до lower(): в таблице есть пары только для заглавных
    text = text.translate(_TO_CYRILLIC).lower().replace("ё", "е")
    return "".join(_NOT_WORD_RE.sub(" ", text).split())


def _tokens(md: str) -> tuple[list[str], list[tuple[int, int]]]:
    """Нормализованные токены стороны и срез каждого из них в исходном md."""
    keys: list[str] = []
    spans: list[tuple[int, int]] = []
    offset = 0
    for line in md.split("\n"):
        if not _SEPARATOR_RE.match(line):
            for match in _TOKEN_RE.finditer(line):
                key = normalize(match.group())
                if key:                      # от токена осталась одна разметка
                    keys.append(key)
                    spans.append((offset + match.start(), offset + match.end()))
        offset += len(line) + 1
    return keys, spans


def _chars(keys: list[str]) -> tuple[str, list[int]]:
    """Символы куска подряд и номер токена-владельца для каждого символа."""
    owners: list[int] = []
    for index, key in enumerate(keys):
        owners += [index] * len(key)
    return "".join(keys), owners


def _slice(md: str, spans: list[tuple[int, int]]) -> str:
    """Дословный срез от начала первого токена до конца последнего, не длиннее SNIPPET_MAX.

    Обрезка — по границе токена; токен длиннее лимита (ссылка на картинку MinerU)
    режется по лимиту, внутри токена пробелов нет.
    """
    start, end = spans[0]
    for token_start, token_end in spans:
        if token_end - start > SNIPPET_MAX:
            break
        end = token_end
    return md[start:min(end, start + SNIPPET_MAX)]


def _anchor(spans: list[tuple[int, int]], index: int) -> list[tuple[int, int]]:
    """Для insert (у primary пусто) — соседний токен слева, нет слева — справа."""
    return spans[index - 1:index] if index else spans[:1]


def _divergences(text_a: str, owners_a: list[int],
                 text_b: str, owners_b: list[int]) -> list[list[list[int]]]:
    """Символьные расхождения куска → пары списков затронутых токенов.

    ПРАВКА #70: опкоды, попавшие в один и тот же токен, — одна находка
    («Hypeeb» против «Hyreev» расходится в двух местах, но место одно).
    Расхождение без единой буквы и цифры (только «№») находкой не считается.
    """
    out: list[list[list[int]]] = []
    for tag, x1, x2, y1, y2 in difflib.SequenceMatcher(
            None, text_a, text_b, autojunk=False).get_opcodes():
        if tag == "equal":
            continue
        if not any(char.isalnum() for char in text_a[x1:x2] + text_b[y1:y2]):
            continue
        tokens_a = sorted(set(owners_a[x1:x2]))
        if not tokens_a:                     # вставка: своих символов нет,
            tokens_a = owners_a[x1 - 1:x1] or owners_a[:1]   # цепляемся к соседу
        tokens_b = sorted(set(owners_b[y1:y2]))
        if out and tokens_a and out[-1][0] and tokens_a[0] <= out[-1][0][-1]:
            out[-1][0] = sorted(set(out[-1][0]) | set(tokens_a))
            out[-1][1] = sorted(set(out[-1][1]) | set(tokens_b))
        else:
            out.append([tokens_a, tokens_b])
    return out


def diff_findings(primary_md: str, secondary_md: str,
                  content_list: list | None = None) -> list[Finding]:
    """Расхождения двух прогонов → находки low_confidence.

    Два уровня: грубое выравнивание по токенам (быстро) и посимвольная сверка
    внутри каждого расхождения. Глобальный посимвольный проход на паре фикстур
    идёт 13 с против 0,1 с у двухуровневого — при том же результате.

    Текст не меняется, «кто прав» не выбирается. Порядок находок — по позиции
    в primary_md; одинаковые подряд не схлопываются: каждая — своё место.
    """
    keys_a, spans_a = _tokens(primary_md)
    keys_b, spans_b = _tokens(secondary_md)
    matcher = difflib.SequenceMatcher(None, keys_a, keys_b, autojunk=False)
    findings = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        text_a, owners_a = _chars(keys_a[i1:i2])
        text_b, owners_b = _chars(keys_b[j1:j2])
        if text_a == text_b:
            continue                         # разошлась только разбивка на слова
        for tokens_a, tokens_b in _divergences(text_a, owners_a, text_b, owners_b):
            spans = [spans_a[i1 + index] for index in tokens_a] or _anchor(spans_a, i1)
            snippet = _slice(primary_md, spans) if spans else ""
            if len(snippet) < MIN_RUN:       # одна-две буквы — джиттер, не находка
                continue
            suggestion = (_slice(secondary_md, [spans_b[j1 + index] for index in tokens_b])
                          if tokens_b else None)
            findings.append(Finding(rule="low_confidence", severity="warning",
                                    page=page_of(snippet, content_list),
                                    snippet=snippet, suggestion=suggestion))
    return findings
