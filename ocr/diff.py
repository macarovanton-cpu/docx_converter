"""ПРАВКА #65: сверка двух OCR-прогонов, расхождения -> находки low_confidence."""

import difflib
import re
import unicodedata

from ocr import Finding
from ocr.postprocess import HOMOGLYPHS_CYR_TO_LAT
from ocr.validate import page_of

SNIPPET_MAX = 120

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
    """NFKC, «№» отдельным словом, гомоглифы → кириллица, регистр, пунктуация → пробел."""
    text = unicodedata.normalize("NFKC", text)
    text = _NUMERO_GAP_RE.sub("№ ", _NUMERO_RE.sub("№ ", text))
    # свёртка до lower(): в таблице есть пары только для заглавных
    text = text.translate(_TO_CYRILLIC).lower().replace("ё", "е")
    return " ".join(_NOT_WORD_RE.sub(" ", text).split())


def _words(md: str) -> tuple[list[str], list[tuple[int, int]]]:
    """Слова стороны и срез исходного токена в md для каждого слова.

    Склеенное OCR-ом «твердыми,сыпучими» даёт два слова с одним срезом.
    """
    words: list[str] = []
    spans: list[tuple[int, int]] = []
    offset = 0
    for line in md.split("\n"):
        if not _SEPARATOR_RE.match(line):
            for match in _TOKEN_RE.finditer(line):
                span = (offset + match.start(), offset + match.end())
                for word in normalize(match.group()).split():
                    words.append(word)
                    spans.append(span)
        offset += len(line) + 1
    return words, spans


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


def diff_findings(primary_md: str, secondary_md: str,
                  content_list: list | None = None) -> list[Finding]:
    """Расхождения пословного выравнивания двух прогонов → находки low_confidence.

    Текст не меняется, «кто прав» не выбирается. Порядок находок — по позиции
    в primary_md; одинаковые подряд не схлопываются: каждая — своё место.
    """
    words_primary, spans_primary = _words(primary_md)
    words_secondary, spans_secondary = _words(secondary_md)
    matcher = difflib.SequenceMatcher(None, words_primary, words_secondary,
                                      autojunk=False)
    findings = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        spans = spans_primary[i1:i2] or _anchor(spans_primary, i1)
        snippet = _slice(primary_md, spans) if spans else ""
        suggestion = _slice(secondary_md, spans_secondary[j1:j2]) if j1 < j2 else None
        findings.append(Finding(rule="low_confidence", severity="warning",
                                page=page_of(snippet, content_list),
                                snippet=snippet, suggestion=suggestion))
    return findings
