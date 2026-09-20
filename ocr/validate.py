"""ПРАВКА #64: валидатор OCR-markdown, report.json, пометки в markdown."""

import re
from collections import Counter
from dataclasses import replace
from datetime import datetime, timezone
from itertools import permutations

from ocr import SEVERITIES, Finding
from ocr.postprocess import fold_to_cyrillic, parse_pipe_tables

REPORT_SCHEMA_VERSION = 1
ANNOTATION_PREFIX = "!! ПРОВЕРИТЬ: "
ANNOTATION_SUFFIX = " !!"
# ПРАВКА #70: находки без места в тексте — в хвост документа, своим разделом
LOST_HEADING = "## Не привязанные находки"

UNITS = ("мм", "см", "м", "км", "г", "кг", "т", "л", "шт", "В", "А", "Вт", "кВт",
         "кВА", "Гц", "лм", "ч", "с", "мин", "сут")
SCALE_FAMILIES = ("ВЕСТА-С", "РЕУС-Д", "ВЕКТОР-ДВ")

# серьёзность правила — константа (README, «Правила и их серьёзность»)
SEVERITY = {
    "inn_checksum": "critical",
    "ogrn_checksum": "critical",
    "kpp_format": "warning",
    "gost_format": "warning",
    "unit_unknown": "warning",
    "scale_model": "warning",
    "table_total_mismatch": "critical",
    "table_row_cells": "warning",
    "table_empty_number": "warning",
    "code_digits_glued": "warning",              # ПРАВКА #79
    "org_name_variant": "warning",               # ПРАВКА #80
}

GLUED_CODE_SUGGESTION = "проверить границу с следующим пунктом"
ORG_NAME_MIN_LEN = 5

ANNOTATION_LIMIT = 200

INN_WEIGHTS_10 = (2, 4, 10, 3, 5, 9, 4, 6, 8)
INN_WEIGHTS_11 = (7, 2, 4, 10, 3, 5, 9, 4, 6, 8)
INN_WEIGHTS_12 = (3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8)

_UNIT_BY_LOWER = {unit.lower(): unit for unit in UNITS}

_INN_RE = re.compile(r"ИНН[^\d]{0,3}(\d+)")
_OGRN_RE = re.compile(r"ОГРН(?:ИП)?[^\d]{0,3}(\d+)")
_KPP_RE = re.compile(r"КПП\W{0,3}(\w+)")
_KPP_VALUE_RE = re.compile(r"\d{4}[\dA-Z]{2}\d{3}")
_WORD_RE = re.compile(r"[^\W\d_]+")
# ПРАВКА #77: тире в обозначении бывает дефисом, en- и em-dash («ГОСТ Р 58760—2019»)
_DASH = "[-–—]"
_GOST_TAIL_RE = re.compile(rf"( Р)? \d+(\.\d+)*{_DASH}\d{{2,4}}")
# ПРАВКА #77: «ГОСТ»/«ГОСТ Р» без номера — упоминание в прозе, а не обозначение.
# Пробелы необязательны: «ГОСТ Р53228» и «ГОСТ8.726» — номер, только слипшийся.
_GOST_NUMBER_RE = re.compile(r"( ?Р)? ?\d")
_TU_RE = re.compile(r"(?<![^\W\d_])ТУ(?= ?\d)")
_TU_TAIL_RE = re.compile(rf" ?\d+(\.\d+)*{_DASH}\d+{_DASH}\d+{_DASH}\d{{2,4}}")
# ПРАВКА #79: обозначение стандарта с номером; последняя числовая группа длиннее
# четырёх цифр — к году прилип номер следующего пункта («СП 76.13330.20163»)
_GLUED_CODE_RE = re.compile(
    rf"(?<![^\W\d_])(?:СП|СНиП|ГОСТ(?: Р)?|ТУ) ?\d+(?:(?:\.|{_DASH})\d+)*")
# ПРАВКА #80: название организации — то, что стоит в «ёлочках» (без вложенных)
_ORG_QUOTED_RE = re.compile(r"«([^«»]+)»")
_UNIT_RE = re.compile(r"\d ?([^\W\d_]{1,3})(?![^\W\d_])")
_NUMBER_CELL_RE = re.compile(r"\d+(\.\d+)*\.?")
_TAG_RE = re.compile(r"<[^>]+>")
_ANNOTATION_RE = re.compile(
    re.escape(ANNOTATION_PREFIX) + ".*" + re.escape(ANNOTATION_SUFFIX))

# краевая пунктуация токена; дефис не снимаем — он внутри марки весов
_EDGE_PUNCT = " \t.,;:!?()[]{}«»\"'|*_`"

SNIPPET_TAIL = 20


def _finding(rule: str, snippet: str, suggestion: str | None = None) -> Finding:
    """Находка валидатора: page всегда None (страницу проставляет build_report)."""
    return Finding(rule=rule, severity=SEVERITY[rule], page=None,
                   snippet=snippet, suggestion=suggestion)


def _has_latin(token: str) -> bool:
    return any("a" <= char <= "z" or "A" <= char <= "Z" for char in token)


# --- реквизиты --------------------------------------------------------------

def _control(digits: str, weights: tuple) -> int:
    return sum(int(d) * w for d, w in zip(digits, weights)) % 11 % 10


def _inn_ok(digits: str) -> bool:
    if len(digits) == 10:
        return _control(digits, INN_WEIGHTS_10) == int(digits[9])
    if len(digits) == 12:
        return (_control(digits, INN_WEIGHTS_11) == int(digits[10])
                and _control(digits, INN_WEIGHTS_12) == int(digits[11]))
    return False


def _ogrn_ok(digits: str) -> bool:
    if len(digits) == 13:
        return int(digits[:12]) % 11 % 10 == int(digits[12])
    if len(digits) == 15:
        return int(digits[:14]) % 13 % 10 == int(digits[14])
    return False


def check_requisites(md: str) -> list[Finding]:
    """ИНН и ОГРН — длина и контрольная сумма; КПП — формат."""
    findings = []
    for match in _INN_RE.finditer(md):
        if not _inn_ok(match.group(1)):
            findings.append(_finding("inn_checksum", match.group()))
    for match in _OGRN_RE.finditer(md):
        if not _ogrn_ok(match.group(1)):
            findings.append(_finding("ogrn_checksum", match.group()))
    for match in _KPP_RE.finditer(md):
        if not _KPP_VALUE_RE.fullmatch(match.group(1)):
            findings.append(_finding("kpp_format", match.group()))
    return findings


# --- ГОСТ / ТУ --------------------------------------------------------------

def _snippet_from(md: str, start: int, end: int) -> str:
    """Токен плюс хвост строки — чтобы находка читалась в отчёте."""
    return md[start:end + SNIPPET_TAIL].split("\n")[0].rstrip()


def check_standards(md: str) -> list[Finding]:
    """«ГОСТ» гомоглифами/регистром и номер не по шаблону; ТУ — только перед цифрой.

    ПРАВКА #77: номер проверяется только там, где он есть. «ссылки на ГОСТ, ТУ»
    и «сертификат (ГОСТ Р)» — проза, а не обозначение с потерянным номером.
    """
    findings = []
    for match in _WORD_RE.finditer(md):
        token = match.group()
        folded = fold_to_cyrillic(token)
        if len(token) != 4 or folded is None or folded.upper() != "ГОСТ":
            continue
        if token != "ГОСТ":
            findings.append(_finding("gost_format", token, "ГОСТ"))
        if (_GOST_NUMBER_RE.match(md, match.end())
                and not _GOST_TAIL_RE.match(md, match.end())):
            findings.append(_finding(
                "gost_format", _snippet_from(md, match.start(), match.end())))
    for match in _TU_RE.finditer(md):
        if not _TU_TAIL_RE.match(md, match.end()):
            findings.append(_finding(
                "gost_format", _snippet_from(md, match.start(), match.end())))
    return findings


def check_glued_codes(md: str) -> list[Finding]:
    """ПРАВКА #79: к номеру стандарта прилип номер следующего пункта.

    Последняя группа составного номера — год, он не длиннее четырёх цифр.
    Длиннее — значит, OCR склеил обозначение со следующим пунктом списка:
    «СП 76.13330.20163. Подрядчик…» — это «СП 76.13330.2016» и пункт «3.».
    Номер из одной группы («ГОСТ Р53228») не год и под правило не попадает.
    Текст не правим: где именно граница, видно только по скану.
    """
    findings = []
    for match in _GLUED_CODE_RE.finditer(md):
        groups = re.findall(r"\d+", match.group())
        if len(groups) > 1 and len(groups[-1]) > 4:
            findings.append(_finding("code_digits_glued", match.group(),
                                     GLUED_CODE_SUGGESTION))
    return findings


# --- единицы и марки весов --------------------------------------------------

def check_units(md: str) -> list[Finding]:
    """Единица после числа, записанная с латиницей: «10MM» → «мм»."""
    findings = []
    for match in _UNIT_RE.finditer(md):
        token = match.group(1)
        folded = fold_to_cyrillic(token)
        if folded is None or not _has_latin(token):
            continue
        canonical = _UNIT_BY_LOWER.get(folded.lower())
        if canonical is not None and token not in UNITS:
            findings.append(_finding("unit_unknown", match.group(), canonical))
    return findings


def _one_letter_apart(rare: str, common: str) -> bool:
    """Одна подстановка буквы. Перенос строки вместо пробела вариантом не считается."""
    if len(rare) != len(common):
        return False
    diff = [(a, b) for a, b in zip(rare, common) if a != b]
    return len(diff) == 1 and diff[0][0].isalpha() and diff[0][1].isalpha()


def check_org_names(md: str) -> list[Finding]:
    """ПРАВКА #80: название организации в кавычках, отличающееся на одну букву.

    «ГПИ имени Д.С. Косьяна» против «ГПП имени Д.С. Косьяна» — одна и та же
    контора, букву разобрало двумя способами. Редкий вариант идёт в находку,
    частый — в предложение; текст не правим, какой верен — видно по скану.
    Равная частота находки не даёт: кто из двоих опечатка, тогда не сказать.
    """
    counts = Counter(match.group(1) for match in _ORG_QUOTED_RE.finditer(md))
    names = [name for name in counts if len(name) >= ORG_NAME_MIN_LEN]
    return [_finding("org_name_variant", f"«{rare}»", f"«{common}»")
            for rare, common in permutations(names, 2)
            if counts[rare] < counts[common] and _one_letter_apart(rare, common)]


def check_scale_models(md: str) -> list[Finding]:
    """Марка весов не каноническим написанием: «BЕСТА-С60» → «ВЕСТА-С60»."""
    findings = []
    for raw in md.split():
        token = raw.strip(_EDGE_PUNCT)
        folded = fold_to_cyrillic(token)
        if not token or folded is None:
            continue
        upper = folded.upper()
        for family in SCALE_FAMILIES:
            if upper.startswith(family) and not token.startswith(family):
                findings.append(_finding(
                    "scale_model", token, family + token[len(family):]))
                break
    return findings


# --- таблицы ----------------------------------------------------------------

def _as_number(cell: str) -> float | None:
    """«1 000» → 1000.0, «10,5» → 10.5; не число → None."""
    text = cell.replace(",", ".").replace(" ", "").replace(" ", "")
    try:
        return float(text)
    except ValueError:
        return None


def _format_number(value: float) -> str:
    return f"{value:.2f}".rstrip("0").rstrip(".")


def _longest_cell(row: list[str]) -> str:
    """Самая длинная ячейка строки — дословный фрагмент, по нему ищется место пометки."""
    return max(row, key=len) if row else ""


def _is_numbered_column(rows: list[list[str]]) -> bool:
    """Первая колонка нумерационная: не меньше половины строк (кроме первой) — номера."""
    rest = rows[1:]
    if not rest:
        return False
    numbers = sum(1 for row in rest if row and _NUMBER_CELL_RE.fullmatch(row[0]))
    return numbers * 2 >= len(rest)


def _check_totals(rows: list[list[str]]) -> list[Finding]:
    findings = []
    for index, row in enumerate(rows):
        first = next((cell for cell in row if cell), "")
        if not first.lower().startswith("итого"):
            continue
        for column in range(len(row)):
            total = _as_number(row[column])
            values = [_as_number(above[column]) for above in rows[1:index]
                      if column < len(above)]
            if total is None or not values or any(v is None for v in values):
                continue
            if abs(sum(values) - total) > 0.01:
                findings.append(_finding("table_total_mismatch", row[column],
                                         _format_number(sum(values))))
    return findings


def check_tables(md: str) -> list[Finding]:
    """«Итого», число ячеек в строке, пустой номер пункта."""
    findings = []
    for rows in parse_pipe_tables(md):
        findings += _check_totals(rows)
        width = len(rows[0])
        for row in rows[1:]:
            if len(row) != width:
                findings.append(_finding("table_row_cells", _longest_cell(row)))
        if not _is_numbered_column(rows):
            continue
        # первая строка проверяется нарочно: у продолжения, разорванного
        # страницей, «шапкой» становится строка с потерянным номером
        for row in rows:
            if row and not row[0] and any(row[1:]):
                findings.append(_finding("table_empty_number", _longest_cell(row)))
    return findings


# --- публичные функции ------------------------------------------------------

def validate(md: str, content_list: list | None = None) -> list[Finding]:
    """Все правила спеки 05 над постобработанным markdown."""
    findings = (check_requisites(md) + check_standards(md) + check_glued_codes(md)
                + check_units(md) + check_org_names(md) + check_scale_models(md)
                + check_tables(md))
    if content_list is None:
        return findings
    return [replace(f, page=page_of(f.snippet, content_list)) for f in findings]


def page_of(snippet: str, content_list: list | None) -> int | None:
    """Страница (1-based) первого блока content_list, содержащего snippet.

    ПРАВКА #72: PLACEHOLDER снят — форма блока сверена с настоящим
    content_list.json из vlm_raw.zip (живой прогон vlm), тест
    tests/test_ocr_validate.py::test_pages_from_real_content_list.
    """
    if not content_list:
        return None
    needle = " ".join(snippet.split())
    if not needle:
        return None
    for block in content_list:
        # text / table_body — взаимоисключающие и необязательные, отсюда .get
        text = block.get("text") or block.get("table_body") or ""
        if needle in " ".join(_TAG_RE.sub(" ", text).split()):
            return block["page_idx"] + 1 if "page_idx" in block else None
    return None


def build_report(*, source: str, sha256: str, provider: str,
                 model_version: str | None, cache_hit: bool, verified: bool,
                 findings: list[Finding],
                 content_list: list | None = None) -> dict:
    """report.json схемы v1: ключи верхнего уровня и ключи находки — закрытые списки."""
    items = []
    summary = {severity: 0 for severity in SEVERITIES}
    for index, finding in enumerate(findings, 1):
        page = finding.page
        if page is None:
            page = page_of(finding.snippet, content_list)
        items.append({"id": index, "rule": finding.rule, "severity": finding.severity,
                      "page": page, "snippet": finding.snippet,
                      "suggestion": finding.suggestion})
        summary[finding.severity] += 1
    return {
        "schema_version": REPORT_SCHEMA_VERSION,
        "source": source,
        "sha256": sha256,
        "provider": provider,
        "model_version": model_version,
        "cache_hit": cache_hit,
        "verified": verified,
        "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "summary": summary,
        "findings": items,
    }


def _annotation(finding: dict) -> str:
    text = (finding["suggestion"] or finding["snippet"]).replace("\n", " ")
    if len(text) > ANNOTATION_LIMIT:
        text = text[:ANNOTATION_LIMIT - 1] + "…"
    return (f"{ANNOTATION_PREFIX}[{finding['id']}] {finding['rule']}: "
            f"{text}{ANNOTATION_SUFFIX}")


def _split_blocks(md: str) -> tuple[list[str], str]:
    """Блоки по '\\n\\n' плюс хвостовые переводы строк — чтобы round-trip был точным."""
    body = md.rstrip("\n")
    return body.split("\n\n"), md[len(body):]


def annotate(md: str, report: dict, *, include_low_confidence: bool = False) -> str:
    """Пометки «!! ПРОВЕРИТЬ: … !!» отдельными блоками после места находки.

    ПРАВКА #70: `low_confidence` по умолчанию пропускаются — их на документ
    десятки, и в тексте они забивают находки правил; отчёт их всё равно
    содержит целиком, а `--annotate-all` возвращает их в текст. Находка, чей
    фрагмент в документе не нашёлся, уходит в хвост под заголовок
    LOST_HEADING — в начало документа ей нельзя, там она врезается в шапку.
    """
    blocks, trailing = _split_blocks(md)
    after: dict[int, list[str]] = {}
    lost: list[str] = []
    for finding in sorted(report["findings"], key=lambda f: f["id"]):
        if finding["rule"] == "low_confidence" and not include_low_confidence:
            continue
        note = _annotation(finding)
        index = next((i for i, block in enumerate(blocks)
                      if finding["snippet"] and finding["snippet"] in block), None)
        if index is None:
            lost.append(note)                    # находка не теряется
        else:
            after.setdefault(index, []).append(note)
    out = []
    for index, block in enumerate(blocks):
        out.append(block)
        out += after.get(index, [])
    if lost:
        out += [LOST_HEADING] + lost
    return "\n\n".join(out) + trailing


def strip_annotations(md: str) -> str:
    """Снять пометки валидатора; прочие callout-ы («!! формула !!») не трогать."""
    blocks, trailing = _split_blocks(md)
    kept = [block for block in blocks if not _ANNOTATION_RE.fullmatch(block)]
    if LOST_HEADING in blocks:
        tail = blocks[blocks.index(LOST_HEADING) + 1:]
        if tail and all(_ANNOTATION_RE.fullmatch(block) for block in tail):
            kept.remove(LOST_HEADING)            # заголовок поставил annotate
    return "\n\n".join(kept) + trailing
