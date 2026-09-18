"""ПРАВКА #63: детерминированный постпроцессор markdown после OCR MinerU."""

import re
from html.parser import HTMLParser

from markdown_cleanup import cleanup_ocr_markdown
from ocr import Finding

# канонические написания; сравнение — по свёртке в латиницу и верхнему регистру
KNOWN_LATIN = ("PoE", "IP", "SNR", "ITV", "VIDEOMAX", "RS", "SFP", "SQL",
               "USB", "PDF", "DWG", "Ethernet", "Parsec")

HOMOGLYPHS_CYR_TO_LAT = {
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O",
    "Р": "P", "С": "C", "Т": "T", "Х": "X", "І": "I",
    "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "у": "y", "х": "x", "і": "i",
}

# обратная таблица однозначна: значения HOMOGLYPHS_CYR_TO_LAT не повторяются,
# I/i достаются украинским І/і — годятся для инициалов и единиц, не для слов
_LAT_TO_CYR = {lat: cyr for cyr, lat in HOMOGLYPHS_CYR_TO_LAT.items()}
_KNOWN_BY_UPPER = {word.upper(): word for word in KNOWN_LATIN}

# серьёзность правила — константа (README, «Правила и их серьёзность»)
SEVERITY = {
    "html_table_unparsed": "critical",
    "table_merge_failed": "critical",
    "table_merged": "warning",
    "table_span": "info",
    "mixed_alphabet_unknown": "warning",
    "translit_suspect": "critical",
    "signature_block": "warning",
}

SNIPPET_LIMIT = 80

# «инициалы + фамилия»; алфавит не различает — латиницу проверяет вызывающий
NAME_RE = re.compile(r"^\s*[A-Za-zА-ЯЁ]\.\s?[A-Za-zА-ЯЁ]{1,3}\.\s+[A-Za-zА-Яа-яЁё-]+\s*$")

_TOKEN_RE = re.compile(r"[^\W_]+")
_NUMERO_RE = re.compile(r"\bNo\.?[ \t]*(?=\d|п/п)")
_DEGREE_RE = re.compile(
    r"[ \t]*\$\s*(?:C\s*\^\s*\{\s*\\circ\s*\}|\^\s*\{\s*\\circ\s*\}\s*C)\s*\$")
_DEGREE_TAIL_RE = re.compile(r"°C[ \t]+(?=[;.,)])")
# ПРАВКА #68: «;» или «:» вплотную к маркеру списка; «:---» — разделитель таблицы
_LIST_GLUE_RE = re.compile(r"(?<=[;:])-(?!-)")
_TABLE_TAG_RE = re.compile(r"</?table\b[^>]*>", re.I)
_CELL_SPLIT_RE = re.compile(r"(?<!\\)\|")
_SEP_CELL_RE = re.compile(r":?-+:?")
_IP_CODE_RE = re.compile(r"IP\d{2}")


def _finding(rule: str, snippet: str, suggestion: str | None = None) -> Finding:
    """Находка постпроцессора: page всегда None (страницу проставляет build_report)."""
    return Finding(rule=rule, severity=SEVERITY[rule], page=None,
                   snippet=snippet, suggestion=suggestion)


def _is_cyrillic(char: str) -> bool:
    return "Ѐ" <= char <= "ӿ"


def _is_latin(char: str) -> bool:
    return "a" <= char <= "z" or "A" <= char <= "Z"


def fold_to_latin(token: str) -> str | None:
    """Кириллические гомоглифы → латиница; буква без пары → None."""
    out = []
    for char in token:
        if _is_cyrillic(char):
            if char not in HOMOGLYPHS_CYR_TO_LAT:
                return None
            out.append(HOMOGLYPHS_CYR_TO_LAT[char])
        else:
            out.append(char)
    return "".join(out)


def fold_to_cyrillic(token: str) -> str | None:
    """Латинские гомоглифы → кириллица; буква без пары → None."""
    out = []
    for char in token:
        if _is_latin(char):
            if char not in _LAT_TO_CYR:
                return None
            out.append(_LAT_TO_CYR[char])
        else:
            out.append(char)
    return "".join(out)


# --- pipe-таблицы -----------------------------------------------------------

def _split_row(line: str) -> list[str]:
    body = line.strip()
    if body.startswith("|"):
        body = body[1:]
    if body.endswith("|") and not body.endswith("\\|"):
        body = body[:-1]
    return [cell.strip().replace("\\|", "|") for cell in _CELL_SPLIT_RE.split(body)]


def _is_separator(cells: list[str]) -> bool:
    return bool(cells) and all(_SEP_CELL_RE.fullmatch(cell) for cell in cells)


def parse_pipe_tables(md: str) -> list[list[list[str]]]:
    r"""Таблицы → строки → ячейки. Строка-разделитель исключена, \| раскрыт."""
    tables: list[list[list[str]]] = []
    current: list[list[str]] | None = None
    for line in md.split("\n"):
        if not line.startswith("|"):
            current = None
            continue
        if current is None:
            current = []
            tables.append(current)
        cells = _split_row(line)
        if not _is_separator(cells):
            current.append(cells)
    return [table for table in tables if table]


def _render_rows(rows: list[list[str]], with_separator: bool) -> list[str]:
    lines = ["| " + " | ".join(cell.replace("|", "\\|") for cell in row) + " |"
             for row in rows]
    if with_separator and lines:
        lines.insert(1, "|" + "|".join("---" for _ in rows[0]) + "|")
    return lines


# --- 1. HTML-таблицы --------------------------------------------------------

class _TableParser(HTMLParser):
    """<table> → строки ячеек (текст, colspan, rowspan). Вложенная table — отказ."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)      # convert_charrefs == html.unescape
        self.rows: list[list[tuple[str, int, int]]] = []
        self.failed = False
        self._depth = 0
        self._row: list[tuple[str, int, int]] | None = None
        self._cell: list[str] | None = None
        self._attrs: dict = {}

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "table":
            self._depth += 1
            if self._depth > 1:
                self.failed = True
        elif tag == "tr":
            self._close_cell()
            self._close_row()
            self._row = []
        elif tag in ("td", "th"):
            self._close_cell()
            if self._row is None:
                self._row = []
            self._cell = []
            self._attrs = dict(attrs)
        elif tag == "br" and self._cell is not None:
            self._cell.append(" ")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("td", "th"):
            self._close_cell()
        elif tag == "tr":
            self._close_cell()
            self._close_row()
        elif tag == "table":
            self._close_cell()
            self._close_row()
            self._depth -= 1

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)

    def _close_cell(self) -> None:
        if self._cell is None:
            return
        text = re.sub(r"\s+", " ", "".join(self._cell)).strip()
        if self._row is None:
            self._row = []
        self._row.append((text, _span(self._attrs, "colspan"), _span(self._attrs, "rowspan")))
        self._cell = None
        self._attrs = {}

    def _close_row(self) -> None:
        if self._row:
            self.rows.append(self._row)
        self._row = None


def _span(attrs: dict, name: str) -> int:
    """colspan/rowspan — необязательные атрибуты, отсюда .get; мусор → 1."""
    try:
        value = int(attrs.get(name) or 1)
    except ValueError:
        return 1
    return value if value >= 1 else 1


def _table_regions(md: str) -> list[tuple[int, int, bool]]:
    """Границы таблиц верхнего уровня: (start, end, closed). Вложенные — внутри."""
    regions = []
    depth = 0
    start = 0
    for match in _TABLE_TAG_RE.finditer(md):
        if match.group().startswith("</"):
            if depth:
                depth -= 1
                if depth == 0:
                    regions.append((start, match.end(), True))
        else:
            if depth == 0:
                start = match.start()
            depth += 1
    if depth:
        regions.append((start, len(md), False))      # незакрытый тег
    return regions


def _grid_from_rows(rows) -> tuple[list[list[str]], str | None]:
    """Развернуть colspan/rowspan в прямоугольные ячейки; текст первой span-ячейки."""
    grid: list[list[str]] = []
    carry: dict[int, list] = {}
    spanned: str | None = None
    for cells in rows:
        row: list[str] = []
        col = 0
        pending = list(cells)
        while True:
            if col in carry:
                row.append(carry[col][0])
                carry[col][1] -= 1
                if carry[col][1] == 0:
                    del carry[col]
                col += 1
                continue
            if not pending:
                break
            text, colspan, rowspan = pending.pop(0)
            if (colspan > 1 or rowspan > 1) and spanned is None:
                spanned = text
            for index in range(colspan):
                row.append(text if index == 0 else "")
                if rowspan > 1:
                    carry[col] = [row[-1], rowspan - 1]
                col += 1
        grid.append(row)
    return grid, spanned


def html_tables_to_pipe(md: str) -> tuple[str, list[Finding]]:
    """<table> → pipe-таблица. Неразобранная таблица остаётся как есть + находка."""
    findings: list[Finding] = []
    out: list[str] = []
    position = 0
    for start, end, closed in _table_regions(md):
        out.append(md[position:start])
        raw = md[start:end]
        position = end
        parser = _TableParser()
        if closed:
            parser.feed(raw)
            parser.close()
        if not closed or parser.failed or not parser.rows:
            findings.append(_finding("html_table_unparsed", raw[:SNIPPET_LIMIT]))
            out.append(raw)
            continue
        grid, spanned = _grid_from_rows(parser.rows)
        if spanned is not None:
            findings.append(_finding("table_span", spanned[:SNIPPET_LIMIT]))
        out.append("\n\n" + "\n".join(_render_rows(grid, True)) + "\n\n")
    out.append(md[position:])
    return "".join(out), findings


# --- 2. склейка таблиц через разрыв страницы --------------------------------

def _rows_of(lines: list[str]) -> tuple[list[list[str]], bool]:
    rows, separator = [], False
    for line in lines:
        cells = _split_row(line)
        if _is_separator(cells):
            separator = True
        else:
            rows.append(cells)
    return rows, separator


def _is_continuation(row: list[str]) -> bool:
    """Первая ячейка пуста и непустая ровно одна — хвост пункта с прошлой страницы."""
    return bool(row) and row[0] == "" and len([cell for cell in row if cell]) == 1


def _merge_pair(a_lines: list[str], b_lines: list[str],
                findings: list[Finding]) -> list[str] | None:
    a_rows, a_separator = _rows_of(a_lines)
    b_rows, _ = _rows_of(b_lines)
    if not a_rows or not b_rows or not a_rows[-1]:
        return None
    rest = [list(row) for row in b_rows]
    tail = []
    while rest and _is_continuation(rest[0]):
        tail.append(next(cell for cell in rest.pop(0) if cell))
    if not tail:
        return None
    if any(len(row) != len(a_rows[0]) for row in rest):
        nonempty = next((cell for row in b_rows for cell in row if cell), "")
        findings.append(_finding("table_merge_failed", nonempty[:SNIPPET_LIMIT]))
        return None
    merged = [list(row) for row in a_rows]
    target = merged[-1]
    added = " ".join(tail)
    target[-1] = (target[-1] + " " + added).strip()
    merged.extend(rest)
    findings.append(_finding(
        "table_merged", added[:SNIPPET_LIMIT],
        f"Продолжение дописано в строку «{target[0]}»; сверить со сканом"))
    return _render_rows(merged, a_separator)


def _is_row_tail(row: list[str], target: list[str]) -> bool:
    """ПРАВКА #69: номер потерян, содержимого ≥ 2 ячеек, ширина как у предыдущей строки.

    Ровно одна непустая ячейка — это не хвост, а разделитель («I. Общие данные»)
    или продолжение с прошлой страницы, которое уже разобрал _merge_pair.
    Ширина обязана совпасть: иначе лишние ячейки пришлось бы выбросить.
    """
    return (len(row) == len(target) and row[0] == ""
            and len([cell for cell in row if cell]) >= 2)


def _merge_row_tails(rows: list[list[str]],
                     findings: list[Finding]) -> list[list[str]] | None:
    """Хвосты дописываются в предыдущую строку поячеечно. Нечего сливать → None."""
    merged: list[list[str]] = []
    for row in rows:
        if merged and _is_row_tail(row, merged[-1]):
            target = merged[-1]
            for index, cell in enumerate(row):
                if cell:
                    target[index] = (target[index] + " " + cell).strip()
            added = " ".join(cell for cell in row if cell)
            findings.append(_finding(
                "table_merged", added[:SNIPPET_LIMIT],
                f"Продолжение дописано в строку «{target[0]}»; сверить со сканом"))
        else:
            merged.append(list(row))
    return merged if len(merged) != len(rows) else None


def _segments(lines: list[str]) -> list[list]:
    """Текст → чередование блоков ['table'|'other', строки]."""
    segments: list[list] = []
    for line in lines:
        kind = "table" if line.startswith("|") else "other"
        if segments and segments[-1][0] == kind:
            segments[-1][1].append(line)
        else:
            segments.append([kind, [line]])
    return segments


def _previous_table(out: list[list]) -> int | None:
    """Индекс предыдущей таблицы, если между ней и текущей только пустые строки."""
    for index in range(len(out) - 1, -1, -1):
        kind, block = out[index]
        if kind == "table":
            return index
        if any(line.strip() for line in block):
            return None
    return None


def merge_split_tables(md: str) -> tuple[str, list[Finding]]:
    """Склеить таблицу с её продолжением и строку с её хвостом (ПРАВКА #69)."""
    findings: list[Finding] = []
    out: list[list] = []
    for kind, block in _segments(md.split("\n")):
        if kind != "table":
            out.append([kind, block])
            continue
        previous = _previous_table(out)
        merged = None if previous is None else _merge_pair(out[previous][1], block, findings)
        if merged is None:
            out.append([kind, block])
        else:
            out[previous][1] = merged
    for entry in out:                            # ПРАВКА #69: хвосты строк внутри таблицы
        if entry[0] != "table":
            continue
        rows, separator = _rows_of(entry[1])
        tails_merged = _merge_row_tails(rows, findings)
        if tails_merged is not None:
            entry[1] = _render_rows(tails_merged, separator)
    return "\n".join(line for _, block in out for line in block), findings


# --- 3-6. точечные замены ---------------------------------------------------

def fix_numero(md: str) -> str:
    """No / No. перед цифрой или п/п → «№ »; Nokia, Note, ПNo1 не трогаем."""
    return _NUMERO_RE.sub("№ ", md)


def fix_degree(md: str) -> str:
    r"""$C^{\circ}$ и $^{\circ}C$ → °C; прочие формулы не трогаем."""
    return _DEGREE_TAIL_RE.sub("°C", _DEGREE_RE.sub(" °C", md))


def fix_list_glue(md: str) -> str:
    """ПРАВКА #68: «на:-» → «на: -». Только «;-» и «:-», разделитель «:---» не трогаем."""
    return _LIST_GLUE_RE.sub(" -", md)


def _canonical(folded: str | None, allow_ip_code: bool) -> str | None:
    if folded is None:
        return None
    upper = folded.upper()
    if upper in _KNOWN_BY_UPPER:
        return _KNOWN_BY_UPPER[upper]
    if allow_ip_code and _IP_CODE_RE.fullmatch(upper):
        return upper
    return None


def fix_mixed_alphabet(md: str) -> tuple[str, list[Finding]]:
    """Смешанный алфавит: известное чиним по словарю, остальное уходит в находки."""
    findings: list[Finding] = []

    def replace(match: re.Match) -> str:
        token = match.group()
        has_latin = any(_is_latin(char) for char in token)
        has_cyrillic = any(_is_cyrillic(char) for char in token)
        folded = fold_to_latin(token)
        if has_latin and has_cyrillic:
            canonical = _canonical(folded, allow_ip_code=True)
            if canonical is not None:
                return canonical
            suggestion = folded if folded is not None else fold_to_cyrillic(token)
            findings.append(_finding("mixed_alphabet_unknown", token, suggestion))
            return token
        if has_cyrillic and not has_latin and len(token) >= 3:
            canonical = _canonical(folded, allow_ip_code=False)
            if canonical is not None:
                return canonical
        return token

    return _TOKEN_RE.sub(replace, md), findings


# --- 7-8. только пометки ----------------------------------------------------

def flag_translit(md: str) -> list[Finding]:
    """Строка «инициалы + фамилия» с латиницей — пометить, текст не менять."""
    findings = []
    for line in md.split("\n"):
        if not NAME_RE.match(line) or not any(_is_latin(char) for char in line):
            continue
        snippet = line.strip()
        suggestion = fold_to_cyrillic(snippet)
        # ПРАВКА #70: «A.III.» свернулось бы в «А.ІІІ.» — украинское І в русских
        # инициалах заведомо не то, там разобранная на палки «Ш». Не гадаем.
        if suggestion is not None and ("І" in suggestion or "і" in suggestion):
            suggestion = None
        findings.append(_finding("translit_suspect", snippet, suggestion))
    return findings


def flag_signature_block(md: str) -> list[Finding]:
    """Три и более подряд абзаца «инициалы + фамилия» — разорванный блок подписей."""
    findings = []
    run: list[str] = []
    for block in re.split(r"\n[ \t]*\n", md) + [""]:
        if NAME_RE.match(block):
            run.append(block.strip())
            continue
        if len(run) >= 3:
            findings.append(_finding(
                "signature_block", run[0],
                "Должности и ФИО идут отдельными списками — сопоставить по скану"))
        run = []
    return findings


# --- 9. цепочка -------------------------------------------------------------

def postprocess(md: str) -> tuple[str, list[Finding]]:
    """1 → 2 → … → 8 → cleanup_ocr_markdown. Находки — в порядке получения."""
    findings: list[Finding] = []
    md, found = html_tables_to_pipe(md)
    findings += found
    md, found = merge_split_tables(md)
    findings += found
    md = fix_numero(md)
    md = fix_degree(md)
    md = fix_list_glue(md)                       # ПРАВКА #68
    md, found = fix_mixed_alphabet(md)
    findings += found
    findings += flag_translit(md)
    findings += flag_signature_block(md)
    return cleanup_ocr_markdown(md), findings
