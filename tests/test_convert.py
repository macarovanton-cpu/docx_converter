"""Тесты табличного рендерера convert.py (запускать из docx_converter/)."""

import pytest
from docx import Document
from docx.oxml.ns import qn

from convert import convert_md_to_docx

# ПРАВКА #34: иконка добавляется к тексту, а не вместо него
CASES = [
    ("Да — 36 месяцев",       "✓ Да — 36 месяцев"),
    ("Нет данных",            "✗ Нет данных"),
    ("Нет.",                  "✗ Нет."),
    ("Да",                    "✓ Да"),
    ("Отсутствует в базовой", "✗ Отсутствует в базовой"),
    ("Данные по объекту",     "Данные по объекту"),   # «Да» без \b — иконки нет
    ("Гарантия",              "Гарантия"),
]


@pytest.fixture(scope="module")
def table(tmp_path_factory):
    """Одна конвертация без шаблона на весь модуль → таблица из готового docx."""
    rows = "\n".join(f"| Параметр {i} | {src} |" for i, (src, _) in enumerate(CASES))
    md = f"| Параметр | Значение |\n|---|---|\n{rows}\n"
    out = tmp_path_factory.mktemp("docx") / "table.docx"
    convert_md_to_docx(md, str(out))
    return Document(str(out)).tables[0]


@pytest.fixture(scope="module")
def cells(table):
    """{исходный текст ячейки: то, что реально попало в word/document.xml}."""
    return {src: row.cells[1].text for (src, _), row in zip(CASES, table.rows[1:])}


def test_table_row_count(table):
    """Страховка: шапка + все строки на месте, иначе ассерты ниже вакуумны."""
    assert len(table.rows) == len(CASES) + 1


@pytest.mark.parametrize("src,expected", CASES)
def test_cell_text_survives_icon(cells, src, expected):
    assert cells[src] == expected


def test_header_repeats_and_rows_do_not_split(table):
    """ПРАВКА #35: tblHeader только на шапке, cantSplit на каждой строке."""
    headers = [i for i, r in enumerate(table.rows)
               if r._tr.find(qn('w:trPr')).find(qn('w:tblHeader')) is not None]
    assert headers == [0]
    assert all(r._tr.find(qn('w:trPr')).find(qn('w:cantSplit')) is not None
               for r in table.rows)


# =============================================================================
# ИНЛАЙН-ПАРСЕР
# =============================================================================

def _runs(md, tmp_path):
    """[(текст, bold, italic), ...] по всем абзацам конвертации без шаблона."""
    out = tmp_path / "inline.docx"
    convert_md_to_docx(md, str(out))
    return [(r.text, bool(r.bold), bool(r.italic))
            for p in Document(str(out)).paragraphs for r in p.runs]


def _text(md, tmp_path):
    return "".join(t for t, _, _ in _runs(md, tmp_path))


# ПРАВКА #36: звёздочки-умножение не съедаются как курсив
def test_multiplication_with_spaces_survives(tmp_path):
    runs = _runs("Габариты платформы 2 * 3 * 4 метра.", tmp_path)
    assert "2 * 3 * 4" in "".join(t for t, _, _ in runs)
    assert not any(italic for _, _, italic in runs)


def test_multiplication_without_spaces_survives(tmp_path):
    runs = _runs("Габариты 2*3*4 метра.", tmp_path)
    assert "2*3*4" in "".join(t for t, _, _ in runs)
    assert not any(italic for _, _, italic in runs)


def test_italic_still_works(tmp_path):
    assert ("курсив", False, True) in _runs("Обычный *курсив* тут.", tmp_path)


def test_bold_still_works(tmp_path):
    assert ("жирный", True, False) in _runs("И **жирный** тут.", tmp_path)


def test_bold_italic_still_works(tmp_path):
    assert ("оба", True, True) in _runs("И ***оба*** тут.", tmp_path)


def test_emphasis_with_inner_space_still_works(tmp_path):
    """Пробелы внутри эмфазы разрешены — правило касается только флангов."""
    assert ("две слова", False, True) in _runs("Тут *две слова* курсивом.",
                                               tmp_path)
