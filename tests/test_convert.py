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
