"""ПРАВКА #63: приёмочные тесты детерминированного постпроцессора.

Фикстурные тесты пропускаются через require_fixture, если нет _test/fixtures/ocr/.
Сети и кэша спека не требует.
"""

import zipfile

from ocr import SEVERITIES
from ocr_fixtures import count_diffs, read_fixture, require_fixture, text_tokens
from ocr.postprocess import (fix_degree, fix_list_glue, fix_mixed_alphabet,
                             fix_numero, flag_signature_block, flag_translit,
                             html_tables_to_pipe, merge_split_tables,
                             parse_pipe_tables, postprocess)
from ocr.validate import validate
from test_ocr_fixtures import VLM_TO_GOLDEN_DIFFS

# Фактический остаток на vlm.md: 5 опкодов — «сыручими», «IR-камерами» и три
# опкода на пяти подписях. Спека считала 7 из расчётных 14 в спеке 00.
# ПРАВКА #68: порог опущен с 7 до факта — запас в две единицы пропускал бы
# регрессию на один опкод. Склейки списка (правка 8) в остаток не попадают:
# fix_list_glue чинит их в тракте ровно так же, как они починены в эталоне.
REMAINING_DIFFS = 5


def run_vlm():
    out, findings = postprocess(read_fixture("vlm.md"))
    return out, findings, [f.rule for f in findings]


def test_metric_and_idempotence():
    out, _, _ = run_vlm()
    golden = read_fixture("golden.md")
    assert count_diffs(out, golden) <= VLM_TO_GOLDEN_DIFFS
    assert count_diffs(out, golden) <= REMAINING_DIFFS
    assert postprocess(out)[0] == out


def test_fixed_artifacts():
    out, _, _ = run_vlm()
    assert not [t for t in text_tokens(out) if t.startswith("No")]
    assert out.count("№ ") == 4 and "№ п/п" in out and "№ 384-ФЗ" in out and "№ 7-ФЗ" in out
    assert "$" not in out and "+50 °C;" in out and "+35 °C." in out
    assert "РоЕ" not in out and out.count("PoE") == 2


def test_not_fixed_without_scan():
    out, _, _ = run_vlm()
    assert "сыручими" in out and "сыпучими" not in out
    assert "IR-камерами" in out
    assert "A.II. Taipov" in out and "P.P. Hypeeb" in out


def test_single_merged_table():
    out, _, _ = run_vlm()
    assert "<table" not in out and "<td" not in out
    tables = parse_pipe_tables(out)
    assert len(tables) == 1
    assert len(tables[0]) == 53                                 # 55 <tr> минус 2 строки-продолжения
    assert all(len(row) == 3 for row in tables[0])
    assert tables[0][0][0] == "№ п/п"
    assert tables[0][1] == ["I. Общие данные", "", ""]
    assert "" not in [row[0] for row in tables[0]]
    row25 = next(r for r in tables[0] if r[0] == "25")
    assert "пуско-наладочные работы; - первичная поверка" in row25[2]
    assert row25[2].rstrip().endswith("до сети ИТСО.")
    assert tables[0][-1][0] == "43"


def test_findings():
    out, findings, rules = run_vlm()
    assert rules.count("table_merged") == 2
    assert rules.count("table_span") == 1
    assert rules.count("signature_block") == 1
    assert rules.count("translit_suspect") == 5
    # единственный смешанный токен vlm — склейка «bzdk@list.ruОГРН»; «Noп» уже снят шагом 3
    assert [f.snippet for f in findings if f.rule == "mixed_alphabet_unknown"] == ["ruОГРН"]
    assert "html_table_unparsed" not in rules and "table_merge_failed" not in rules
    assert all(f.page is None for f in findings)
    assert all(f.severity in SEVERITIES for f in findings)
    tr = {f.snippet: f for f in findings if f.rule == "translit_suspect"}
    assert tr["A.P. Сиражитдинов"].suggestion == "А.Р. Сиражитдинов"
    assert tr["E.K. Кустова"].suggestion == "Е.К. Кустова"
    assert tr["P.P. Hypeeb"].suggestion is None                 # 'b', 'y' без пары — не гадать
    assert all(f.severity == "critical" for f in tr.values())
    for f in findings:
        assert f.snippet in out or f.rule in ("table_merged", "table_span")


def test_row_tail_merged_on_live_raw():
    """ПРАВКА #69: в живом прогоне пункт 13 приехал двумя строками — стал одной."""
    with zipfile.ZipFile(require_fixture("vlm_raw.zip")) as archive:
        raw = archive.read("full.md").decode("utf-8")
    out, findings = postprocess(raw)
    table = parse_pipe_tables(out)[0]

    rows13 = [row for row in table if row[0] == "13"]
    assert len(rows13) == 1                                     # было две строки
    assert "(персональный компьютер," in rows13[0][1]           # хвост первой ячейки
    assert "Ethernet. Для передачи данных" in rows13[0][2]      # хвост третьей
    assert "" not in [row[0] for row in table]
    assert [f.rule for f in findings].count("table_merged") == 3
    assert "table_empty_number" not in [f.rule for f in validate(out)]


def test_noisy_pipeline_run():
    pipeline = read_fixture("pipeline.md")
    out, findings = postprocess(pipeline)
    assert "mixed_alphabet_unknown" in [f.rule for f in findings]
    assert len(text_tokens(out)) >= len(text_tokens(pipeline)) - 5


# --- синтетика: по одному тесту на функцию ----------------------------------

def test_html_tables_to_pipe():
    md, f = html_tables_to_pipe(
        '<table><tr><td rowspan="2">A</td><td>1</td></tr><tr><td>2</td></tr></table>')
    assert parse_pipe_tables(md) == [[["A", "1"], ["A", "2"]]] and f[0].rule == "table_span"

    md, f = html_tables_to_pipe("<table><tr><td>a|b</td><td>x<br>y &amp; z</td></tr></table>")
    assert "| a\\|b | x y & z |" in md

    bad = "<table><tr><td><table><tr><td>in</td></tr></table></td></tr></table>"
    md, f = html_tables_to_pipe(bad)
    assert md == bad and [x.rule for x in f] == ["html_table_unparsed"]

    md, f = html_tables_to_pipe("<table><tr><td>1</td><td>2</td></tr><tr><td>3</td></tr></table>")
    assert [len(r) for r in parse_pipe_tables(md)[0]] == [2, 1]     # не дополняем


def test_merge_split_tables():
    a = "| № | Что |\n|---|---|\n| 1 | начало |"
    assert parse_pipe_tables(merge_split_tables(a + "\n\n|  | конец |\n|---|---|")[0]) == \
        [[["№", "Что"], ["1", "начало конец"]]]

    matrix = a + "\n\n|  | 2024 | 2025 |\n|---|---|---|\n| план | 1 | 2 |"
    assert merge_split_tables(matrix) == (matrix, [])               # шапка матрицы — не продолжение

    wide = a + "\n\n|  | хвост |\n|---|---|\n| 2 | x | лишняя |"
    md, f = merge_split_tables(wide)
    assert md == wide and [x.rule for x in f] == ["table_merge_failed"]

    assert merge_split_tables(a + "\n\nАбзац.\n\n|  | конец |\n|---|---|")[1] == []


def test_merge_row_tails():
    """ПРАВКА #69: строка с пустым номером и ≥2 непустыми — хвост предыдущей."""
    head = "| № | A | B |\n|---|---|---|\n"
    md, f = merge_split_tables(head + "| 1 | a | b |\n|  | хвост | ещё |")
    assert parse_pipe_tables(md) == [[["№", "A", "B"], ["1", "a хвост", "b ещё"]]]
    assert [x.rule for x in f] == ["table_merged"]

    one = head + "| 1 | a | b |\n| I. Общие данные |  |  |\n|  |  | одна |"
    assert merge_split_tables(one) == (one, [])          # одна непустая — не хвост
    first = "|  | x | y |\n|---|---|---|\n| 1 | a | b |"
    assert merge_split_tables(first) == (first, [])      # первой строке некуда дописывать
    wide = head + "| 1 | a | b |\n|  | x | y | лишняя |"
    assert merge_split_tables(wide) == (wide, [])        # ширина не та — ячейку не выбросим


def test_fix_list_glue():
    """ПРАВКА #68: только «;-» и «:-»; «+-», «--» и разделитель «:---» не трогаем."""
    assert fix_list_glue("на:- один;- два") == "на: - один; - два"
    assert fix_list_glue("+-30кг, 60 т. +- 50кг") == "+-30кг, 60 т. +- 50кг"
    assert fix_list_glue("|:---|---:|") == "|:---|---:|"
    assert fix_list_glue(fix_list_glue("на:-")) == fix_list_glue("на:-")


def test_fix_numero_and_degree():
    assert fix_numero("No1, No 12, No.7, Noп/п") == "№ 1, № 12, № 7, № п/п"
    assert fix_numero("Nokia Note ПNo1") == "Nokia Note ПNo1"
    assert fix_degree("от +5 до +35  $C^{\\circ}$ .") == "от +5 до +35 °C."
    assert fix_degree("$x^2$") == "$x^2$"


def test_fix_mixed_alphabet():
    assert fix_mixed_alphabet("функции РоЕ; IР68; 12В; Ст3")[0] == "функции PoE; IP68; 12В; Ст3"
    md, f = fix_mixed_alphabet("ЦСмц и Sм")
    assert md == "ЦСмц и Sм" and [x.snippet for x in f] == ["Sм"]
    assert fix_mixed_alphabet("РОЕ")[0] == "PoE" and fix_mixed_alphabet("Ре")[0] == "Ре"


def test_signature_flags():
    assert flag_translit("Текст A.P. Сиражитдинов в строке") == []   # не отдельной строкой
    assert flag_translit("А.Р. Сиражитдинов") == []                  # латиницы нет
    assert flag_signature_block("А.А. Иванов\n\nБ.Б. Петров") == []  # двух мало
    assert len(flag_signature_block("А.А. Иванов\n\nБ.Б. Петров\n\nВ.В. Сидоров")) == 1
