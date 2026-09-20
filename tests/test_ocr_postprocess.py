"""ПРАВКА #63: приёмочные тесты детерминированного постпроцессора.

Фикстурные тесты пропускаются через require_fixture, если нет _test/fixtures/ocr/.
Сети и кэша спека не требует.
"""

import re
import zipfile

from ocr import SEVERITIES
from ocr_fixtures import (count_diffs, read_fixture, read_raw, require_fixture,
                          text_tokens)
from ocr.postprocess import (fix_degree, fix_list_glue, fix_list_number_glue,
                             fix_mixed_alphabet, fix_numero, fix_sentence_glue,
                             flag_signature_block, flag_translit, html_tables_to_pipe,
                             merge_split_tables, parse_pipe_tables, postprocess,
                             recover_dropped_blocks)
from ocr.validate import validate
from test_ocr_fixtures import VLM_TO_GOLDEN_DIFFS

# Фактический остаток на vlm.md: 5 опкодов — «сыручими», «IR-камерами» и три
# опкода на пяти подписях. Спека считала 7 из расчётных 14 в спеке 00.
# ПРАВКА #68: порог опущен с 7 до факта — запас в две единицы пропускал бы
# регрессию на один опкод. Склейки списка (правка 8) в остаток не попадают:
# fix_list_glue чинит их в тракте ровно так же, как они починены в эталоне.
# ПРАВКА #72: 5 → 6. Правка 9 в остаток не попадает (fix_sentence_glue чинит её
# в тракте), а правка 10 попадает: «(персональныйкомпьютер,» пришёл из MinerU
# уже слитным словом, разделить его без скана нельзя. В эталоне пробел стоит —
# так эту же строку склеивает #69 на сыром прогоне, где ячейки ещё раздельны.
REMAINING_DIFFS = 6


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
    # ПРАВКА #72: точка между предложениями разведена, соседние случаи — нет
    assert "RS-485,Ethernet. Для" in out and "и ПО. Работы," in out
    assert "в т.ч.дистрибутивы" in out and "Приложение 1. План" in out


def test_not_fixed_without_scan():
    out, _, _ = run_vlm()
    assert "сыручими" in out and "сыпучими" not in out
    assert "IR-камерами" in out
    assert "A.II. Taipov" in out and "P.P. Hypeeb" in out
    # ПРАВКА #72: MinerU прислал слово слитным — тракт слова не режет (эталонный
    # «(персональный компьютер,» достижим только на сыром прогоне, через #69)
    assert "(персональныйкомпьютер," in out


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
    # ПРАВКА #70: «II»/«III» свернулись бы в украинское «І» — это не инициалы
    assert tr["A.II. Taipov"].suggestion is None
    assert tr["A.III. Ямалов"].suggestion is None
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


# --- обе фикстуры, сырые прогоны (ПРАВКА #74, #75, #76) ----------------------

def run_raw(name):
    """postprocess сырого прогона с content_list и без него."""
    markdown, content_list = read_raw(name)
    return postprocess(markdown, content_list), postprocess(markdown)


def test_recovered_blocks_bakeoff2():
    """ПРАВКА #74: «СОГЛАСОВАНО … Д.В. Майер» и «УТВЕРЖДАЮ …» со стр. 1 вернулись."""
    (out, findings), (plain, plain_findings) = run_raw("vlm_raw2.zip")

    for lost in ("СОГЛАСОВАНО", "Д.В. Майер", "УТВЕРЖДАЮ", "К.В. Древко",
                 "ООО «Мангазея Майнинг»"):
        assert lost not in plain and lost in out
    recovered = [f for f in findings if f.rule == "recovered_block"]
    assert len(recovered) == 1                       # три блока подряд — одна врезка
    assert recovered[0].severity == "info" and recovered[0].snippet == "СОГЛАСОВАНО"
    assert "header" in recovered[0].suggestion
    # порядок блоков сохранён, подписант остался при своём блоке
    assert out.index("СОГЛАСОВАНО") < out.index("Д.В. Майер") < out.index("УТВЕРЖДАЮ")
    # ПРАВКА #78: врезка встала перед первым блоком своей страницы, а не за
    # таблицей, растянутой на весь документ
    assert out.startswith("СОГЛАСОВАНО")
    assert (out.index("УТВЕРЖДАЮ") < out.index("«23» 04 2026 г.")
            < out.index("## ТЕХНИЧЕСКОЕ ЗАДАНИЕ"))
    # кроме врезки постпроцессор не изменился: те же правила в том же порядке
    assert [f.rule for f in findings if f.rule != "recovered_block"] == \
        [f.rule for f in plain_findings]


def test_recovered_blocks_bakeoff():
    """ПРАВКА #74: на bakeoff возвращается один колонтитул, номера страниц — нет."""
    (out, findings), (plain, plain_findings) = run_raw("vlm_raw.zip")

    recovered = [f for f in findings if f.rule == "recovered_block"]
    assert len(recovered) == 1 and recovered[0].snippet == "с. Сафарово, 2026 г."
    assert "с. Сафарово, 2026 г." not in plain and "с. Сафарово, 2026 г." in out
    # ПРАВКА #78: нижний колонтитул остался на своём месте — под титулом, не над ним
    assert out.startswith("Утверждаю:")
    assert out.index("## ТЕХНИЧЕСКОЕ ЗАДАНИЕ") < out.index("с. Сафарово, 2026 г.")
    for page_number in ("Лист 1 из 9", "Лист 5 из 9", "Лист 9 из 9"):
        assert page_number not in out                # номера страниц не возвращаем
    assert [f.rule for f in findings if f.rule != "recovered_block"] == \
        [f.rule for f in plain_findings]
    # одиночных цифр отдельными абзацами тоже не прибавилось
    assert [b for b in out.split("\n\n") if b.strip().isdigit()] == []


def test_recovery_needs_content_list():
    """ПРАВКА #74: без content_list шаг 0 молчит — так тракт идёт по md-фикстурам."""
    markdown, _ = read_raw("vlm_raw2.zip")
    assert recover_dropped_blocks(markdown, None) == (markdown, [])
    assert recover_dropped_blocks(markdown, []) == (markdown, [])
    out, findings = postprocess(read_fixture("vlm.md"))
    assert "recovered_block" not in [f.rule for f in findings]


def test_glue_fixes_bakeoff2():
    """ПРАВКА #75: склейки а/б/в на втором документе."""
    (out, _), _ = run_raw("vlm_raw2.zip")

    assert "СП 20.13330.2016. Климатический" in out            # б: слева цифра
    assert "из своих материалов. 2. Подрядчик" in out          # а: номер пункта
    assert "Заказчик: 1. Передает" in out
    assert "форма № КС-2" in out and "форме № КС-3. Заказчик" in out    # в + б
    assert "No " not in out and "20x4,5x5,0м" in out
    assert not re.search(r"[.;:]\d{1,2}\.\s", out)
    assert postprocess(out)[0] == out                           # идемпотентно


def test_cell_tail_merged_bakeoff2():
    """ПРАВКА #76: продолжение пункта 2.1 — текст только в последней ячейке."""
    (out, findings), _ = run_raw("vlm_raw2.zip")
    table = parse_pipe_tables(out)[0]

    rows = [row for row in table if row[0] == "2.1"]
    assert len(rows) == 1                                      # было две строки
    assert "-Монтаж навеса под автовесовую" in rows[0][2]      # хвост дописан в конец
    assert rows[0][2].index("Разработка рабочей документации") < \
        rows[0][2].index("-Монтаж навеса")
    assert "" not in [row[0] for row in table]
    assert "table_empty_number" not in [f.rule for f in validate(out)]
    merged = [f for f in findings if f.rule == "table_merged"]
    assert len(merged) == 1 and "«2.1»" in merged[0].suggestion


def test_cell_tail_leaves_bakeoff_alone():
    """ПРАВКА #76: на bakeoff новая ветка ничего не склеила сверх прежних трёх."""
    (out, findings), _ = run_raw("vlm_raw.zip")
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


def test_fix_sentence_glue():
    """ПРАВКА #72: точка между предложениями; инициалы и версии — мимо."""
    assert fix_sentence_glue("проектом.Предусмотреть") == "проектом. Предусмотреть"
    assert fix_sentence_glue("Ethernet.Для и ПО.Работы") == "Ethernet. Для и ПО. Работы"
    # ПРАВКА #75б: слева цифра — тоже склейка
    assert fix_sentence_glue("СП 20.13330.2016.Климатический") == "СП 20.13330.2016. Климатический"
    assert fix_sentence_glue("1.3.4.Требования") == "1.3.4. Требования"
    assert fix_sentence_glue("Приложение 1.План") == "Приложение 1. План"
    for same in ("И.М. Халиуллин", "т.е.Х", "Windows 8.1", "СП 131.13330.2020",
                 "в т.ч.дистрибутивы", "компания».Юридический",
                 "A.II. Taipov", "А.Ш. Таипов"):
        assert fix_sentence_glue(same) == same
    assert fix_sentence_glue(fix_sentence_glue("ПО.Работы")) == fix_sentence_glue("ПО.Работы")


def test_fix_list_number_glue():
    """ПРАВКА #75а: номер пункта, прилипший к концу фразы."""
    assert fix_list_number_glue("материалов.2. Подрядчик") == "материалов. 2. Подрядчик"
    assert fix_list_number_glue("Заказчик:1. Передает") == "Заказчик: 1. Передает"
    assert fix_list_number_glue("передачи;2. Предоставляет") == "передачи; 2. Предоставляет"
    assert fix_list_number_glue("ГОСТ 380-2005.2. Размер") == "ГОСТ 380-2005. 2. Размер"
    for same in ("СП 131.13330.2020, СП", "п. 2.1, требованиями", "20x4,5x5,0м",
                 "версия 1.2.3 сборки", "итого:100. Всего"):
        assert fix_list_number_glue(same) == same
    once = fix_list_number_glue("материалов.2. Подрядчик")
    assert fix_list_number_glue(once) == once


def test_fix_numero_and_degree():
    assert fix_numero("No1, No 12, No.7, Noп/п") == "№ 1, № 12, № 7, № п/п"
    assert fix_numero("Nokia Note ПNo1") == "Nokia Note ПNo1"
    # ПРАВКА #75в: заглавная кириллица справа — тоже номер
    assert fix_numero("форма No КС-2 и No КС-3") == "форма № КС-2 и № КС-3"
    assert fix_numero("No problem, Note") == "No problem, Note"
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
