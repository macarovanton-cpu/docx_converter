"""Приёмочные тесты спеки 00: эталон golden.md, метрика vlm→golden, помощники.

Фикстуры лежат в _test/fixtures/ocr/ (папка в .gitignore) — без них тесты
пропускаются через require_fixture, а не падают.
"""

import hashlib
import os
import re
import subprocess
import sys

from ocr_fixtures import FIXTURES, count_diffs, read_fixture, require_fixture, text_tokens

REPO_DIR = FIXTURES.parents[2]

# Фактическое значение, снятое исполнителем спеки 00; расчётное в спеке — 14.
# ПРАВКА #68: было 12, стало 76 — правка 8 («;-» → «; -») разводит 66 склеек на
# два токена. 64 из них дают свой опкод, две попали внутрь соседних («°C;-» и
# «РоЕ;-»). ПРАВКА #72: 76 → 79 — правка 9 («Ethernet.Для», «ПО.Работы») даёт два
# опкода, правка 10 («(персональныйкомпьютер,») — один.
# ПРАВКА #75: 79 → 99 — правка 11 («материалов.2. » → «материалов. 2. ») даёт
# 19 опкодов, расширение правки 9 на цифру слева («Приложение 1.План») — один.
# Разбор — в docs/docx_converter_docs_sync.md, «Метрика vlm→golden».
VLM_TO_GOLDEN_DIFFS = 99
# golden.md вне git — sha ловит тихую подмену эталона.
GOLDEN_SHA256 = "f96c73ac39b3d2bdab4103963fa09305f36ed2eb8ed9191a68b71484740e1e59"


def test_metric_and_golden_unchanged():
    assert count_diffs(read_fixture("vlm.md"), read_fixture("golden.md")) == VLM_TO_GOLDEN_DIFFS
    golden = read_fixture("golden.md")
    assert count_diffs(golden, golden) == 0
    assert hashlib.sha256(require_fixture("golden.md").read_bytes()).hexdigest() == GOLDEN_SHA256


def test_fixes_1_to_5():
    vlm, golden = read_fixture("vlm.md"), read_fixture("golden.md")

    assert not [t for t in text_tokens(golden) if t.startswith("No")]
    assert golden.count("№ ") == 4 and "№ п/п" in golden and "№ 384-ФЗ" in golden
    assert "$" not in golden and "\\circ" not in golden
    assert "+50 °C;" in golden and "+35 °C." in golden
    assert "РоЕ" not in golden            # кириллица
    assert golden.count("PoE") == 2       # латиница
    assert "сыпучими" in golden and "сыручими" not in golden
    assert "IR-" not in golden
    assert golden.count("IP-камерами") == 1
    assert golden.count("IP-") == vlm.count("IP-") + 1      # остальные IP- не тронуты


def test_fix_6_signatures():
    golden = read_fixture("golden.md")
    for name in ("А.Ш. Таипов", "Р.Р. Нуреев", "А.Р. Сиражитдинов", "А.Ш. Ямалов", "Е.К. Кустова"):
        assert name in golden
    tail = golden[golden.index("СОГЛАСОВАНО:"):]
    assert not re.search(r"[A-Za-z]", tail)


def test_fix_7_single_pipe_table():
    golden = read_fixture("golden.md")

    assert "<table" not in golden and "<td" not in golden
    lines = golden.split("\n")
    idx = [i for i, l in enumerate(lines) if l.startswith("|")]
    assert idx == list(range(idx[0], idx[-1] + 1))            # один сплошной блок
    rows = [l for l in lines if l.startswith("|")]
    assert all(len(r.strip().strip("|").split("|")) == 3 for r in rows)
    assert re.fullmatch(r"\|\s*-+\s*\|\s*-+\s*\|\s*-+\s*\|", rows[1])
    first = [r.strip().strip("|").split("|")[0].strip() for r in rows[2:]]
    assert "" not in first                                     # номер пункта не потерян
    assert first.count("25") == 1 and first.count("26") == 1
    row25 = next(r for r in rows if r.lstrip("| ").startswith("25 "))
    assert "первичная поверка весов" in row25                   # хвост со стр. 6
    assert "ITV Интеллект - УРММ" in row25                      # хвост со стр. 7
    assert any(r.startswith("| I. Общие данные |") for r in rows)


def test_fix_8_list_glue():
    """ПРАВКА #68: правка 8 — «;»/«:» вплотную к маркеру списка разводятся пробелом."""
    vlm, golden = read_fixture("vlm.md"), read_fixture("golden.md")
    assert len(re.findall(r"[;:]-", vlm)) == 66                # склеек в исходнике
    assert not re.search(r"[;:]-", golden)                     # в эталоне ни одной
    # 5 своих у vlm + 66 разведённых правкой 8 + 2 стыка, где правка 7 дописала
    # продолжение таблицы через пробел к ячейке, кончавшейся на «;»
    assert len(re.findall(r"[;:] -", golden)) == len(re.findall(r"[;:] -", vlm)) + 66 + 2


def test_fix_9_sentence_glue():
    """ПРАВКА #72: правка 9 — точка между предложениями разводится пробелом.

    ПРАВКА #75б: слева от точки теперь и цифра — «Приложение 1.План» тоже
    разводится, из «осталось как было» этот случай ушёл.
    """
    vlm, golden = read_fixture("vlm.md"), read_fixture("golden.md")
    glue = r"(?:(?<=[^\W\d_]{2})|(?<=\d))\.(?=[А-ЯЁ])"
    assert len(re.findall(glue, vlm)) == 3     # «Ethernet.Для», «ПО.Работы», «1.План»
    assert not re.search(glue, golden)
    assert "RS-485,Ethernet. Для" in golden and "и ПО. Работы," in golden
    assert "Приложение 1. План" in golden                      # ПРАВКА #75б
    # соседние точки под правило не попали и остались слитными
    assert "в т.ч.дистрибутивы" in golden                      # справа строчная
    assert "компания».Юридический" in golden                   # слева кавычка


def test_fix_11_list_number_glue():
    """ПРАВКА #75а: правка 11 — номер пункта, прилипший к концу фразы."""
    vlm, golden = read_fixture("vlm.md"), read_fixture("golden.md")
    glue = r"[.;:](?=\d{1,2}\.\s)"
    assert len(re.findall(glue, vlm)) == 19                    # склеек в исходнике
    assert not re.search(glue, golden)                         # в эталоне ни одной
    assert "требования: 1. Платформа" in golden
    assert "ГОСТ 380-2005. 2. Размер" in golden                # слева год, не номер пункта
    # даты и номера стандартов под правило не попали: за точкой не «N. »
    for same in ("30.12.2019", "10.01.2002", "21.1101-2020"):
        assert same in golden


def test_fix_10_row_tail_space():
    """ПРАВКА #72: правка 10 — хвост строки 13 пристыкован через пробел, как склеивает #69."""
    vlm, golden = read_fixture("vlm.md"), read_fixture("golden.md")
    assert "(персональныйкомпьютер," in vlm                    # в исходнике слитно
    assert golden.count("(персональный компьютер,") == 1
    assert "(персональныйкомпьютер," not in golden
    # «Программно-техническийкомплекс» склеен в самом vlm.md — не трогаем
    assert "Программно-техническийкомплекс" in golden


def test_live_marker_needs_explicit_flag():
    """ПРАВКА #71: ключ в окружении сеть не включает — включает только MINERU_LIVE=1."""
    env = dict(os.environ, MINERU_API_KEY="k-не-используется")
    env.pop("MINERU_LIVE", None)
    run = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "pytest", "-m", "live", "-q", "-rs",
         "-p", "no:cacheprovider"],
        cwd=REPO_DIR, env=env, capture_output=True, text=True, encoding="utf-8")
    assert run.returncode == 0, run.stdout                 # сеть не тронута
    assert "skipped" in run.stdout and "MINERU_LIVE=1" in run.stdout


def test_text_tokens_strips_markup():
    assert text_tokens("<td>a</td><td>b|c</td>\n|---|---|\n| d |") == ["a", "b", "c", "d"]
