"""Приёмочные тесты спеки 00: эталон golden.md, метрика vlm→golden, помощники.

Фикстуры лежат в _test/fixtures/ocr/ (папка в .gitignore) — без них тесты
пропускаются через require_fixture, а не падают.
"""

import hashlib
import re

from ocr_fixtures import count_diffs, read_fixture, require_fixture, text_tokens

# Фактическое значение, снятое исполнителем спеки 00; расчётное в спеке — 14.
# Разница объяснена в docs/docx_converter_docs_sync.md, раздел «Метрика vlm→golden».
VLM_TO_GOLDEN_DIFFS = 12
# golden.md вне git — sha ловит тихую подмену эталона.
GOLDEN_SHA256 = "dc3f377c0d54210a9bacbc8369aa19ee866c13cd7e094f2a94f3963f16bbd887"


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


def test_text_tokens_strips_markup():
    assert text_tokens("<td>a</td><td>b|c</td>\n|---|---|\n| d |") == ["a", "b", "c", "d"]
