"""ПРАВКА #64: приёмочные тесты валидатора, report.json и пометок.

Фикстурные тесты пропускаются через require_fixture, если нет _test/fixtures/ocr/.
Сети спека не требует.
"""

import json
import zipfile

from ocr_fixtures import read_fixture, read_raw, require_fixture
from ocr.postprocess import html_tables_to_pipe, parse_pipe_tables, postprocess
from ocr.validate import (ANNOTATION_PREFIX, LOST_HEADING, annotate, build_report,
                          page_of, strip_annotations, validate)

# правила 05, которые на чистом vlm молчать обязаны
QUIET_ON_VLM = ("ogrn_checksum", "inn_checksum", "gost_format", "unit_unknown",
                "table_row_cells", "table_empty_number", "table_total_mismatch")


def run_vlm():
    md, post = postprocess(read_fixture("vlm.md"))
    return md, post, validate(md)


def test_quiet_on_clean_markdown():
    md, _, val = run_vlm()
    assert "ОГРН 1020202283287" in md          # 102020228328 % 11 % 10 == 7
    assert [f for f in val if f.rule in QUIET_ON_VLM] == []
    assert validate(read_fixture("golden.md")) == []


def test_lost_row_number_without_merge():
    unmerged = html_tables_to_pipe(read_fixture("vlm.md"))[0]
    rules = [f.rule for f in validate(unmerged)]
    # первая строка табл. 3: '', '', текст; табл. 2 (одна строка '', текст)
    # под правило не попадает — первая колонка не нумерационная
    assert rules.count("table_empty_number") == 1


def test_noisy_pipeline_run():
    pv = validate(postprocess(read_fixture("pipeline.md"))[0])
    # «ГОст», «ГОСт», «гОСт», «ГОСТ Р53228»
    assert [f.rule for f in pv].count("gost_format") >= 3
    assert any(f.rule == "unit_unknown" and f.suggestion == "мм" for f in pv)   # «10MM.»


def test_report_schema():
    _, post, val = run_vlm()
    findings = post + val
    report = build_report(source="bakeoff.pdf", sha256="ab" * 32, provider="mineru",
                          model_version="vlm", cache_hit=False, verified=False,
                          findings=findings)
    assert list(report) == ["schema_version", "source", "sha256", "provider",
                            "model_version", "cache_hit", "verified", "created_at",
                            "summary", "findings"]
    assert report["schema_version"] == 1
    assert [f["id"] for f in report["findings"]] == list(range(1, len(findings) + 1))
    assert all(list(f) == ["id", "rule", "severity", "page", "snippet", "suggestion"]
               for f in report["findings"])
    assert report["summary"]["critical"] == 5                  # пять ФИО с латиницей
    assert sum(report["summary"].values()) == len(report["findings"])
    assert set(report["summary"]) == {"critical", "warning", "info"}
    assert json.loads(json.dumps(report, ensure_ascii=False)) == report


def test_pages_from_content_list():
    cl = [{"type": "text", "text": "P.P. Hypeeb", "page_idx": 8},
          {"type": "table", "table_body": "<table><tr><td>сыручими  и</td></tr></table>",
           "page_idx": 1}]
    assert page_of("P.P. Hypeeb", cl) == 9 and page_of("сыручими и", cl) == 2
    assert page_of("нет такого", cl) is None and page_of("x", None) is None
    assert page_of("P.P. Hypeeb", [{"type": "text", "text": "P.P. Hypeeb"}]) is None

    _, post, _ = run_vlm()
    r2 = build_report(source="s", sha256="0" * 64, provider="mineru", model_version="vlm",
                      cache_hit=True, verified=False, findings=post, content_list=cl)
    assert next(f for f in r2["findings"] if f["snippet"] == "P.P. Hypeeb")["page"] == 9


def test_pages_from_real_content_list():
    """PLACEHOLDER 3 спеки 05 закрыт: page_of проверен на настоящем content_list.json.

    Архив живого прогона vlm кладёт файл под именем «<uuid>_content_list.json» —
    так его и ищет result_from_zip. Форма блока совпала с докой: type, page_idx
    (с нуля) и взаимоисключающие text / table_body; сверх них — bbox, img_path,
    text_level, table_caption, table_footnote.
    """
    with zipfile.ZipFile(require_fixture("vlm_raw.zip")) as archive:
        names = [n for n in archive.namelist() if n.endswith("content_list.json")]
        assert len(names) == 1, names
        content_list = json.loads(archive.read(names[0]).decode("utf-8"))

    assert all("page_idx" in block and "type" in block for block in content_list)
    assert page_of("сыручими", content_list) == 2          # абзац стр. 2
    assert page_of("A.II. Taipov", content_list) == 9      # блок подписей, стр. 9
    assert page_of("такого в документе нет", content_list) is None


def test_annotations():
    md, post, val = run_vlm()
    report = build_report(source="bakeoff.pdf", sha256="ab" * 32, provider="mineru",
                          model_version="vlm", cache_hit=False, verified=False,
                          findings=post + val)
    ann = annotate(md, report)
    assert strip_annotations(ann) == md
    assert ann.count(ANNOTATION_PREFIX) == len(report["findings"])
    assert parse_pipe_tables(ann) == parse_pipe_tables(md)      # таблица не разорвана
    block = next(b for b in ann.split("\n\n") if "translit_suspect" in b and "Hypeeb" in b)
    assert block.startswith("!!") and block.endswith("!!") and "\n" not in block
    assert ann.split("\n\n").index(block) == ann.split("\n\n").index("P.P. Hypeeb") + 1


def _low(id_, snippet, suggestion=None):
    return {"id": id_, "rule": "low_confidence", "severity": "warning", "page": None,
            "snippet": snippet, "suggestion": suggestion}


def test_annotations_synthetic():
    assert strip_annotations("!! E = mc2 !!\n\nТекст") == "!! E = mc2 !!\n\nТекст"

    all_ = dict(include_low_confidence=True)
    lost = annotate("Абзац.", {"findings": [_low(1, "нет в тексте")]}, **all_)
    assert lost.endswith("!! ПРОВЕРИТЬ: [1] low_confidence: нет в тексте !!")
    assert strip_annotations(lost) == "Абзац."

    long_text = "я" * 300
    one = annotate("Абзац.", {"findings": [_low(2, "Абзац.", "а\nб")]}, **all_)
    assert one == "Абзац.\n\n!! ПРОВЕРИТЬ: [2] low_confidence: а б !!"
    cut = annotate("Абзац.", {"findings": [_low(3, long_text)]}, **all_)
    assert cut.endswith("я…" + " !!") and len(cut.split(": ", 2)[2]) == 200 + len(" !!")


def test_low_confidence_skipped_by_default():
    """ПРАВКА #70: в текст low_confidence попадают только по --annotate-all."""
    report = {"findings": [_low(1, "Абзац."),
                           {"id": 2, "rule": "translit_suspect", "severity": "critical",
                            "page": None, "snippet": "Абзац.", "suggestion": None}]}
    assert annotate("Абзац.", report) == "Абзац.\n\n!! ПРОВЕРИТЬ: [2] translit_suspect: Абзац. !!"
    assert annotate("Абзац.", report, include_low_confidence=True).count(ANNOTATION_PREFIX) == 2


def test_lost_findings_go_to_the_tail():
    """ПРАВКА #70: не найденный якорь — в конец, своим разделом, а не в шапку."""
    md = "# Шапка\n\nТело."
    report = {"findings": [_low(1, "нет в тексте"), _low(2, "")]}
    ann = annotate(md, report, include_low_confidence=True)
    assert ann.startswith(md)                            # шапку и тело не тронули
    blocks = ann.split("\n\n")
    assert blocks[2] == LOST_HEADING and len(blocks) == 5
    assert strip_annotations(ann) == md                  # заголовок раздела снимается вместе
    # свой такой же заголовок в документе — не наш, не трогаем
    own = "## Не привязанные находки\n\nТекст."
    assert strip_annotations(own) == own


# --- синтетика: по тесту на правило -----------------------------------------

def rules(text):
    return [f.rule for f in validate(text)]


def test_rule_inn():
    assert rules("ИНН 7707083893") == [] and rules("ИНН 7707083894") == ["inn_checksum"]
    assert rules("ИНН: 500100732259") == [] and rules("ИНН 77070838") == ["inn_checksum"]


def test_rule_ogrn():
    assert rules("ОГРН 1027700132195") == [] and rules("ОГРН 1027700132196") == ["ogrn_checksum"]
    assert rules("ОГРНИП 304500116000157") == []


def test_rule_kpp():
    assert rules("КПП 773601001") == [] and rules("КПП 77360100") == ["kpp_format"]


def test_rule_gost():
    assert rules("ГОСТ Р 53228-2008, ГОСТ 8.726-2010") == []
    assert rules("ГОСт 380-2005") == ["gost_format"] and rules("ГОСТ 380") == ["gost_format"]
    assert rules("(ТУ) на подключение") == [] and rules("ТУ 4274-001-12345678-2015") == []


def test_rule_gost_prose_and_dashes():
    """ПРАВКА #77: упоминание без номера — не находка; тире бывает трёх видов."""
    for prose in ("ссылки на ГОСТ, ТУ), подписью", "сертификат (ГОСТ Р), а также",
                  "по ГОСТ, рабочей документации", "требования ГОСТ."):
        assert rules(prose) == []
    for dash in ("-", "–", "—"):
        assert rules(f"ГОСТ Р 58760{dash}2019") == []
        assert rules(f"ТУ 4274-001-12345678{dash}2015".replace("-", dash)) == []
    # номер есть, но не по шаблону — находка осталась
    assert rules("ГОСТ Р53228") == ["gost_format"]
    # гомоглиф ловится и в прозе: это про написание слова, не про номер
    assert rules("по ГОСт, без номера") == ["gost_format"]


def test_gost_quiet_on_both_fixtures():
    """ПРАВКА #77: на обеих фикстурах gost_format молчит."""
    for name in ("vlm_raw.zip", "vlm_raw2.zip"):
        markdown, content_list = read_raw(name)
        md, _ = postprocess(markdown, content_list)
        assert [f.rule for f in validate(md)].count("gost_format") == 0, name
    # в bakeoff2 обозначение с em-dash и упоминания в прозе — всё это было находками
    md = postprocess(*read_raw("vlm_raw2.zip"))[0]
    assert "ГОСТ Р 58760—2019" in md and "ГОСТ, ТУ)" in md and "(ГОСТ Р)" in md


def test_rule_code_digits_glued():
    """ПРАВКА #79: к году обозначения прилип номер следующего пункта."""
    assert rules("по СП 76.13330.20163. Подрядчик") == ["code_digits_glued"]
    for clean in ("СП 131.13330.2020", "СП 20.13330.2016", "ГОСТ Р 58760—2019",
                  "ГОСТ 8.726-2010", "ТУ 4274-001-12345678-2015", "СНиП 3.05.06-85",
                  "ГОСТ Р 53228", "СП 13330"):     # одна группа — номер, не год
        assert "code_digits_glued" not in rules(clean), clean
    found = validate("по СП 76.13330.20163. Подрядчик")[0]
    assert found.snippet == "СП 76.13330.20163" and found.severity == "warning"
    assert found.suggestion == "проверить границу с следующим пунктом"


def test_code_digits_glued_on_fixtures():
    """ПРАВКА #79: на bakeoff2 ровно одна находка, на bakeoff — ни одной."""
    for name, expected in (("vlm_raw2.zip", ["СП 76.13330.20163"]), ("vlm_raw.zip", [])):
        md = postprocess(*read_raw(name))[0]
        assert [f.snippet for f in validate(md)
                if f.rule == "code_digits_glued"] == expected, name


def test_rule_units_and_models():
    assert rules("не менее 10 MM") == ["unit_unknown"] and rules("не менее 10 мм, 12В") == []
    assert rules("весы BЕСТА-С60") == ["scale_model"] and rules("весы ВЕСТА-С60") == []


def test_rule_tables():
    t = "| № | Сумма |\n|---|---|\n| 1 | 10,5 |\n| 2 | 1 000 |\n| Итого | 1 010,5 |"
    assert rules(t) == [] and rules(t.replace("1 010,5", "1 011")) == ["table_total_mismatch"]
    assert rules("| № | A |\n|---|---|\n| 1 | x | y |") == ["table_row_cells"]
    assert rules("| № | A |\n|---|---|\n| 1 | a |\n|  | x |\n| 2 | b |") == ["table_empty_number"]
    # потерян номер в «шапке» продолжения
    assert rules("|  | x |\n|---|---|\n| 26 | a |\n| 27 | b |") == ["table_empty_number"]
    # матрица, не нумерация
    assert rules("|  | 2024 |\n|---|---|\n| план | 1 |\n| факт | 2 |") == []
    assert rules("| № | A | B |\n|---|---|---|\n| 1 | a | b |\n| Раздел |  |  |") == []
