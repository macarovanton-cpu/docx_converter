"""ПРАВКА #87: приёмочные тесты ocr.measure. Verifier — фейковый, ответы строятся из эталонов. Сети нет.
ПРАВКА #88: модель переписывает полосы, вердикт — judge; фейки отвечают транскрипцией полос.
ПРАВКА #89: выбор бэкенда, --fixture, настоящий ClaudeCodeVerifier без CLAUDE_CODE_LIVE.
ПРАВКА #90: физическая страница (*_model.json), стык страниц, narrow, приклеенный маркер, промах кэша — не стоп."""

import functools
import json
import socket
import subprocess
from collections import Counter
from io import BytesIO

import PIL.Image
import pytest

from ocr import measure as measure_module
from ocr.board import BOARD, FIXTURES, GOLDEN, build_board
from ocr.claude_code_verifier import CLAUDE_MODEL, ClaudeCodeTimeoutError, ClaudeCodeVerifier
from ocr.gemini_verifier import CACHE_ROOT, VerifierConfigError, VerifierError, VerifierQuotaError
from ocr.measure import (CONTROL_OUTCOMES, ERROR_OUTCOMES, ERROR_REASONS, MIN_RATIO, PAGE_SOURCES, TILE_HEIGHT,
                         TRANSCRIBE_QUESTION, Case, _content_list, build_cases, error_kind, error_reason, judge, main,
                         make_verifier, model_pages, norm, page_tiles, run_measure, split_tiles, table_run)
from ocr_fixtures import require_fixture

OFFICE = ("docx1.docx", "xlsx1.xlsx")
bare_make_verifier = make_verifier


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("тест полез в сеть или запустил процесс")
    monkeypatch.setattr(socket, "socket", refuse)
    monkeypatch.setattr(subprocess, "run", refuse)
    for name in ("GEMINI_LIVE", "GEMINI_API_KEY", "CLAUDE_CODE_LIVE", "VERIFIER"):
        monkeypatch.delenv(name, raising=False)


class Transcriber:
    """Фейковый Verifier: полоса -> sep.join(<attr> всех случаев, у кого она в tiles или next_tiles). Вызовы пишутся."""
    name, model = "fake", "golden"

    def __init__(self, cases, attr, sep):
        texts = {}
        for case in cases:
            for tile in case.tiles + case.next_tiles:     # ПРАВКА #90: стык
                texts.setdefault(tile, []).append(getattr(case, attr))
        self.texts = {tile: sep.join(parts) for tile, parts in texts.items()}
        self.calls = []

    def transcribe(self, images, question):
        assert question == TRANSCRIBE_QUESTION
        self.calls.append(images)
        return [self.texts[image] for image in images]


def Echo(cases):
    return Transcriber(cases, "fragment", " ")


def Golden(cases):
    return Transcriber(cases, "expected", " ¶ ")


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    for pair in BOARD:
        for name in pair:
            if name is not None:
                require_fixture(name)
    for golden_name in GOLDEN.values():
        require_fixture(golden_name)
    board, outputs = build_board(tmp_path_factory.mktemp("board"))
    return board, outputs, build_cases(outputs)


def test_build_cases(built):
    board, outputs, cases = built
    errors = [c for c in cases if c.kind == "error"]
    controls = [c for c in cases if c.kind == "control"]
    assert len(errors) == sum(r["count_diffs"] for r in board["rows"])          # 63 на 21.09.2026 — замер и табло едины
    assert sum(not c.tiles and c.fixture in OFFICE for c in errors) == 9
    assert all(not c.tiles for c in cases if c.fixture in OFFICE)
    # ПРАВКА #90: 63 опкода = 46 с полосами + 4 narrow + 4 unlocated + 9 no_image; контроль — 50, как в v2
    located = [c for c in errors if c.tiles]
    narrow = [c for c in errors if c.page_source == "narrow"]
    unlocated = [c for c in errors if not c.tiles and c.fixture.endswith(".pdf") and c.page_source != "narrow"]
    assert len(located) == 46 and len(unlocated) == 4 and sum(c.ambiguous for c in located) == 0
    assert {c.id for c in narrow} == {"bakeoff-e12", "bakeoff-e13", "bakeoff2-e19", "bakeoff2-e20"}
    assert all(c.page is not None and not c.tiles and not c.next_tiles for c in narrow)
    assert Counter(c.error_kind for c in located) == {"merge": 31, "homoglyph": 8, "chars": 7}
    assert all((c.span[0] == c.span[1]) == (c.tag == "insert") for c in located)
    assert all((c.gold_span[0] == c.gold_span[1]) == (c.tag == "delete") for c in located)
    assert {t: sum(c.block_type == t for c in located) for t in {c.block_type for c in located}} == {
        "table": 44, "text": 1, "header": 1}
    assert len(controls) == 50 and all(c.tiles for c in controls)
    assert Counter(c.page_source for c in located) == {"content_list": 20, "model_json": 21, "seam": 5}
    seams = [c for c in located if c.page_source == "seam"]
    assert {c.id for c in seams} == {"bakeoff2-e07", "bakeoff2-e08", "bakeoff2-e11", "bakeoff3-e03", "bakeoff3-e05"}
    assert all(c.next_tiles and c.error_kind == "merge" for c in seams)
    assert all(not c.next_tiles for c in cases if c.page_source != "seam")
    assert sum(c.page_source == "model_json" for c in controls) >= 27    # 8 bakeoff + 17 bakeoff2 + 2 bakeoff3
    assert all(c.page_source in PAGE_SOURCES for c in controls)            # unresolved-окно заменено следующим
    assert next(c for c in controls if c.id == "textpdf1-c11").fragment != "+50 Диапазон температуры для приборов °C:– М0601"
    page = {c.id: c.page for c in cases}                                   # регрессия v2: физические страницы
    assert (page["bakeoff-e02"], page["bakeoff-c13"], page["bakeoff2-e12"], page["textpdf1-e11"]) == (3, 9, 8, 9)
    assert (page["bakeoff3-e05"], next(c for c in cases if c.id == "bakeoff-c05").list_page) == (5, 2)
    # «после»: текст каждого перенесённого случая есть на его странице *_model.json (у стыка — на стыке двух)
    for c in located + controls:
        if c.page_source in ("model_json", "seam"):
            mp = model_pages(FIXTURES / dict(BOARD)[c.fixture])
            text = mp[c.page - 1] + (mp[c.page] if c.page_source == "seam" else "")
            tokens = c.fragment.split()
            needle = "".join(tokens if c.kind == "control" else tokens[slice(*c.span)])
            assert not needle or needle in text, c.id
    assert all(c.expected == c.fragment for c in controls) and all(c.expected != c.fragment for c in errors)
    assert all(c.span is None and c.gold_span is None and c.error_kind is None for c in controls)
    assert all(0 <= c.gold_span[0] <= c.gold_span[1] <= len(c.expected.split()) for c in errors)
    docx_e01 = next(c for c in errors if c.id == "docx1-e01")                   # вставка в начало документа
    assert docx_e01.tag == "insert" and docx_e01.span == (0, 0) and docx_e01.gold_span[1] > 0
    assert all(t.startswith(b"\x89PNG") and len(t) < 500_000 for c in located for t in c.tiles)
    assert all(PIL.Image.open(BytesIO(t)).size[1] <= TILE_HEIGHT for c in located for t in c.tiles)
    for name, _ in BOARD:
        own_controls = [c for c in controls if c.fixture == name]
        own_located = [c for c in errors if c.fixture == name and c.page_source is not None]   # с блоком, как в v2
        assert [c.id[-2:] for c in own_controls] == [e.id[-2:] for e in own_located]
        assert all(len(c.fragment.split()) == len(e.fragment.split()) for c, e in zip(own_controls, own_located))
    assert len({c.id for c in cases}) == len(cases)
    assert all(c.fragment not in e.fragment for c in controls for e in errors)   # контроль не задевает окна ошибок
    assert build_cases(outputs) == cases                                        # детерминизм, random не участвует
    ir = next(c for c in errors if "IR-камерами" in c.fragment)
    assert "IP-камерами" in ir.expected and ir.page is not None and ir.error_kind == "chars"
    assert ir.fragment.split()[ir.span[0]] == "IR-камерами" and ir.expected.split()[ir.gold_span[0]] == "IP-камерами"
    # expected — эталон на весь пролёт окна: соседний опкод (IP54не) тоже исправлен
    e07 = next(c for c in errors if c.id == "bakeoff-e07")
    assert e07.fragment.split()[slice(*e07.span)] == ["IP65не"] and "IP54не" in e07.fragment
    assert e07.expected == "комплекс не ниже IP65 не ниже IP54 не ниже" and e07.gold_span == (3, 5)


def test_split_tiles():
    image = PIL.Image.new("RGB", (40, 1500))
    for y in range(1500):
        for x in range(40):
            image.putpixel((x, y), (y // 6, 0, 0))
    buffer = BytesIO()
    image.save(buffer, "PNG")
    tiles = split_tiles(buffer.getvalue())
    opened = [PIL.Image.open(BytesIO(t)) for t in tiles]
    assert [o.size for o in opened] == [(40, 700)] * 3 and all(o.mode == "L" for o in opened)
    reference = image.convert("L")
    assert [o.getpixel((0, 0)) for o in opened] == [reference.getpixel((0, y)) for y in (0, 500, 800)]
    assert opened[-1].getpixel((0, 699)) == reference.getpixel((0, 1499))
    assert split_tiles(buffer.getvalue()) == tiles                              # детерминизм

    small = BytesIO()
    image.crop((0, 0, 40, 600)).save(small, "PNG")
    one = split_tiles(small.getvalue())
    assert len(one) == 1 and PIL.Image.open(BytesIO(one[0])).size == (40, 600)


def test_table_run_model_pages(built):
    """ПРАВКА #90: склейка межстраничной таблицы в content_list и постраничный текст *_model.json."""
    cl = {name: _content_list(FIXTURES / zip_name) for name, zip_name in BOARD if zip_name is not None}
    assert table_run(cl["bakeoff2.pdf"], 3) == [3, 8, 10, 12, 14, 16, 18, 20]
    assert table_run(cl["bakeoff.pdf"], 17) == [17, 19, 21] and table_run(cl["textpdf1.pdf"], 102) == [102]
    assert table_run(cl["bakeoff.pdf"], 7) == [7, 9, 11, 13] and table_run(cl["bakeoff3.pdf"], 23) == [23, 26]
    assert [len(model_pages(FIXTURES / dict(BOARD)[n])) for n in ("bakeoff.pdf", "bakeoff2.pdf", "bakeoff3.pdf",
                                                                  "textpdf1.pdf")] == [9, 8, 9, 10]


def test_norm_and_error_reason():
    # ПРАВКА #90: маркер списка, приклеенный MinerU к предыдущему токену, отделяется; перенос и «-слово» — нет
    assert norm(["DBM14G;–", "связи).-", "°C:–", "помещений-", "-АКЗ", "–"]) == [
        "DBM14G;", "связи).", "°C:", "помещений-", "-АКЗ"]
    assert [error_reason(e) for e in (
        VerifierConfigError("нет в кэше, вызов выключен: нужен CLAUDE_CODE_LIVE=1"), VerifierQuotaError("лимит"),
        ClaudeCodeTimeoutError("claude -p: нет ответа"), VerifierError("claude -p: ответ не JSON: …"),
        VerifierError("ответов 1, картинок 2"), VerifierError("ответ модели не JSON: …"),
        VerifierError("ответ не по файлам: …"), VerifierError("прочее"))] == [
        "miss", "stop", "timeout", "cli", "count", "parse", "parse", "other"]


def error_case(fragment, expected, span, gold_span):
    return Case("x-e01", "bakeoff.pdf", "error", "replace", fragment, expected, span, gold_span,
                error_kind(fragment.split()[span[0]:span[1]], expected.split()[gold_span[0]:gold_span[1]]),
                (), 1, "table", False, (b"png",))


def control_case(fragment):
    return Case("x-c01", "bakeoff.pdf", "control", None, fragment, fragment, None, None, None,
                (), 1, "table", False, (b"png",))


IR = error_case("оснащение уличными PoE IR-камерами видеонаблюдения с моторизированным",
                "оснащение уличными PoE IP-камерами видеонаблюдения с моторизированным", (3, 4), (3, 4))


def test_judge_errors():
    # склейка при соседней склейке в окне: обе исправлены, свой пролёт чист
    ip = error_case("комплекс не ниже IP65не ниже IP54не ниже", "комплекс не ниже IP65 не ниже IP54 не ниже",
                    (3, 4), (3, 5))
    assert ip.error_kind == "merge"
    r = judge(ip, "программно-технический\nкомплекс не ниже IP65\nне ниже IP54\nне ниже IP20")
    assert r["outcome"] == "found" and r["verdict"] == "fix" and len(r["ops"]) == 2
    assert [op[3] for op in r["ops"]] == ["merge", "merge"]

    # маркер списка выброшен norm
    r = judge(IR, "- оснащение уличными PoE IP-камерами видеонаблюдения с моторизированным объективом")
    assert r["outcome"] == "found" and r["clipped"] == 0
    assert r["reading"] == "оснащение уличными PoE IP-камерами видеонаблюдения с моторизированным"
    assert r["ops"] == [["replace", "IR-камерами", "IP-камерами", "chars"]]

    # вставка: пустой пролёт опкода
    ins = error_case("форме № КС-3. Заказчик перечисляет Подрядчику",
                     "форме № КС-3. 3. Заказчик перечисляет Подрядчику", (3, 3), (3, 4))
    assert judge(ins, "по форме № КС-3.\n3. Заказчик перечисляет Подрядчику")["outcome"] == "found"

    # гомоглиф
    homo = error_case("04 2026 г. A.C. Гузь A. C.", "04 2026 г. А.С. Гузь А. С.", (3, 4), (3, 4))
    assert homo.error_kind == "homoglyph"
    r = judge(homo, "04 2026 г. А.С. Гузь А. С.")
    assert r["outcome"] == "found" and r["ops"][0][3] == "homoglyph"

    # правка не туда / туда, но не как в эталоне / согласие
    assert judge(IR, "оснащение уличными PoE IR-камерами видеонаблюденя с моторизированным")["outcome"] == "neighbor"
    r = judge(IR, "оснащение уличными PoE IK-камерами видеонаблюдения с моторизированным")
    assert r["outcome"] == "wrong_fix" and r["verdict"] == "fix"
    r = judge(IR, "оснащение уличными PoE IR-камерами видеонаблюдения с моторизированным")
    assert r["outcome"] == "not_found" and r["verdict"] == "agree" and r["ops"] == [] and r["ratio"] == 1.0

    # мусор и пустая транскрипция
    r = judge(IR, "совсем другой текст про погоду на завтра")
    assert r["ratio"] < MIN_RATIO and r["outcome"] == "not_found" and r["verdict"] == "unreadable"
    assert judge(IR, "")["outcome"] == "not_found"

    # хвост фрагмента за краем полосы: не в вердикте, исход по остальному
    r = judge(IR, "оснащение уличными PoE IP-камерами видеонаблюдения с")
    assert r["clipped"] == 1 and r["outcome"] == "found" and len(r["ops"]) == 1
    # сам пролёт опкода за краем: not_found, чтение "clipped"
    r = judge(IR, "видеонаблюдения с моторизированным")
    assert r["ratio"] >= MIN_RATIO and r["outcome"] == "not_found" and r["reading"] == "clipped"


def test_judge_glued_marker():
    """ПРАВКА #90: единственный wrong_fix v2 был артефактом сравнения — тире, приклеенным MinerU к токену."""
    e10 = Case(id="t-e10", fixture="textpdf1.pdf", kind="error", tag="replace", error_kind="homoglyph",
               fragment="740, DHM9B, DBM14G;– MB150; C;– H4, M100;– ZSFY, ZSFY-D,",
               expected="740, DHM9B, DBM14G;– МВ150; С;– Н4, M100;– ZSFY, ZSFY-D,", span=(3, 6), gold_span=(3, 6),
               rules=(), page=8, block_type="table", ambiguous=False, tiles=())
    r = judge(e10, "740, DHM9B, DBM14G;\n– MB150; C;\n– H4, M100;\n– ZSFY, ZSFY-D,")
    assert r["outcome"] == "not_found" and r["ops"] == [] and r["clipped"] == 0      # в v2 здесь был wrong_fix


def test_judge_control():
    c = control_case("числом оптических волокон не менее 4(тип волокон")
    r = judge(c, "- числом оптических волокон не менее 4(тип волокон")
    assert r["outcome"] == "agree" and r["verdict"] == "agree"
    r = judge(c, "числом оптических волоконне менее 4(тип волокон")
    assert r["outcome"] == "false_alarm" and r["ops"] == [["replace", "волокон не", "волоконне", "merge"]]
    # PLACEHOLDER 2: расхождение у края окна неотличимо от края полосы — срезано, ложная тревога недосчитана
    r = judge(c, "числом оптических волокон не менее 4 (тип волокон")
    assert r["outcome"] == "agree" and r["clipped"] == 2
    assert judge(c, "совсем другой текст про погоду на завтра")["outcome"] == "unreadable"


def test_run_measure(built):
    _, _, cases = built
    errors = [c for c in cases if c.kind == "error"]
    located = [c for c in errors if c.tiles]
    pages = page_tiles(cases)
    assert sum(len(v) for v in pages.values()) == len({t for c in cases for t in c.tiles + c.next_tiles})

    echo = Echo(cases)
    m = run_measure(cases, echo)
    assert len(echo.calls) == len(pages)                                               # страница — один вызов
    assert [len(x) for x in echo.calls] == [len(v) for v in pages.values()]            # полосы без повторов
    assert m["complete"] and m["totals"]["not_found"] == 46 and m["totals"]["agree"] == 50
    assert m["totals"]["false_alarm"] == 0 and m["totals"]["narrow"] == 4 and not m["missing_pages"]

    golden = Golden(cases)
    m = run_measure(cases, golden)
    missed = [(r["id"], r["outcome"], r["reading"]) for r in m["cases"] if r["kind"] == "error"
              and r["page"] is not None and r["outcome"] not in ("found", "narrow")]
    assert not missed                                   # проба 22.09.2026: 46 из 46; алгоритм под тест не подкручивать
    assert m["totals"]["found"] == 46 and m["totals"]["false_alarm"] == 0 and m["totals"]["neighbor"] == 0
    assert m["totals"]["no_image"] == 9 and m["totals"]["opcodes"] == len(errors) == 63
    assert m["totals"]["measured"] == len(located) and m["totals"]["unlocated"] == 4 and m["totals"]["control"] == 50
    assert (m["verifier"], m["model"], m["calls"], m["schema_version"]) == ("fake", "golden", [], 3)
    sources = m["totals"]["page_sources"]                                  # 46 ошибок + 50 контрольных с полосами
    assert sources["seam"] == 5 and sources["model_json"] >= 48 and sum(sources.values()) == 96
    assert set(m["totals"]["error_reasons"]) == set(ERROR_REASONS) and not any(m["totals"]["error_reasons"].values())
    assert sum(r["moved"] for r in m["by_fixture"]) == sources["model_json"] + sources["seam"]
    seam_row = next(r for r in m["cases"] if r["id"] == "bakeoff3-e05")     # окно через стык судится по двум полосам
    assert seam_row["outcome"] == "found" and seam_row["page"] == 5 and seam_row["list_page"] == 5
    assert seam_row["crop"].startswith("verify_crops/bakeoff3/p05-")
    assert seam_row["crop_next"].startswith("verify_crops/bakeoff3/p06-")
    assert all(r["crop_next"] is None for r in m["cases"] if r["page_source"] != "seam")
    assert [r["fixture"] for r in m["by_fixture"]] == [name for name, _ in BOARD]
    assert all("neighbor" in r for r in m["by_fixture"] + m["by_rule"] + m["by_block_type"])
    assert [(r["error_kind"], r["opcodes"], r["found"]) for r in m["by_kind"]] == [
        ("merge", 31, 31), ("homoglyph", 8, 8), ("chars", 7, 7)]
    assert sum(r["opcodes"] for r in m["by_block_type"]) == m["totals"]["measured"]
    assert any(r["rule"] == "—" for r in m["by_rule"])                          # «сыручими» правилом не ловится
    row = next(r for r in m["cases"] if r["id"] == "bakeoff-e01")
    assert row["crop"].startswith("verify_crops/bakeoff/p") and row["crop"].endswith(".png")
    assert row["span"] == [3, 4] and row["ratio"] > 0.5 and row["cache_hit"] is False
    assert all(r["crop"] is None for r in m["cases"] if r["outcome"] in ("unlocated", "no_image", "narrow"))
    assert all("tiles" not in r for r in m["cases"]) and json.dumps(m)          # картинки в JSON не пишутся

    # лимит на 3-м вызове: две первые страницы отвечены, на третьей — текст ошибки, дальше исходов нет
    class Limited(Transcriber):
        def transcribe(self, images, question):
            if len(self.calls) == 2:
                self.calls.append(images)
                raise VerifierQuotaError("лимит подписки Claude Code: usage limit reached")
            return super().transcribe(images, question)
    limited = Limited(cases, "expected", " ¶ ")
    m = run_measure(cases, limited)
    keys = list(pages)
    # строки страницы: только случаи с полосами — narrow сидит на своей странице и исход имеет всегда
    rows = {key: [r for r in m["cases"] if (r["fixture"], r["page"]) == key and r["page_source"] in PAGE_SOURCES]
            for key in keys}
    assert not m["complete"] and len(limited.calls) == 3 and not m["missing_pages"]
    assert all(r["outcome"] in ("found", "agree") for key in keys[:2] for r in rows[key])
    assert all(r["outcome"] is None and "usage limit" in r["error"] for r in rows[keys[2]])
    assert all(r["error_reason"] == "stop" for r in rows[keys[2]])
    assert all(r["outcome"] is None and r["error"] is None for key in keys[3:] for r in rows[key])
    assert m["totals"]["error_reasons"]["stop"] == len(rows[keys[2]])

    # ПРАВКА #90: промах кэша на 2-й странице — не стоп: спрошены все страницы, оплаченные пересужены
    class Miss(Transcriber):
        def transcribe(self, images, question):
            if len(self.calls) == 1:
                self.calls.append(images)
                raise VerifierConfigError("нет в кэше, вызов выключен: нужен CLAUDE_CODE_LIVE=1")
            return super().transcribe(images, question)
    miss = Miss(cases, "expected", " ¶ ")
    m = run_measure(cases, miss)
    assert not m["complete"] and len(miss.calls) == len(pages)          # спрошены все страницы, а не до первого промаха
    assert m["missing_pages"] == [list(keys[1])]
    got = {key: [r for r in m["cases"] if (r["fixture"], r["page"]) == key and r["page_source"] in PAGE_SOURCES]
           for key in keys}
    assert all(r["outcome"] is None and r["error_reason"] == "miss" for r in got[keys[1]])
    assert m["totals"]["error_reasons"]["miss"] == len(got[keys[1]]) > 0
    assert all(r["outcome"] in ("found", "agree") for key in keys if key != keys[1] for r in got[key])

    # VerifierError на одной странице (и ответ не той длины на другой): их случаи — error, замер полный
    class Broken(Transcriber):
        def transcribe(self, images, question):
            answer = super().transcribe(images, question)
            if len(self.calls) == 1:
                raise VerifierError("непонятный сбой")           # ПРАВКА #90: причина other
            return answer[:-1] if len(self.calls) == 2 else answer
    m = run_measure(cases, Broken(cases, "expected", " ¶ "))
    broken = [r for key in keys[:2] for r in m["cases"] if (r["fixture"], r["page"]) == key]
    assert m["complete"] and m["totals"]["error"] == len(broken)
    assert all(r["outcome"] == "error" for r in broken) and "ответов" in broken[-1]["error"]
    assert {r["error_reason"] for r in broken} == {"other", "count"}


def test_main(built, tmp_path, monkeypatch, capsys):
    _, _, cases = built
    crops = tmp_path / "verify_crops"
    monkeypatch.setattr(measure_module, "MEASURE_DIR", tmp_path)
    monkeypatch.setattr(measure_module, "CROPS_DIR", crops)
    monkeypatch.setattr(measure_module, "make_verifier", lambda name, model: Golden(cases))
    assert main([]) == 0
    written = json.loads((tmp_path / "verify_measure.v3.fake.golden.json").read_text(encoding="utf-8"))
    assert written["complete"] and written["totals"]["found"] == 46 and written["missing_pages"] == []
    names = sorted(p.relative_to(crops).as_posix() for p in crops.rglob("*.png"))
    assert names == sorted(f"{fixture.removesuffix('.pdf')}/p{page:02d}-{n:02d}.png"
                           for (fixture, page), tiles in page_tiles(cases).items() for n in range(1, len(tiles) + 1))
    assert all((tmp_path / r["crop"]).is_file() for r in written["cases"] if r["crop"])
    assert all((tmp_path / r["crop_next"]).is_file() for r in written["cases"] if r["crop_next"])
    assert (tmp_path / "verify_review.v3.fake.golden.md").is_file()
    out = capsys.readouterr().out
    assert "сумма строк может превышать measured" in out and "вызовов 0, картинок 0" in out


@pytest.fixture
def quick(built, tmp_path, monkeypatch):
    """main без пересборки табло: build_board / build_cases отдают готовое, пути — во временной папке."""
    _, outputs, cases = built
    monkeypatch.setattr(measure_module, "MEASURE_DIR", tmp_path)
    monkeypatch.setattr(measure_module, "CROPS_DIR", tmp_path / "verify_crops")
    monkeypatch.setattr(measure_module, "build_board", lambda work, fixtures: (None, outputs))
    monkeypatch.setattr(measure_module, "build_cases", lambda outputs, fixtures: cases)
    return cases


def test_review(quick, tmp_path, monkeypatch):
    # одна ложная тревога -> в отчёте для сверки её id, полоса и фрагмент
    cases = quick
    alarm = next(c for c in cases if c.kind == "control")
    tokens = alarm.fragment.split()
    altered = " ".join(tokens[:3] + ["лишнее"] + tokens[3:])

    class OneAlarm(Transcriber):
        def transcribe(self, images, question):
            return [text.replace(alarm.fragment, altered) for text in super().transcribe(images, question)]
    monkeypatch.setattr(measure_module, "make_verifier", lambda name, model: OneAlarm(cases, "expected", " ¶ "))
    assert main([]) == 0
    review = (tmp_path / "verify_review.v3.fake.golden.md").read_text(encoding="utf-8")
    row = next(r for r in json.loads((tmp_path / "verify_measure.v3.fake.golden.json").read_text(encoding="utf-8"))["cases"]
               if r["id"] == alarm.id)
    assert row["outcome"] == "false_alarm"
    assert f"### {alarm.id} · false_alarm" in review and f"![]({row['crop']})" in review
    assert f"- фрагмент: `{alarm.fragment}`" in review and "Модель ошиблась — случай остаётся ложной тревогой" in review
    assert "`∅` → `лишнее` (chars)" in review


def test_review_seam(quick, tmp_path, monkeypatch):
    """ПРАВКА #90: у случая через стык в отчёте две полосы — низ страницы N и верх N+1."""
    cases = quick
    seam = next(c for c in cases if c.page_source == "seam")
    tokens = seam.fragment.split()
    altered = " ".join(tokens[:1] + ["лишнее"] + tokens[1:])      # правка вне пролёта опкода -> neighbor

    class Neighbor(Transcriber):
        def transcribe(self, images, question):
            return [text.replace(seam.fragment, altered) for text in super().transcribe(images, question)]
    monkeypatch.setattr(measure_module, "make_verifier", lambda name, model: Neighbor(cases, "fragment", " "))
    assert main([]) == 0
    row = next(r for r in json.loads((tmp_path / "verify_measure.v3.fake.golden.json").read_text(
        encoding="utf-8"))["cases"] if r["id"] == seam.id)
    review = (tmp_path / "verify_review.v3.fake.golden.md").read_text(encoding="utf-8")
    assert row["outcome"] == "neighbor" and row["crop_next"]
    assert f"### {seam.id} · neighbor · стр. {seam.page}|{seam.page + 1}" in review
    assert f"![]({row['crop']})" in review and f"![]({row['crop_next']})" in review


def test_backends(quick, tmp_path, monkeypatch, capsys):
    cases = quick
    with pytest.raises(VerifierConfigError, match="#87"):
        make_verifier("gemini", None)
    assert isinstance(make_verifier("claude-code", None), ClaudeCodeVerifier)
    assert make_verifier("claude-code", None).model == CLAUDE_MODEL
    assert make_verifier("claude-code", "claude-opus-5").model == "claude-opus-5"

    assert main(["--verifier", "gemini"]) == 1 and "#87" in capsys.readouterr().err
    monkeypatch.setenv("VERIFIER", "gemini")
    assert main([]) == 1 and "#87" in capsys.readouterr().err
    monkeypatch.delenv("VERIFIER")
    assert main(["--verifier", "openai"]) == 1 and main(["--fixture", "nope.pdf"]) == 1

    # --fixture: остаются случаи только этой фикстуры, имя файла с суффиксом
    asked = []
    monkeypatch.setattr(measure_module, "make_verifier", lambda name, model: asked.append((name, model)) or Golden(cases))
    assert main(["--fixture", "bakeoff.pdf", "--model", "claude-opus-5"]) == 0    # имя выхода — v3.<бэкенд>.<модель>
    assert asked == [("claude-code", "claude-opus-5")]
    one = json.loads((tmp_path / "verify_measure.v3.fake.golden.bakeoff.json").read_text(encoding="utf-8"))
    assert {r["fixture"] for r in one["cases"]} == {"bakeoff.pdf"} and [r["fixture"] for r in one["by_fixture"]] == [
        "bakeoff.pdf"]
    assert (tmp_path / "verify_review.v3.fake.golden.bakeoff.md").is_file()
    # полосы остальных фикстур (записаны прогонами с gemini выше) не стёрты
    assert sorted(p.name for p in (tmp_path / "verify_crops").iterdir()) == ["bakeoff", "bakeoff2", "bakeoff3", "textpdf1"]
    assert main(["--fixture", "textpdf1.pdf", "--fixture", "bakeoff.pdf"]) == 0
    assert (tmp_path / "verify_measure.v3.fake.golden.bakeoff+textpdf1.json").is_file()

    # настоящий ClaudeCodeVerifier, холодный кэш, без CLAUDE_CODE_LIVE: стоп с кодом 1, claude не запускался
    runs = []
    monkeypatch.setattr(subprocess, "run", lambda *args, **kwargs: runs.append(args))
    monkeypatch.setattr(measure_module, "make_verifier", bare_make_verifier)
    monkeypatch.setattr(measure_module, "ClaudeCodeVerifier",
                        functools.partial(ClaudeCodeVerifier, cache_root=tmp_path / "cold"))
    capsys.readouterr()
    assert main([]) == 1
    err = capsys.readouterr().err
    assert "CLAUDE_CODE_LIVE" in err and "нет в кэше" in err and runs == []
    cold = json.loads((tmp_path / "verify_measure.v3.claude-code.claude-sonnet-5.json").read_text(encoding="utf-8"))
    assert not cold["complete"] and cold["calls"] == []
    assert [tuple(key) for key in cold["missing_pages"]] == list(page_tiles(cases))    # холодный кэш — все страницы
    assert all(r["outcome"] in (None, "unlocated", "no_image", "narrow") for r in cold["cases"])
    assert all(r["error_reason"] in (None, "miss") for r in cold["cases"])
    assert not (tmp_path / "cold").exists()


V3_NEW_PAGES = ({("bakeoff.pdf", p) for p in (3, 4, 5, 8, 9)} | {("bakeoff2.pdf", p) for p in range(2, 9)}
                | {("bakeoff3.pdf", 2), ("bakeoff3.pdf", 6), ("textpdf1.pdf", 9)})    # живой добор человека, #90


@pytest.mark.parametrize("model, fixture", [("claude-sonnet-5", None), ("claude-opus-5", "bakeoff.pdf")])
def test_paid_pages_are_free(built, model, fixture):
    """ПРАВКА #90, условие приёмки: ключи кэша не изменились — оплаченные в v2 страницы пересуживаются без вызовов."""
    _, _, cases = built
    if not any(CACHE_ROOT.glob("*.json")):
        pytest.skip("нет .cache/ocr/verify: живого замера на этой машине не было")
    own = [c for c in cases if fixture in (None, c.fixture)]
    verifier = ClaudeCodeVerifier(model=model)          # CLAUDE_CODE_LIVE снят фикстурой: только кэш
    m = run_measure(own, verifier)
    missing = {tuple(key) for key in m["missing_pages"]}
    assert verifier.network_calls == 0 and missing <= V3_NEW_PAGES     # промах на странице v2 — уехали ключи кэша
    assert m["totals"]["error_reasons"]["parse"] == 0 and m["totals"]["error"] == 0
    judged = 0
    for row in m["cases"]:
        keys = {(row["fixture"], row["page"])}
        if row["page_source"] == "seam":
            keys.add((row["fixture"], row["page"] + 1))
        if row["page_source"] in PAGE_SOURCES and not keys & missing:
            assert row["cache_hit"] and row["outcome"] in ERROR_OUTCOMES + CONTROL_OUTCOMES, row["id"]
            judged += 1
    assert judged >= (40 if fixture is None else 6)
    outcome = {row["id"]: row["outcome"] for row in m["cases"]}
    if fixture is None:      # два ответа с дописанной копией JSON: 39 error v2 -> исходы (ПРАВКА #90, parse_texts)
        assert (outcome["bakeoff3-e04"], outcome["bakeoff3-c03"], outcome["textpdf1-e10"]) == (
            "found", "agree", "not_found")
        assert all(row["outcome"] for row in m["cases"] if row["fixture"] == "bakeoff2.pdf" and row["page"] == 1)
