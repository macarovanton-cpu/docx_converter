"""ПРАВКА #87: приёмочные тесты ocr.measure. Verifier — фейковый, ответы строятся из эталонов. Сети нет.
ПРАВКА #88: модель переписывает полосы, вердикт — judge; фейки отвечают транскрипцией полос."""

import json
import socket
import subprocess
from collections import Counter
from io import BytesIO

import PIL.Image
import pytest

from ocr import measure as measure_module
from ocr.board import BOARD, GOLDEN, build_board
from ocr.gemini_verifier import VerifierError, VerifierQuotaError
from ocr.measure import (MIN_RATIO, TILE_HEIGHT, TRANSCRIBE_QUESTION, Case, build_cases, error_kind, judge, main,
                         page_tiles, run_measure, split_tiles)
from ocr_fixtures import require_fixture

OFFICE = ("docx1.docx", "xlsx1.xlsx")


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("тест полез в сеть или запустил процесс")
    monkeypatch.setattr(socket, "socket", refuse)
    monkeypatch.setattr(subprocess, "run", refuse)
    for name in ("GEMINI_LIVE", "GEMINI_API_KEY", "CLAUDE_CODE_LIVE", "VERIFIER"):
        monkeypatch.delenv(name, raising=False)


class Transcriber:
    """Фейковый Verifier: полоса -> sep.join(<attr> всех случаев на ней). Вызовы записываются."""
    name, model = "fake", "golden"

    def __init__(self, cases, attr, sep):
        texts = {}
        for case in cases:
            for tile in case.tiles:
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
    located = [c for c in errors if c.tiles]
    unlocated = [c for c in errors if not c.tiles and c.fixture.endswith(".pdf")]
    assert len(located) == 50 and len(unlocated) == 4 and sum(c.ambiguous for c in located) == 2
    assert Counter(c.error_kind for c in located) == {"merge": 31, "homoglyph": 11, "chars": 8}
    assert all((c.span[0] == c.span[1]) == (c.tag == "insert") for c in located)
    assert all((c.gold_span[0] == c.gold_span[1]) == (c.tag == "delete") for c in located)
    assert {c.block_type: sum(x.block_type == c.block_type for x in located) for c in located} == {
        "table": 44, "text": 5, "header": 1}
    assert len(controls) == len(located)
    assert all(c.expected == c.fragment for c in controls) and all(c.expected != c.fragment for c in errors)
    assert all(c.span is None and c.gold_span is None and c.error_kind is None for c in controls)
    assert all(0 <= c.gold_span[0] <= c.gold_span[1] <= len(c.expected.split()) for c in errors)
    docx_e01 = next(c for c in errors if c.id == "docx1-e01")                   # вставка в начало документа
    assert docx_e01.tag == "insert" and docx_e01.span == (0, 0) and docx_e01.gold_span[1] > 0
    assert all(t.startswith(b"\x89PNG") and len(t) < 500_000 for c in located for t in c.tiles)
    assert all(PIL.Image.open(BytesIO(t)).size[1] <= TILE_HEIGHT for c in located for t in c.tiles)
    for name, _ in BOARD:
        own_controls = [c for c in controls if c.fixture == name]
        own_located = [c for c in located if c.fixture == name]
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
    assert sum(len(v) for v in pages.values()) == len({t for c in cases for t in c.tiles})

    echo = Echo(cases)
    m = run_measure(cases, echo)
    assert len(echo.calls) == len(pages)                                               # страница — один вызов
    assert [len(x) for x in echo.calls] == [len(v) for v in pages.values()]            # полосы без повторов
    assert m["complete"] and m["totals"]["not_found"] == 50 and m["totals"]["agree"] == 50
    assert m["totals"]["false_alarm"] == 0

    golden = Golden(cases)
    m = run_measure(cases, golden)
    missed = [(r["id"], r["outcome"], r["reading"]) for r in m["cases"] if r["kind"] == "error"
              and r["page"] is not None and r["outcome"] != "found"]
    assert not missed                                   # проба 21.09.2026: 50 из 50; алгоритм под тест не подкручивать
    assert m["totals"]["found"] == 50 and m["totals"]["false_alarm"] == 0 and m["totals"]["neighbor"] == 0
    assert m["totals"]["no_image"] == 9 and m["totals"]["opcodes"] == len(errors) == 63
    assert m["totals"]["measured"] == len(located) and m["totals"]["unlocated"] == 4 and m["totals"]["control"] == 50
    assert (m["verifier"], m["model"], m["calls"], m["schema_version"]) == ("fake", "golden", [], 2)
    assert [r["fixture"] for r in m["by_fixture"]] == [name for name, _ in BOARD]
    assert all("neighbor" in r for r in m["by_fixture"] + m["by_rule"] + m["by_block_type"])
    assert [(r["error_kind"], r["opcodes"], r["found"]) for r in m["by_kind"]] == [
        ("merge", 31, 31), ("homoglyph", 11, 11), ("chars", 8, 8)]
    assert sum(r["opcodes"] for r in m["by_block_type"]) == m["totals"]["measured"]
    assert any(r["rule"] == "—" for r in m["by_rule"])                          # «сыручими» правилом не ловится
    row = next(r for r in m["cases"] if r["id"] == "bakeoff-e01")
    assert row["crop"].startswith("verify_crops/bakeoff/p") and row["crop"].endswith(".png")
    assert row["span"] == [3, 4] and row["ratio"] > 0.5 and row["cache_hit"] is False
    assert all(r["crop"] is None for r in m["cases"] if r["outcome"] in ("unlocated", "no_image"))
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
    rows = {key: [r for r in m["cases"] if (r["fixture"], r["page"]) == key] for key in keys}
    assert not m["complete"] and len(limited.calls) == 3
    assert all(r["outcome"] in ("found", "agree") for key in keys[:2] for r in rows[key])
    assert all(r["outcome"] is None and "usage limit" in r["error"] for r in rows[keys[2]])
    assert all(r["outcome"] is None and r["error"] is None for key in keys[3:] for r in rows[key])

    # VerifierError на одной странице (и ответ не той длины на другой): их случаи — error, замер полный
    class Broken(Transcriber):
        def transcribe(self, images, question):
            answer = super().transcribe(images, question)
            if len(self.calls) == 1:
                raise VerifierError("ответ не по файлам")
            return answer[:-1] if len(self.calls) == 2 else answer
    m = run_measure(cases, Broken(cases, "expected", " ¶ "))
    broken = [r for key in keys[:2] for r in m["cases"] if (r["fixture"], r["page"]) == key]
    assert m["complete"] and m["totals"]["error"] == len(broken)
    assert all(r["outcome"] == "error" for r in broken) and "ответов" in broken[-1]["error"]


def test_main(built, tmp_path, monkeypatch, capsys):
    _, _, cases = built
    bare_make_verifier = measure_module.make_verifier
    crops = tmp_path / "verify_crops"
    monkeypatch.setattr(measure_module, "MEASURE_DIR", tmp_path)
    monkeypatch.setattr(measure_module, "CROPS_DIR", crops)
    monkeypatch.setattr(measure_module, "make_verifier", lambda: Golden(cases))
    assert main([]) == 0
    written = json.loads((tmp_path / "verify_measure.fake.golden.json").read_text(encoding="utf-8"))
    assert written["complete"] and written["totals"]["found"] == 50
    names = sorted(p.relative_to(crops).as_posix() for p in crops.rglob("*.png"))
    assert names == sorted(f"{fixture.removesuffix('.pdf')}/p{page:02d}-{n:02d}.png"
                           for (fixture, page), tiles in page_tiles(cases).items() for n in range(1, len(tiles) + 1))
    assert all((tmp_path / r["crop"]).is_file() for r in written["cases"] if r["crop"])
    assert (tmp_path / "verify_review.fake.golden.md").is_file()
    out = capsys.readouterr().out
    assert "сумма строк может превышать measured" in out and "вызовов 0, картинок 0" in out

    # одна ложная тревога -> в отчёте для сверки её id, полоса и фрагмент
    alarm = next(c for c in cases if c.kind == "control")

    tokens = alarm.fragment.split()
    altered = " ".join(tokens[:3] + ["лишнее"] + tokens[3:])

    class OneAlarm(Transcriber):
        def transcribe(self, images, question):
            return [text.replace(alarm.fragment, altered) for text in super().transcribe(images, question)]
    monkeypatch.setattr(measure_module, "make_verifier", lambda: OneAlarm(cases, "expected", " ¶ "))
    assert main([]) == 0
    review = (tmp_path / "verify_review.fake.golden.md").read_text(encoding="utf-8")
    row = next(r for r in json.loads((tmp_path / "verify_measure.fake.golden.json").read_text(encoding="utf-8"))["cases"]
               if r["id"] == alarm.id)
    assert row["outcome"] == "false_alarm"
    assert f"### {alarm.id} · false_alarm" in review and f"![]({row['crop']})" in review
    assert f"- фрагмент: `{alarm.fragment}`" in review and "Модель ошиблась — случай остаётся ложной тревогой" in review

    # без подмены (#88): бэкенда нет -> код 1, полосы уже записаны
    monkeypatch.setattr(measure_module, "make_verifier", bare_make_verifier)
    monkeypatch.setattr(measure_module, "MEASURE_DIR", tmp_path / "bare")
    monkeypatch.setattr(measure_module, "CROPS_DIR", tmp_path / "bare" / "verify_crops")
    capsys.readouterr()
    assert main([]) == 1
    assert "ПРАВКА #89" in capsys.readouterr().err
    assert len(list((tmp_path / "bare" / "verify_crops").rglob("p*-*.png"))) == len(
        {t for c in cases for t in c.tiles})
