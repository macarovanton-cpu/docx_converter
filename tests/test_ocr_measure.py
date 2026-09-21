"""ПРАВКА #87: приёмочные тесты ocr.measure. Verifier — фейковый, ответы строятся из эталонов. Сети нет."""

import json
import socket

import pytest

from ocr import measure as measure_module
from ocr.board import BOARD, GOLDEN, build_board
from ocr.gemini_verifier import VerifierError, VerifierQuotaError
from ocr.measure import Case, build_cases, main, run_measure, score
from ocr_fixtures import require_fixture
from pdf_core import VerifyResult

OFFICE = ("docx1.docx", "xlsx1.xlsx")


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("тест полез в сеть")
    monkeypatch.setattr(socket, "socket", refuse)
    monkeypatch.delenv("GEMINI_LIVE", raising=False)
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)


class FakeVerifier:
    """«Фейковый Verifier из фикстур»: fix при expected != fragment, иначе agree."""
    def __init__(self, answers): self.answers, self.calls = answers, 0
    def verify(self, image_png, fragment, question):
        self.calls += 1
        fixed = self.answers[fragment]
        return VerifyResult("agree" if fixed == fragment else "fix", None if fixed == fragment else fixed, 0.9, "")


class QuotaVerifier(FakeVerifier):
    def verify(self, image_png, fragment, question):
        if self.calls == 9:
            raise VerifierQuotaError("Gemini: HTTP 429")
        return super().verify(image_png, fragment, question)


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
    assert sum(c.image_png is None and c.fixture in OFFICE for c in errors) == 9
    assert all(c.image_png is None for c in cases if c.fixture in OFFICE)
    located = [c for c in errors if c.image_png is not None]
    unlocated = [c for c in errors if c.image_png is None and c.fixture.endswith(".pdf")]
    assert len(located) + len(unlocated) == 54 and len(unlocated) == 4           # проба 21.09.2026: 50 + 4, без запаса
    assert sum(c.ambiguous for c in located) == 2
    assert {c.block_type: sum(x.block_type == c.block_type for x in located) for c in located} == {
        "table": 44, "text": 5, "header": 1}
    assert len(controls) == len(located)
    assert all(c.expected == c.fragment for c in controls) and all(c.expected != c.fragment for c in errors)
    for name, _ in BOARD:
        own_controls = [c for c in controls if c.fixture == name]
        own_located = [c for c in located if c.fixture == name]
        assert [c.id[-2:] for c in own_controls] == [e.id[-2:] for e in own_located]
        assert all(len(c.fragment.split()) == len(e.fragment.split()) for c, e in zip(own_controls, own_located))
    assert len({c.id for c in cases}) == len(cases)
    assert all(c.fragment not in e.fragment for c in controls for e in errors)   # контроль не задевает окна ошибок
    assert build_cases(outputs) == cases                                        # детерминизм, random не участвует
    ir = next(c for c in errors if "IR-камерами" in c.fragment)
    assert "IP-камерами" in ir.expected and ir.page is not None and ir.image_png.startswith(b"\x89PNG")


def test_run_measure_oracle_and_friends(built):
    _, _, cases = built
    errors = [c for c in cases if c.kind == "error"]
    located = [c for c in errors if c.image_png is not None]
    answers = {c.fragment: c.expected for c in cases}

    # «оракул»: знает эталон -> все привязанные найдены, ложных тревог нет
    oracle = FakeVerifier(answers)
    m = run_measure(cases, oracle)
    assert oracle.calls == 2 * len(located)                                     # один случай — один вызов; без картинки — ни одного
    assert m["complete"] and m["totals"]["found"] == len(located) and m["totals"]["false_alarm"] == 0
    assert m["totals"]["no_image"] == 9 and m["totals"]["opcodes"] == len(errors)
    assert m["totals"]["measured"] == len(located) and m["totals"]["unlocated"] == 4
    assert m["totals"]["control"] == m["totals"]["agree"] == len(located)
    assert sum(r["opcodes"] for r in m["by_fixture"]) == len(errors) and len(m["by_fixture"]) == 6
    assert any(r["rule"] == "—" for r in m["by_rule"])                          # «сыручими» правилом не ловится
    assert sum(r["opcodes"] for r in m["by_rule"]) >= m["totals"]["measured"]
    assert sum(r["opcodes"] for r in m["by_block_type"]) == m["totals"]["measured"]
    assert all("image_png" not in row for row in m["cases"])
    json.dumps(m)

    # «всегда согласен»: ничего не найдено, ложных тревог нет
    m = run_measure(cases, FakeVerifier({c.fragment: c.fragment for c in cases}))
    assert m["totals"]["found"] == 0 and m["totals"]["not_found"] == len(located) and m["totals"]["false_alarm"] == 0
    # «всегда правит мусором»: wrong_fix и false_alarm
    m = run_measure(cases, FakeVerifier({c.fragment: "мусор" for c in cases}))
    assert m["totals"]["wrong_fix"] == m["totals"]["false_alarm"] == len(located)

    # «квота на 10-м вызове»: complete False, первые девять исходов на месте
    m = run_measure(cases, QuotaVerifier(answers))
    with_image = [row for row in m["cases"] if row["page"] is not None]
    assert not m["complete"] and [row["outcome"] for row in with_image[:9]] == ["found"] * 9
    assert with_image[9]["outcome"] is None and "429" in with_image[9]["error"]
    assert all(row["outcome"] is None for row in with_image[9:])

    # разбор сломан на одном случае -> outcome "error", замер не остановился
    class Broken(FakeVerifier):
        def verify(self, image_png, fragment, question):
            if fragment == located[0].fragment:
                raise VerifierError("ответ модели не разобран")
            return super().verify(image_png, fragment, question)
    m = run_measure(cases, Broken(answers))
    assert m["complete"] and m["totals"]["error"] == 1 and m["totals"]["found"] == len(located) - 1
    assert next(row for row in m["cases"] if row["outcome"] == "error")["error"].startswith("VerifierError")


def test_score_table():
    def case(kind, image=b"png", fixture="bakeoff.pdf"):
        return Case("x-e01", fixture, kind, None, "было так", "было так" if kind == "control" else "стало так",
                    (), 1, "text", False, image)
    agree, unreadable = VerifyResult("agree", None, 1.0, ""), VerifyResult("unreadable", None, 1.0, "")
    right, wrong = VerifyResult("fix", "стало  так", 1.0, ""), VerifyResult("fix", "мимо", 1.0, "")
    assert [score(case("error"), r) for r in (agree, right, wrong, unreadable)] == [
        "not_found", "found", "wrong_fix", "not_found"]
    assert [score(case("control"), r) for r in (agree, VerifyResult("fix", "было так", 1.0, ""), wrong, unreadable)] == [
        "agree", "false_alarm", "false_alarm", "unreadable"]
    assert score(case("error"), None) == "error"
    assert score(case("error", image=None), agree) == "unlocated"
    assert score(case("error", image=None, fixture="docx1.docx"), agree) == "no_image"


def test_main(built, tmp_path, monkeypatch, capsys):
    _, _, cases = built
    answers = {c.fragment: c.expected for c in cases}
    monkeypatch.setattr(measure_module, "MEASURE_JSON", tmp_path / "m.json")
    monkeypatch.setattr(measure_module, "CROPS_DIR", tmp_path / "crops")
    monkeypatch.setattr(measure_module, "GeminiVerifier", lambda: FakeVerifier(answers))
    assert main([]) == 0
    written = json.loads((tmp_path / "m.json").read_text(encoding="utf-8"))
    assert written["complete"] and "image_png" not in json.dumps(written)
    assert len(list((tmp_path / "crops").rglob("*.png"))) == sum(c.image_png is not None for c in cases)
    assert "сумма строк может превышать measured" in capsys.readouterr().out

    monkeypatch.setattr(measure_module, "GeminiVerifier", lambda: QuotaVerifier(answers))
    assert main([]) == 1                                                        # complete False -> 1, JSON записан
    assert not json.loads((tmp_path / "m.json").read_text(encoding="utf-8"))["complete"]
    assert "429" in capsys.readouterr().err

    # настоящий GeminiVerifier без GEMINI_LIVE и с холодным кэшем: стоп на первом случае, не поход в сеть
    from ocr import gemini_verifier
    monkeypatch.setattr(measure_module, "GeminiVerifier",
                        lambda: gemini_verifier.GeminiVerifier(cache_root=tmp_path / "cold"))
    assert main([]) == 1
    stopped = json.loads((tmp_path / "m.json").read_text(encoding="utf-8"))
    assert not stopped["complete"] and all(row["outcome"] in (None, "unlocated", "no_image") for row in stopped["cases"])
    assert "GEMINI_LIVE" in capsys.readouterr().err
