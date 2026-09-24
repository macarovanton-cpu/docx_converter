"""ПРАВКА #91: приёмочные тесты ocr.vision — сверка по картинке в тракте. Verifier — фейковый ("эталон" / "эхо"),
make_verifier подменяется; сети и claude нет: socket.socket и subprocess.run — заглушки."""

import io
import json
import socket
import subprocess
import zipfile
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path

import pytest

from file_converter import get_pdf_page_count
from ocr import vision as vision_module
from ocr.board import BOARD, GOLDEN, build_board
from ocr.cache import LocalCache, build_meta, cache_key
from ocr.claude_code_verifier import (ClaudeCodeLimitError, ClaudeCodeMissingError, ClaudeCodeTimeoutError,
                                      ClaudeCodeVerifier)
from ocr.cli import _parse_args, main
from ocr.gemini_verifier import CONTEXT_TOKENS
from ocr.ingest import ingest
from ocr.measure import TRANSCRIBE_QUESTION, model_pages, page_tiles, run_measure
from ocr.validate import strip_annotations
from ocr.vision import (PAGE_HINT_LIMIT, VISION_MODEL, VISION_STRIDE, Hint, _touches, build_windows, hints,
                        is_cosmetic, limit_pages, make_verifier, plain_tokens, vision_findings)
from ocr_fixtures import FIXTURES, read_raw, require_fixture

PDFS = [name for name, zip_name in BOARD if zip_name is not None]
ZIPS = dict(BOARD)
KEYS = ["id", "rule", "severity", "page", "snippet", "suggestion", "reading", "model"]

# состав окон и покрытие — числа Шага 0 спеки 16 (проба 22.09.2026)
WINDOWS = {"bakeoff.pdf": 853, "bakeoff2.pdf": 690, "bakeoff3.pdf": 776, "textpdf1.pdf": 729}
PLACED = {"bakeoff.pdf": 829, "bakeoff2.pdf": 673, "bakeoff3.pdf": 753, "textpdf1.pdf": 602}
SKIPPED = {"bakeoff.pdf": {"unlocated": 22, "unresolved": 2}, "bakeoff2.pdf": {"unlocated": 17, "unresolved": 0},
           "bakeoff3.pdf": {"unlocated": 23, "unresolved": 0}, "textpdf1.pdf": {"unlocated": 124, "unresolved": 3}}
SOURCES = {"bakeoff.pdf": Counter(content_list=321, model_json=491, narrow=9, seam=8),
           "bakeoff2.pdf": Counter(content_list=69, model_json=565, narrow=18, seam=21),
           "bakeoff3.pdf": Counter(content_list=526, model_json=180, narrow=41, seam=6),
           "textpdf1.pdf": Counter(content_list=428, model_json=30, narrow=144)}
MISSED = {"bakeoff.pdf": ["IR-камерами", "A.II. Taipov P.P. Hypeeb A.P.", "A.III."], "bakeoff2.pdf": ["+3."],
          "bakeoff3.pdf": [],
          "textpdf1.pdf": ["+М0601 i 30, i 35, i 40-SS МИ ВДА/12Я МИ ВДА/12ЯС DIS2116 М0808",
                           "+Федеральное государственное унитарное"]}
COVERED = {"bakeoff.pdf": 7, "bakeoff2.pdf": 15, "bakeoff3.pdf": 7, "textpdf1.pdf": 6}


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("тест полез в сеть или запустил процесс")
    monkeypatch.setattr(socket, "socket", refuse)
    monkeypatch.setattr(subprocess, "run", refuse)
    for name in ("CLAUDE_CODE_LIVE", "VERIFIER", "GEMINI_LIVE"):
        monkeypatch.delenv(name, raising=False)


# --- чистые функции ---------------------------------------------------------

def test_is_cosmetic():
    # пары из v3 и из пробы
    assert is_cosmetic(["помещений-"], ["помещений"]) and is_cosmetic(["°C:"], ["°С:"])
    assert is_cosmetic(["M0601"], ["М0601"]) and is_cosmetic(["C;–"], ["C;"]) and is_cosmetic(["II–"], ["II-"])
    assert is_cosmetic(["MB150;", "C;–", "H4,"], ["МВ150;", "С;–", "Н4,"]) and is_cosmetic(["\\-"], [])
    assert not is_cosmetic(["IP65не"], ["IP65", "не"]) and not is_cosmetic(["сыручими"], ["сыпучими"])
    assert not is_cosmetic(["комплексом).В"], ["комплексом).", "В"])                      # решение 3
    assert not is_cosmetic(["организации.-Произвести"], ["организации.", "-Произвести"])    # c10 — остаётся
    assert not is_cosmetic([], ["Значение"]) and not is_cosmetic(["IR-камерами"], ["IP-камерами"])


def test_plain_tokens():
    md = "## Заголовок\n\n| № | Имя |\n|---|---|\n| 1 | \\- датчик ![](images/a.jpg) |\n"
    tokens, spans = plain_tokens(md)
    assert tokens == ["Заголовок", "№", "Имя", "1", "-", "датчик"]
    assert [md[s:e] for s, e in spans] == ["Заголовок", "№", "Имя", "1", "\\-", "датчик"]


def test_limit_pages():
    h = lambda p, n: Hint(n, n + 1, "x", p, "x")
    kept, flooded = limit_pages([h(2, n) for n in range(PAGE_HINT_LIMIT + 1)] + [h(3, 99)])
    assert flooded == {2: 16} and kept == [h(3, 99)]
    assert limit_pages([h(2, n) for n in range(PAGE_HINT_LIMIT)])[1] == {}                  # 15 — ещё можно


def test_make_verifier_without_call():
    verifier = make_verifier("claude-opus-5")
    assert isinstance(verifier, ClaudeCodeVerifier)
    assert verifier.model == "claude-opus-5" and verifier._live is True


# --- фейки --------------------------------------------------------------------

def bounds(case) -> tuple[int, int]:
    """(lo, hi) окна в plain-токенах: id хранит начало пролёта."""
    lo = int(case.id[1:]) - case.span[0]
    return lo, lo + len(case.fragment.split())


class Fake:
    """Verifier: полоса -> готовый текст. Вызовы пишутся; fail — {номер вызова (1-based): исключение}."""
    name, model = "fake", "fake-model"

    def __init__(self, texts: dict, fail: dict | None = None):
        self.texts, self.fail, self.calls = texts, fail or {}, 0

    def transcribe(self, images, question):
        assert question == TRANSCRIBE_QUESTION
        self.calls += 1
        if self.calls in self.fail:
            raise self.fail[self.calls]
        return [self.texts[image] for image in images]


def golden_texts(cases, tokens, g) -> dict:
    """«Эталон», точный: полоса -> эталонный текст всех окон, у которых она в tiles или next_tiles."""
    opcodes = SequenceMatcher(None, tokens, g, autojunk=False).get_opcodes()

    def to_b(i):            # как в ocr.measure.build_cases
        if i == len(tokens):
            return len(g)
        if i == 0:
            return 0
        tag, i1, _, j1, _ = next(op for op in opcodes if op[1] <= i < op[2])
        return j1 + (i - i1) if tag == "equal" else j1

    spans = {}
    for case in cases:
        lo, hi = bounds(case)
        for tile in case.tiles + case.next_tiles:
            old = spans.get(tile, (lo, hi))
            spans[tile] = (min(old[0], lo), max(old[1], hi))
    return {tile: " ".join(g[to_b(lo):to_b(hi)]) for tile, (lo, hi) in spans.items()}


def echo_texts(cases, tokens) -> dict:
    """«Эхо», реалистичная полоса: у кортежа полос блока — свой отрезок токенов на полосу, с перекрытием."""
    ranges = {}
    for case in cases:
        lo, hi = bounds(case)
        for tiles in filter(None, (case.tiles, case.next_tiles)):
            old = ranges.get(tiles, (lo, hi))
            ranges[tiles] = (min(old[0], lo), max(old[1], hi))
    ov = 2 * (VISION_STRIDE + 2 * CONTEXT_TOKENS)
    parts = {}
    for tiles, (lo, hi) in ranges.items():
        n, length = len(tiles), hi - lo
        for k, tile in enumerate(tiles):
            a, b = max(lo, lo + k * length // n - ov), min(hi, lo + (k + 1) * length // n + ov)
            parts.setdefault(tile, []).append(" ".join(tokens[a:b]))
    return {tile: " ¶ ".join(texts) for tile, texts in parts.items()}


# --- четыре PDF-фикстуры: состав окон, «эхо», «эталон» ---------------------------

@pytest.fixture(scope="module")
def windows(tmp_path_factory):
    for name in PDFS:
        require_fixture(name), require_fixture(ZIPS[name]), require_fixture(GOLDEN[name])
    _, outputs = build_board(tmp_path_factory.mktemp("board"))
    built = {}
    for name in PDFS:
        tokens = plain_tokens(outputs[name][0])[0]
        content_list = read_raw(ZIPS[name])[1]
        pages = model_pages(FIXTURES / ZIPS[name])
        cases, skipped = build_windows(tokens, (FIXTURES / name).read_bytes(), content_list, pages)
        built[name] = tokens, cases, skipped
    return built


@pytest.mark.parametrize("name", PDFS)
def test_windows_composition(windows, name):
    tokens, cases, skipped = windows[name]
    assert len(range(0, len(tokens), VISION_STRIDE)) == WINDOWS[name]
    assert len(cases) == PLACED[name] and dict(skipped) == SKIPPED[name]
    assert Counter(c.page_source for c in cases) == SOURCES[name]
    assert all(c.tiles for c in cases) and all(bool(c.next_tiles) == (c.page_source == "seam") for c in cases)


@pytest.mark.parametrize("name", PDFS)
def test_echo_gives_no_hints(windows, name):
    tokens, cases, _ = windows[name]
    assert hints(run_measure(cases, Fake(echo_texts(cases, tokens)))["cases"]) == []


@pytest.mark.parametrize("name", PDFS)
def test_golden_hints_only_at_golden_opcodes(windows, name):
    tokens, cases, _ = windows[name]
    g = plain_tokens((FIXTURES / GOLDEN[name]).read_text(encoding="utf-8"))[0]
    opcodes = SequenceMatcher(None, tokens, g, autojunk=False).get_opcodes()
    found = hints(run_measure(cases, Fake(golden_texts(cases, tokens, g)))["cases"])
    ops = [op for op in opcodes if op[0] != "equal"]
    near = lambda h, op: _touches(h.start, h.end, op[1] - (op[1] == op[2]), op[2] + (op[1] == op[2]))
    assert all(any(near(h, op) for op in ops) for h in found)                    # ложных — 0
    want = [op for op in ops if not is_cosmetic(tokens[op[1]:op[2]], g[op[3]:op[4]])]
    missed = [" ".join(tokens[op[1]:op[2]]) or "+" + " ".join(g[op[3]:op[4]]) for op in want
              if not any(near(h, op) for h in found)]
    assert missed == MISSED[name]
    assert len(want) - len(missed) == COVERED[name]
    # каждый непокрытый касается окна без полос — это «не привязано», а не «не увидели»
    placed = {int(c.id[1:]) for c in cases}
    for op in want:
        if not any(near(h, op) for h in found):
            lo, hi = op[1] - (op[1] == op[2]), op[2] + (op[1] == op[2])
            spans = [i for i in range(0, len(tokens), VISION_STRIDE) if _touches(i, i + VISION_STRIDE, lo, hi)]
            assert any(i not in placed for i in spans), (name, op)


def test_office_inputs_refuse_vision(tmp_path):
    for name in ("docx1.docx", "xlsx1.xlsx"):
        with pytest.raises(ValueError, match="только для маршрутов MinerU"):
            ingest(require_fixture(name).read_bytes(), source_name=name, work_dir=tmp_path,
                   cache=LocalCache(tmp_path / "cache"), vision=VISION_MODEL)


# --- vision_findings и ingest на textpdf1 -------------------------------------------

def offline(model_version):
    raise AssertionError(f"провайдер {model_version} создан: промах кэша")


def seed(root: Path) -> LocalCache:
    """Как seeded_cache в tests/test_ocr_ingest.py, только textpdf1."""
    cache = LocalCache(root)
    pdf_path = require_fixture("textpdf1.pdf")
    pdf, zip_bytes = pdf_path.read_bytes(), require_fixture(ZIPS["textpdf1.pdf"]).read_bytes()
    key = cache_key(pdf, "mineru", "vlm")
    cache.put(key, zip_bytes, build_meta(key, pdf, zip_bytes, provider="mineru", model_version="vlm",
                                         page_range=None, pages=get_pdf_page_count(str(pdf_path))))
    return cache


@pytest.fixture(scope="module")
def textpdf1(tmp_path_factory):
    """(pdf, zip, md0, отчёт без сверки, cases, echo, cache) — фейки строятся тем же build_windows."""
    tmp = tmp_path_factory.mktemp("textpdf1")
    cache = seed(tmp / "cache")
    pdf, zip_bytes = require_fixture("textpdf1.pdf").read_bytes(), require_fixture(ZIPS["textpdf1.pdf"]).read_bytes()
    md0, report0 = ingest(pdf, source_name="textpdf1.pdf", work_dir=tmp / "w", cache=cache,
                          provider_factory=offline)
    tokens = plain_tokens(md0)[0]
    cases, _ = build_windows(tokens, pdf, read_raw(ZIPS["textpdf1.pdf"])[1], model_pages(FIXTURES / ZIPS["textpdf1.pdf"]))
    return dict(pdf=pdf, zip=zip_bytes, md0=md0, report0=report0, tokens=tokens, cases=cases,
                echo=echo_texts(cases, tokens), cache=cache, tmp=tmp)


def use(monkeypatch, fake):
    made = []
    monkeypatch.setattr(vision_module, "make_verifier", lambda model: made.append(model) or fake)
    return made


def run_ingest(t, **kw):
    return ingest(t["pdf"], source_name="textpdf1.pdf", work_dir=t["tmp"] / "w", cache=t["cache"],
                  provider_factory=offline, **kw)


def vision_of(report):
    return [f for f in report["findings"] if f["rule"].startswith("vision_")]


def one_edit(t):
    """«Одна правка»: w — первый токен пролёта первого подходящего текстового окна; в его полосах w -> wъ."""
    case = next(c for c in t["cases"] if c.block_type == "text"
                and (w := c.fragment.split()[c.span[0]]).isalpha() and len(w) >= 6)
    w = case.fragment.split()[case.span[0]]
    texts = dict(t["echo"])
    for tile in case.tiles + case.next_tiles:
        texts[tile] = " ".join(w + "ъ" if token == w else token for token in texts[tile].split(" "))
    return case, w, texts


def test_echo_on_textpdf1(textpdf1, monkeypatch):
    use(monkeypatch, Fake(textpdf1["echo"]))
    md, report = run_ingest(textpdf1, vision="fake-model")
    assert md == textpdf1["md0"]                                                 # текст не меняется
    found = vision_of(report)
    assert [f["rule"] for f in found] == ["vision_skipped"]
    assert "из 729" in found[0]["suggestion"] and "не привязаны к скану — 127" in found[0]["suggestion"]
    assert report["schema_version"] == 2
    assert all(list(f) == KEYS for f in report["findings"])
    assert all(f["reading"] is None and f["model"] is None for f in report["findings"]
               if not f["rule"].startswith("vision_"))
    # vision=None — прежние находки: те же rule / page / snippet / suggestion
    rest = [(f["rule"], f["page"], f["snippet"], f["suggestion"]) for f in report["findings"]
            if not f["rule"].startswith("vision_")]
    assert rest == [(f["rule"], f["page"], f["snippet"], f["suggestion"]) for f in textpdf1["report0"]["findings"]]
    assert textpdf1["report0"]["schema_version"] == 2 and all(list(f) == KEYS for f in textpdf1["report0"]["findings"])


def test_one_edit_gives_vision_diff(textpdf1, monkeypatch):
    case, w, texts = one_edit(textpdf1)
    use(monkeypatch, Fake(texts))
    md, report = run_ingest(textpdf1, vision="fake-model", annotate=False)
    assert md == textpdf1["md0"]
    diffs = [f for f in report["findings"] if f["rule"] == "vision_diff"]
    assert diffs and all(f["snippet"] == w and f["suggestion"] == w + "ъ" and f["model"] == "fake-model"
                         and w + "ъ" in f["reading"] and f["severity"] == "warning" for f in diffs)
    assert diffs[0]["page"] == case.page

    md_annotated, report = run_ingest(textpdf1, vision="fake-model", annotate=True)
    note = next(f for f in report["findings"] if f["rule"] == "vision_diff")
    blocks = md_annotated.split("\n\n")
    mark = blocks.index(f"!! ПРОВЕРИТЬ: [{note['id']}] vision_diff: {w}ъ !!")
    assert any(w in block for block in blocks[:mark] if not block.startswith("!! ПРОВЕРИТЬ"))
    assert w in [b for b in blocks[:mark] if not b.startswith("!! ПРОВЕРИТЬ")][-1]      # сразу после блока с w
    assert strip_annotations(md_annotated) == textpdf1["md0"]


def test_progress(textpdf1, monkeypatch):
    use(monkeypatch, Fake(textpdf1["echo"]))
    calls = []
    run_ingest(textpdf1, vision="fake-model", vision_progress=lambda *args: calls.append(args))
    keys = list(page_tiles(textpdf1["cases"]))
    assert len(keys) == 10
    assert calls == [(page, done, 10) for done, (_, page) in enumerate(keys, 1)]


def test_failures_become_vision_skipped(textpdf1, monkeypatch, capsys):
    keys = [page for _, page in page_tiles(textpdf1["cases"])]
    echo = textpdf1["echo"]

    use(monkeypatch, Fake(echo, {1: ClaudeCodeMissingError("не найден claude (Claude Code) в PATH")}))
    md, report = run_ingest(textpdf1, vision="fake-model")
    assert md == textpdf1["md0"] and not [f for f in report["findings"] if f["rule"] == "vision_diff"]
    stopped = [f for f in vision_of(report) if "остановлена" in f["suggestion"]]
    assert len(stopped) == 1 and stopped[0]["page"] == keys[0]
    assert stopped[0]["suggestion"] == (f"сверка остановлена на стр. {keys[0]}: ClaudeCodeMissingError: не найден "
                                        f"claude (Claude Code) в PATH; не сверены стр. " + ", ".join(map(str, keys)))
    assert not [f for f in vision_of(report) if "не сверена" in f["suggestion"]]

    use(monkeypatch, Fake(echo, {3: ClaudeCodeLimitError("лимит подписки Claude Code: usage limit")}))
    md, report = run_ingest(textpdf1, vision="fake-model")
    assert md == textpdf1["md0"]
    stopped = [f for f in vision_of(report) if "остановлена" in f["suggestion"]]
    assert len(stopped) == 1 and stopped[0]["page"] == keys[2]
    assert stopped[0]["suggestion"].startswith(f"сверка остановлена на стр. {keys[2]}: ClaudeCodeLimitError: ")
    assert stopped[0]["suggestion"].endswith("; не сверены стр. " + ", ".join(map(str, keys[2:])))
    assert not [f for f in vision_of(report) if "не сверена" in f["suggestion"]]

    use(monkeypatch, Fake(echo, {4: ClaudeCodeTimeoutError("claude -p: нет ответа за 600 с (3 картинок)")}))
    md, report = run_ingest(textpdf1, vision="fake-model")
    assert md == textpdf1["md0"]
    failed = [f for f in vision_of(report) if "не сверена" in f["suggestion"]]
    assert [f["suggestion"] for f in failed] == [
        f"стр. {keys[3]} не сверена: ClaudeCodeTimeoutError: claude -p: нет ответа за 600 с (3 картинок)"]
    assert failed[0]["page"] == keys[3] and not [f for f in vision_of(report) if "остановлена" in f["suggestion"]]

    use(monkeypatch, Fake(echo, {2: RuntimeError("boom")}))
    capsys.readouterr()
    md, report = run_ingest(textpdf1, vision="fake-model")
    assert md == textpdf1["md0"]
    assert [(f["page"], f["suggestion"]) for f in vision_of(report)] == [(None, "сверка не выполнена: RuntimeError: boom")]
    assert "Traceback" in capsys.readouterr().err

    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(textpdf1["zip"])) as source, zipfile.ZipFile(buffer, "w") as target:
        for item in source.namelist():
            if not item.endswith("_model.json"):
                target.writestr(item, source.read(item))
    made = use(monkeypatch, Fake(echo))
    found = vision_findings(textpdf1["md0"], textpdf1["pdf"], buffer.getvalue(), source_name="textpdf1.pdf",
                            model="fake-model")
    assert len(found) == 1 and found[0].rule == "vision_skipped"
    assert found[0].suggestion.startswith("сверка не выполнена: ValueError: ") and made == []


def test_ingest_refuses_vision(textpdf1, tmp_path):
    with pytest.raises(ValueError, match="нужен кэш"):
        ingest(textpdf1["pdf"], source_name="textpdf1.pdf", work_dir=tmp_path, vision="fake-model")
    with pytest.raises(ValueError, match="только для маршрутов MinerU"):
        ingest(textpdf1["pdf"], source_name="textpdf1.pdf", work_dir=tmp_path, engine="ocrmypdf",
               cache=textpdf1["cache"], vision="fake-model")


# --- CLI --------------------------------------------------------------------------

def test_parse_args_vision():
    assert _parse_args(["a.pdf", "--out", "o"]).vision is None
    assert _parse_args(["a.pdf", "--out", "o", "--vision"]).vision == VISION_MODEL
    assert _parse_args(["a.pdf", "--out", "o", "--vision", "claude-sonnet-5"]).vision == "claude-sonnet-5"


def test_main_with_vision(textpdf1, tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)                     # .cache/ocr — во временной папке
    seed(Path(".cache") / "ocr")
    made = use(monkeypatch, Fake(textpdf1["echo"]))
    out = tmp_path / "out"
    code = main([str(require_fixture("textpdf1.pdf")), "--out", str(out), "--vision"], provider_factory=offline)
    captured = capsys.readouterr()
    assert code == 0 and captured.out.count("\n") == 1 and json.loads(captured.out)["error"] is None
    report = json.loads((out / "report.json").read_text(encoding="utf-8"))
    assert report["schema_version"] == 2 and any(f["rule"] == "vision_skipped" for f in report["findings"])
    assert made == [VISION_MODEL]
    assert "сверка по картинке: стр. 1 (1 из 10)" in captured.err
    assert f"сверка по картинке ({VISION_MODEL}): стр. 10, вызовов claude 0, подсказок 0" in captured.err
