"""ПРАВКА #65: приёмочные тесты сверки прогонов vlm/pipeline.

Фикстурные тесты пропускаются через require_fixture, если нет _test/fixtures/ocr/.
Сети спека не требует.
"""

from ocr_fixtures import read_fixture
from ocr.diff import SNIPPET_MAX, diff_findings, normalize
from ocr.postprocess import postprocess
from ocr.validate import annotate, build_report, strip_annotations

# ПРАВКА #70: число находок low_confidence на паре фикстур. Было 212 при
# пословном выравнивании: три четверти давали потерянные на переносах пробелы.
# Посимвольная сверка их не видит, остаток — 65 настоящих расхождений прогонов.
# Вход фиксированный, поэтому число точное, а не граница.
FIXTURE_FINDINGS = 65


def run_pair():
    a = postprocess(read_fixture("vlm.md"))[0]
    b = postprocess(read_fixture("pipeline.md"))[0]
    return a, b, diff_findings(a, b)


def test_known_defects_highlighted():
    _, _, fs = run_pair()
    syr = [f for f in fs if "сыручими" in f.snippet]
    assert len(syr) == 1 and "сыпучими" in syr[0].suggestion    # второй прогон видел правильно
    ir = [f for f in fs if "IR-камерами" in f.snippet]
    assert len(ir) == 1 and normalize("IP-камерами") in normalize(ir[0].suggestion)
    assert any("Taipov" in f.snippet for f in fs)               # блок подписей
    assert any("Hypeeb" in f.snippet for f in fs)


def test_finding_contract():
    a, b, fs = run_pair()
    assert all((f.rule, f.severity) == ("low_confidence", "warning") for f in fs)
    assert all(f.page is None for f in fs)          # content_list в фикстурах нет
    assert all(f.snippet and f.snippet in a for f in fs)
    assert all(len(f.snippet) <= SNIPPET_MAX for f in fs)
    assert all(f.suggestion is None or len(f.suggestion) <= SNIPPET_MAX for f in fs)
    assert diff_findings(a, a) == [] and diff_findings(b, b) == []


def test_noise_level():
    a, b, fs = run_pair()
    assert len(fs) == FIXTURE_FINDINGS
    # совпадающее не шумит: ОГРН оба прогона прочитали одинаково
    assert not [f for f in fs if "1020202283287" in f.snippet]


def test_findings_fit_report_and_annotations():
    a, _, fs = run_pair()
    report = build_report(source="bakeoff.pdf", sha256="0" * 64, provider="mineru",
                          model_version="vlm", cache_hit=True, verified=True,
                          findings=fs)
    assert strip_annotations(annotate(a, report)) == a


def test_page_from_content_list():
    a, b, _ = run_pair()
    cl = [{"type": "table", "table_body": "<td>твердыми, сыручими и жидкими</td>",
           "page_idx": 1}]
    assert [f.page for f in diff_findings(a, b, cl) if "сыручими" in f.snippet] == [2]


def test_normalize():
    """ПРАВКА #70: пробелов в результате нет — сверяются последовательности символов."""
    assert normalize("No384-ФЗ") == normalize("№ 384-ФЗ") == "№384фз"
    assert normalize("Noп/п") == normalize("№ п/п")
    assert normalize("IР54") == normalize("IP54")               # кир. Р против лат. P
    assert normalize("«НПФ»  БЗК") == normalize("НПф БЗК")
    assert normalize("Ёлка,  ёж.") == "елкаеж"
    assert normalize("№384") == normalize("№ 384")
    # разбивка на слова больше не различает стороны — ради этого всё и затевалось
    assert normalize("твердыми, сыпучими") == normalize("твердыми,сыпучими")


def test_token_level_alignment():
    f = diff_findings("Груз сыручими и жидкими.", "Груз твердыми,сыпучими и жидкими.")
    assert [(x.snippet, x.suggestion) for x in f] == [("сыручими", "твердыми,сыпучими")]
    assert diff_findings("| 1 | а |\n|---|---|\n| 2 | б |", "1 а 2 б") == []  # разметка не шум
    ins = diff_findings("один три", "один два три")
    assert [(x.snippet, x.suggestion) for x in ins] == [("один", "два")]   # insert → сосед слева
    dele = diff_findings("один два три", "один три")
    assert [(x.snippet, x.suggestion) for x in dele] == [("два", None)]


def test_snippet_cut_on_token_boundary():
    long = diff_findings(" ".join(f"а{i}" for i in range(100)), "совсем другое")
    assert len(long[0].snippet) <= SNIPPET_MAX and not long[0].snippet.endswith(" ")


def test_short_divergence_is_not_a_finding():
    """ПРАВКА #70: расхождение короче трёх символов — джиттер OCR, не находка."""
    assert diff_findings("код и цифра 5", "код в цифра 6") == []
    assert [f.snippet for f in diff_findings("код или цифра", "код либо цифра")] == ["или"]
    # разошёлся только «№» — ни буквы, ни цифры, находки нет
    assert diff_findings("пункт № 12", "пункт 12") == []
