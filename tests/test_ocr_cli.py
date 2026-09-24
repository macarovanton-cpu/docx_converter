"""ПРАВКА #66: приёмочные тесты CLI OCR-тракта.

Сети нет: zip собирается в памяти, провайдер подменяется фабрикой-журналом.
Фикстурные тесты пропускаются через require_fixture, если нет _test/fixtures/ocr/.

ПРАВКА #67: у живого бейкоффа метрика относительная — см. test_live_cli_bakeoff.
"""

import hashlib
import io
import json
import re
import zipfile
from pathlib import Path

import docx
import pytest

from convert import convert_md_to_docx
from ocr_fixtures import count_diffs, read_fixture, require_fixture
from ocr.cache import LocalCache
from ocr.cli import main, run_pipeline
from ocr.postprocess import parse_pipe_tables, postprocess
from ocr.validate import strip_annotations
from pdf_core import PageInfo


def make_zip(full_md: str) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("full.md", full_md)
    return buffer.getvalue()


def make_fake(**texts):
    """Фабрика провайдеров + журнал вызовов сети (по одному model_version на вызов)."""
    calls: list[str] = []

    class FakeProvider:
        def __init__(self, mv):
            self.mv = mv

        def page_infos(self, pdf_bytes, page_range=None):
            return [PageInfo(n, False, True) for n in range(1, 10)]

        def fetch_raw_zip(self, pdf_bytes, page_range=None):
            calls.append(self.mv)
            return make_zip(texts[self.mv])

    return FakeProvider, calls


def fixtures():
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    return (pdf, *(read_fixture(name) for name in ("vlm.md", "pipeline.md")))


# --- run_pipeline -----------------------------------------------------------

def test_cache_hit_skips_provider_and_keeps_result(tmp_path):
    pdf, vlm, pipeline = fixtures()
    fake, calls = make_fake(vlm=vlm, pipeline=pipeline)
    kw = dict(source_name="bakeoff.pdf", work_dir=tmp_path / "w",
              cache=LocalCache(tmp_path / "cache"), provider_factory=fake)

    md, report = run_pipeline(pdf, **kw)
    assert calls == ["vlm"] and report["cache_hit"] is False
    md2, report2 = run_pipeline(pdf, **kw)
    assert calls == ["vlm"] and report2["cache_hit"] is True      # провайдер не вызван
    assert md2 == md == postprocess(vlm)[0]
    assert report2["findings"] == report["findings"]              # кэш не меняет результат
    assert (report["provider"], report["model_version"], report["verified"]) == (
        "mineru", "vlm", False)
    assert report["sha256"] == hashlib.sha256(pdf).hexdigest()
    assert report["summary"]["critical"] == 5
    assert (tmp_path / "w" / "raw" / "vlm" / "full.md").exists()

    # результат годится конвертеру без правок convert.py
    assert "<table" not in md and "!! ПРОВЕРИТЬ" not in md


def test_verify_runs_second_model_with_own_cache_key(tmp_path):
    pdf, vlm, pipeline = fixtures()
    fake, calls = make_fake(vlm=vlm, pipeline=pipeline)
    kw = dict(source_name="bakeoff.pdf", work_dir=tmp_path / "w",
              cache=LocalCache(tmp_path / "cache"), provider_factory=fake)

    md_v, report_v = run_pipeline(pdf, verify=True, **kw)
    assert calls == ["vlm", "pipeline"] and report_v["verified"] is True
    rules = [f["rule"] for f in report_v["findings"]]
    assert "low_confidence" in rules
    assert any("сыручими" in f["snippet"] for f in report_v["findings"])
    assert any("IR-камерами" in f["snippet"] for f in report_v["findings"])

    md, _ = run_pipeline(pdf, **kw)
    assert md_v == md                                             # сверка текст не меняет
    run_pipeline(pdf, verify=True, **kw)
    assert calls == ["vlm", "pipeline"]                           # оба уже в кэше


def test_annotate_marks_every_finding(tmp_path):
    pdf, vlm, pipeline = fixtures()
    fake, _ = make_fake(vlm=vlm, pipeline=pipeline)
    kw = dict(source_name="bakeoff.pdf", work_dir=tmp_path / "w",
              cache=LocalCache(tmp_path / "cache"), provider_factory=fake)

    md, _ = run_pipeline(pdf, **kw)
    md_a, report_a = run_pipeline(pdf, annotate=True, **kw)
    assert strip_annotations(md_a) == md
    assert md_a.count("!! ПРОВЕРИТЬ: ") == len(report_a["findings"])


def test_annotate_all_adds_low_confidence(tmp_path):
    """ПРАВКА #70: --annotate обходит low_confidence, --annotate-all вставляет их."""
    pdf, vlm, pipeline = fixtures()
    fake, _ = make_fake(vlm=vlm, pipeline=pipeline)
    kw = dict(source_name="bakeoff.pdf", work_dir=tmp_path / "w",
              cache=LocalCache(tmp_path / "cache"), provider_factory=fake)

    md, _ = run_pipeline(pdf, verify=True, **kw)
    md_a, report = run_pipeline(pdf, verify=True, annotate=True, **kw)
    md_all, _ = run_pipeline(pdf, verify=True, annotate_all=True, **kw)

    low = [f for f in report["findings"] if f["rule"] == "low_confidence"]
    assert low                                                    # сверка нашла расхождения
    assert md_a.count("!! ПРОВЕРИТЬ: ") == len(report["findings"]) - len(low)
    assert md_all.count("!! ПРОВЕРИТЬ: ") == len(report["findings"])
    assert strip_annotations(md_a) == strip_annotations(md_all) == md


def test_without_cache_provider_is_called_every_time(tmp_path):
    pdf, vlm, pipeline = fixtures()
    fake, calls = make_fake(vlm=vlm, pipeline=pipeline)
    kw = dict(source_name="b.pdf", work_dir=tmp_path / "n", cache=None,
              provider_factory=fake)

    run_pipeline(pdf, **kw)
    run_pipeline(pdf, **kw)
    assert calls == ["vlm", "vlm"]


def test_bad_arguments_raise(tmp_path):
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    with pytest.raises(ValueError, match="verify"):
        run_pipeline(pdf, source_name="b.pdf", work_dir=tmp_path,
                     engine="ocrmypdf", verify=True)
    with pytest.raises(ValueError):
        run_pipeline(pdf, source_name="b.pdf", work_dir=tmp_path, engine="tesseract")
    with pytest.raises(ValueError):
        run_pipeline(pdf, source_name="b.pdf", work_dir=tmp_path, mode="best")


def test_out_md_converts_to_docx(tmp_path):
    pdf, vlm, pipeline = fixtures()
    fake, _ = make_fake(vlm=vlm, pipeline=pipeline)
    md, _ = run_pipeline(pdf, source_name="bakeoff.pdf", work_dir=tmp_path / "w",
                         provider_factory=fake)
    convert_md_to_docx(md, str(tmp_path / "t.docx"), template_path=None)
    doc = docx.Document(str(tmp_path / "t.docx"))
    assert max(len(t.rows) for t in doc.tables) >= 50   # таблица ТЗ доехала целиком


# --- main -------------------------------------------------------------------

def test_main_writes_files_and_one_json_line(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)                     # .cache/ не в репозиторий
    pdf, vlm, pipeline = fixtures()
    fake, _ = make_fake(vlm=vlm, pipeline=pipeline)
    src, out = require_fixture("bakeoff.pdf"), tmp_path / "out"

    code = main([str(src), "--out", str(out)], provider_factory=fake)
    stdout = capsys.readouterr().out
    assert code == 2 and stdout.count("\n") == 1
    summary = json.loads(stdout)
    assert list(summary) == ["status", "out_md", "report", "cache_hit",
                             "findings", "error"]
    assert summary["status"] == "findings" and summary["error"] is None
    assert summary["findings"]["critical"] == 5
    assert Path(summary["out_md"]).is_absolute()
    assert Path(summary["out_md"]).read_text(encoding="utf-8") == postprocess(vlm)[0]
    on_disk = json.loads(Path(summary["report"]).read_text(encoding="utf-8"))
    assert on_disk["schema_version"] == 2 and on_disk["source"] == "bakeoff.pdf"     # ПРАВКА #91
    assert all(list(f) == ["id", "rule", "severity", "page", "snippet", "suggestion", "reading", "model"]
               for f in on_disk["findings"])
    # не \uXXXX: отчёт читает человек
    assert "Сиражитдинов" in Path(summary["report"]).read_text(encoding="utf-8")


def test_main_clean_document_exits_zero(tmp_path, monkeypatch, capsys):
    """Эталон: латиницы в ФИО нет, остались только warning/info."""
    monkeypatch.chdir(tmp_path)
    src = require_fixture("bakeoff.pdf")
    fake, _ = make_fake(vlm=read_fixture("golden.md"), pipeline=read_fixture("pipeline.md"))

    code = main([str(src), "--out", str(tmp_path / "g")], provider_factory=fake)
    payload = json.loads(capsys.readouterr().out)
    assert code == 0 and payload["status"] == "ok" and payload["findings"]["critical"] == 0


def test_main_errors_always_report_json_and_write_nothing(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    pdf, vlm, pipeline = fixtures()
    fake, _ = make_fake(vlm=vlm, pipeline=pipeline)
    src = require_fixture("bakeoff.pdf")

    errors = {}
    for name, argv in {
            "no_file": [str(tmp_path / "нет.pdf"), "--out", str(tmp_path / "e1")],
            "drive": [str(src), "--out", str(tmp_path / "e2"), "--cache", "drive"],
            "verify": [str(src), "--out", str(tmp_path / "e3"),
                       "--engine", "ocrmypdf", "--verify"],
            "no_out": [str(src)],                              # ошибка argparse
            "bad_mode": [str(src), "--out", str(tmp_path / "e5"),
                         "--mode", "best"]}.items():
        code = main(argv, provider_factory=fake)               # SystemExit наружу не летит
        payload = json.loads(capsys.readouterr().out)
        assert code == 1 and payload["status"] == "error" and payload["error"]
        assert payload["out_md"] is None and payload["findings"] is None
        errors[name] = payload["error"]
    assert "этапе 8" in errors["drive"]
    assert not (tmp_path / "e2" / "out.md").exists()


def test_main_does_not_leak_api_key(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("MINERU_API_KEY", "k-secret-123")
    pdf, vlm, pipeline = fixtures()
    fake, _ = make_fake(vlm=vlm, pipeline=pipeline)

    main([str(require_fixture("bakeoff.pdf")), "--out", str(tmp_path / "s")],
         provider_factory=fake)
    captured = capsys.readouterr()
    assert "k-secret-123" not in captured.out + captured.err


@pytest.mark.live
def test_live_cli_bakeoff(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    out_dir = tmp_path / "o"
    code = main([str(require_fixture("bakeoff.pdf")), "--out", str(out_dir), "--verify"])
    payload = json.loads(capsys.readouterr().out)
    assert code in (0, 2) and payload["status"] in ("ok", "findings")

    out = Path(payload["out_md"]).read_text(encoding="utf-8")
    golden = read_fixture("golden.md")
    # ПРАВКА #67: vlm не детерминирован, абсолютный порог мерил совпадение с прогоном,
    # из которого собран эталон. Меряем работу тракта: сырой markdown -> out.md.
    raw = (out_dir / "raw" / "vlm" / "full.md").read_text(encoding="utf-8")
    assert count_diffs(out, golden) < count_diffs(raw, golden)
    # жёсткое, от прогона не зависит: что постпроцессор обязан починить — починено
    assert not re.search(r"\bNo\.?[ \t]*(?=\d|п/п)", out)
    assert "$C^" not in out
    assert "РоЕ" not in out                       # кириллические Р, о, Е
    assert len(parse_pipe_tables(out)) == 1
