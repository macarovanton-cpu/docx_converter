"""ПРАВКА #81: приёмочные тесты единого входа ocr.ingest. Сети нет вообще."""

import hashlib
import json
import socket
from pathlib import Path

import pytest

from file_converter import get_pdf_page_count
from ocr_fixtures import require_fixture
from ocr.cache import LocalCache, build_meta, cache_key
from ocr.cli import main, run_pipeline
from ocr.ingest import detect_route, ingest
from ocr.postprocess import postprocess

SAMPLE_PDF = Path(__file__).resolve().parents[1] / "test_files" / "sample.pdf"
RAW = {"bakeoff.pdf": "vlm_raw.zip", "textpdf1.pdf": "vlm_raw_textpdf1.zip"}
KEYS = ["schema_version", "source", "sha256", "provider", "model_version",
        "cache_hit", "verified", "created_at", "summary", "findings"]


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def refuse(*args, **kwargs):
        raise AssertionError("тест полез в сеть")
    monkeypatch.setattr(socket, "socket", refuse)


def offline(model_version):
    raise AssertionError(f"провайдер {model_version} создан: промах кэша")


def seeded_cache(root: Path) -> LocalCache:
    cache = LocalCache(root)
    for pdf_name, zip_name in RAW.items():
        pdf_path = require_fixture(pdf_name)
        pdf, zip_bytes = pdf_path.read_bytes(), require_fixture(zip_name).read_bytes()
        key = cache_key(pdf, "mineru", "vlm")
        cache.put(key, zip_bytes, build_meta(
            key, pdf, zip_bytes, provider="mineru", model_version="vlm",
            page_range=None, pages=get_pdf_page_count(str(pdf_path))))
    return cache


@pytest.fixture
def two_page_text_pdf(tmp_path) -> Path:
    """Синтетическая вырезка: страницы 1-2 textpdf1.pdf — слой есть, таблиц нет (маршрут text)."""
    from pypdf import PdfReader, PdfWriter

    writer = PdfWriter()
    for page in PdfReader(str(require_fixture("textpdf1.pdf"))).pages[:2]:
        writer.add_page(page)
    path = tmp_path / "text2p.pdf"
    with path.open("wb") as handle:
        writer.write(handle)
    return path


def read(name: str) -> bytes:
    return require_fixture(name).read_bytes()


def test_routes_on_fixtures(two_page_text_pdf):
    for name, route in {"bakeoff.pdf": "scan", "bakeoff2.pdf": "scan", "bakeoff3.pdf": "scan",
                        "textpdf1.pdf": "text_tables",
                        "docx1.docx": "office", "xlsx1.xlsx": "office"}.items():
        assert detect_route(require_fixture(name).read_bytes(), name) == route, name
    # смешанный PDF (13-я из 13 страниц без слоя) -> scan, не text
    assert detect_route(SAMPLE_PDF.read_bytes(), "SAMPLE.PDF") == "scan"
    assert detect_route(two_page_text_pdf.read_bytes(), "text2p.pdf") == "text"
    with pytest.raises(ValueError, match="Неподдерживаемый формат"):
        detect_route(b"x", "a.pptx")


def test_one_report_on_all_routes(tmp_path, two_page_text_pdf):
    cache = seeded_cache(tmp_path / "cache")
    for name, provider in (("bakeoff.pdf", "mineru"), ("textpdf1.pdf", "mineru"),
                           ("text2p.pdf", "markitdown"),
                           ("docx1.docx", "markitdown"), ("xlsx1.xlsx", "markitdown")):
        data = two_page_text_pdf.read_bytes() if name == "text2p.pdf" else read(name)
        md, report = ingest(data, source_name=name, work_dir=tmp_path / name,
                            cache=cache, provider_factory=offline)
        assert list(report) == KEYS and report["provider"] == provider, name
        assert report["sha256"] == hashlib.sha256(data).hexdigest()
        assert set(report["summary"]) == {"critical", "warning", "info"}
        assert md.strip()
        assert postprocess(md)[0] == md, name       # постпроцессор прошёл и идемпотентен


def test_mineru_route_equals_run_pipeline(tmp_path):
    bakeoff = read("bakeoff.pdf")
    kw = dict(source_name="bakeoff.pdf", cache=seeded_cache(tmp_path / "cache"),
              provider_factory=offline)
    md, report = ingest(bakeoff, work_dir=tmp_path / "a", **kw)
    assert md == run_pipeline(bakeoff, work_dir=tmp_path / "b", **kw)[0]
    assert report["cache_hit"] is True               # провайдер не создавался


def test_markitdown_routes(tmp_path):
    md_docx, report_docx = ingest(read("docx1.docx"), source_name="docx1.docx",
                                  work_dir=tmp_path / "d")
    assert report_docx["model_version"] is None and report_docx["cache_hit"] is False
    assert all(item["page"] is None for item in report_docx["findings"])
    md_xlsx, _ = ingest(read("xlsx1.xlsx"), source_name="xlsx1.xlsx", work_dir=tmp_path / "x")
    assert "|" in md_xlsx                            # листы доехали таблицами


def test_verify_and_engine(tmp_path):
    with pytest.raises(ValueError, match="MinerU"):
        ingest(read("docx1.docx"), source_name="docx1.docx", work_dir=tmp_path, verify=True)
    with pytest.raises(ValueError, match="--verify"):
        ingest(read("bakeoff.pdf"), source_name="bakeoff.pdf", work_dir=tmp_path,
               engine="ocrmypdf", verify=True)


def test_cli_office_input(tmp_path, capsys):
    code = main([str(require_fixture("docx1.docx")), "--out", str(tmp_path / "d")])
    payload = json.loads(capsys.readouterr().out)
    assert code in (0, 2) and payload["error"] is None
    assert list(payload) == ["status", "out_md", "report", "cache_hit", "findings", "error"]
    assert (tmp_path / "d" / "out.md").exists() and (tmp_path / "d" / "report.json").exists()
    (tmp_path / "a.pptx").write_bytes(b"x")
    assert main([str(tmp_path / "a.pptx"), "--out", str(tmp_path / "p")]) == 1
    assert not (tmp_path / "p").exists()
