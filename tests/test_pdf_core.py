import dataclasses
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pytest

from file_converter import analyze_pdf_pages, convert_with_markitdown
from ocr_auto_mode import convert_pdf_with_optional_ocr
from ocr_fixtures import read_fixture, require_fixture
from pdf_core import (
    OcrmypdfProvider,
    OcrResult,
    PageInfo,
    pdf_to_markdown,
    pdf_to_markdown_with_status,
)

REPO_DIR = Path(__file__).resolve().parents[1]
TEXT_PDF = REPO_DIR / "test_files" / "sample.pdf"
SCAN_PDF = REPO_DIR / "test_files" / "ocr_sample.pdf"


# ПРАВКА #60: провайдер возвращает OcrResult, а не строку.
def _fake_result(markdown="provider markdown"):
    return OcrResult(markdown=markdown, content_list=None, pages=[],
                     provider="fake", model_version=None, raw_dir=None)


def _write_temp_pdf(pdf_bytes: bytes) -> str:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    try:
        tmp.write(pdf_bytes)
        return tmp.name
    finally:
        tmp.close()


class PdfCoreParityTests(unittest.TestCase):
    """Ядро даёт тот же markdown, что старый путь в app._convert_uploaded_file."""

    @classmethod
    def setUpClass(cls):
        cls.pdf_bytes = TEXT_PDF.read_bytes()

    def test_mode_off_matches_legacy_markitdown_path(self):
        tmp_path = _write_temp_pdf(self.pdf_bytes)
        try:
            legacy_markdown = convert_with_markitdown(tmp_path, page_range=None)
        finally:
            os.unlink(tmp_path)

        markdown, status = pdf_to_markdown_with_status(self.pdf_bytes, mode="off")

        self.assertEqual(markdown, legacy_markdown)
        self.assertIsNone(status)
        self.assertEqual(pdf_to_markdown(self.pdf_bytes, mode="off"), legacy_markdown)

    def test_mode_auto_matches_legacy_optional_ocr_path(self):
        # страницы 1-12 имеют текстовый слой (13-я — image-only),
        # поэтому OCR не запускается и тест не требует ocrmypdf
        tmp_path = _write_temp_pdf(self.pdf_bytes)
        try:
            pages = analyze_pdf_pages(tmp_path)
            legacy_markdown, legacy_status = convert_pdf_with_optional_ocr(
                tmp_path, page_range="1-12", pages=pages)
        finally:
            os.unlink(tmp_path)

        markdown, status = pdf_to_markdown_with_status(
            self.pdf_bytes, page_range="1-12", mode="auto")

        self.assertEqual(markdown, legacy_markdown)
        self.assertEqual(status, legacy_status)
        self.assertEqual(status["status"], "not_needed")


class PdfCoreNoStreamlitTests(unittest.TestCase):
    def test_core_imports_and_runs_without_streamlit(self):
        script = (
            "import sys; sys.modules['streamlit'] = None\n"
            "import pdf_core\n"
            f"md = pdf_core.pdf_to_markdown(open(r'{TEXT_PDF}', 'rb').read(), mode='off')\n"
            "assert md.strip(), 'empty markdown'\n"
        )
        result = subprocess.run(
            [sys.executable, "-c", script],
            cwd=REPO_DIR, capture_output=True, text=True,
        )
        self.assertEqual(result.returncode, 0, result.stderr)


class PdfCoreProviderTests(unittest.TestCase):
    def test_explicit_provider_used_for_image_only_pdf(self):
        calls = []

        class FakeProvider:
            def ocr_pdf(self, pdf_bytes, page_range=None):
                calls.append((pdf_bytes, page_range))
                return _fake_result()

        pdf_bytes = SCAN_PDF.read_bytes()
        markdown, status = pdf_to_markdown_with_status(
            pdf_bytes, page_range="1-2", provider=FakeProvider())

        self.assertEqual(markdown, "provider markdown")
        self.assertEqual(calls, [(pdf_bytes, "1-2")])
        self.assertEqual(status["status"], "applied")
        self.assertIn("OCR применён", status["message"])
        self.assertEqual(status["pages_without_text_layer"], [1, 2])

    def test_explicit_provider_skipped_when_text_layer_present(self):
        class FailingProvider:
            def ocr_pdf(self, pdf_bytes, page_range=None):
                raise AssertionError("provider must not be called")

        markdown, status = pdf_to_markdown_with_status(
            TEXT_PDF.read_bytes(), page_range="1-12", provider=FailingProvider())

        self.assertTrue(markdown.strip())
        self.assertEqual(status["status"], "not_needed")
        self.assertIn("текстовый слой найден", status["message"])

    def test_ocrmypdf_provider_pipeline_and_temp_cleanup(self):
        seen = {}

        def ocr_func(input_path, output_path):
            seen["src_bytes"] = Path(input_path).read_bytes()
            seen["src_path"] = input_path
            Path(output_path).write_bytes(b"%PDF searchable")

        def convert_func(path, page_range=None):
            seen["converted_path"] = path
            seen["page_range"] = page_range
            return "ocr markdown"

        provider = OcrmypdfProvider(
            ocr_func=ocr_func, convert_func=convert_func,
            analyze_func=lambda path: [{"page_number": 1, "has_text_layer": False}])
        result = provider.ocr_pdf(b"%PDF fake scan", page_range="3")

        self.assertIsInstance(result, OcrResult)
        self.assertEqual(result.markdown, "ocr markdown")
        self.assertEqual((result.provider, result.model_version), ("ocrmypdf", None))
        self.assertIsNone(result.content_list)
        self.assertIsNone(result.raw_dir)
        self.assertEqual(seen["src_bytes"], b"%PDF fake scan")
        self.assertEqual(seen["page_range"], "3")
        self.assertNotEqual(seen["src_path"], seen["converted_path"])
        self.assertFalse(os.path.exists(seen["src_path"]))
        self.assertFalse(os.path.exists(seen["converted_path"]))


# ПРАВКА #60: контракт датаклассов и сведения о страницах от OcrmypdfProvider.
def test_ocr_result_and_page_info_are_frozen_with_fixed_fields():
    result = _fake_result()
    with pytest.raises(dataclasses.FrozenInstanceError):
        result.markdown = "x"
    assert [f.name for f in dataclasses.fields(OcrResult)] == [
        "markdown", "content_list", "pages", "provider", "model_version", "raw_dir"]
    assert [f.name for f in dataclasses.fields(PageInfo)] == [
        "number", "has_text_layer", "ocr_applied", "warnings"]


def test_ocrmypdf_provider_reports_pages_from_real_analyze():
    pdf = require_fixture("bakeoff.pdf").read_bytes()
    vlm = read_fixture("vlm.md")

    def ocr_func(src, dst):
        Path(dst).write_bytes(b"%PDF searchable")

    def convert_func(path, page_range=None):
        return vlm

    result = OcrmypdfProvider(ocr_func=ocr_func, convert_func=convert_func).ocr_pdf(pdf)
    assert result.markdown == vlm
    assert [p.number for p in result.pages] == list(range(1, 10))
    assert all(not p.has_text_layer and p.ocr_applied for p in result.pages)
    assert all(p.warnings == () for p in result.pages)

    ranged = OcrmypdfProvider(
        ocr_func=ocr_func, convert_func=convert_func).ocr_pdf(pdf, "2-3")
    assert [p.number for p in ranged.pages] == [2, 3]


def test_status_dict_unchanged_for_ui():
    class FakeProvider:
        def ocr_pdf(self, pdf_bytes, page_range=None):
            return _fake_result()

    md, status = pdf_to_markdown_with_status(SCAN_PDF.read_bytes(), page_range="1-2",
                                             provider=FakeProvider())
    assert md == "provider markdown"
    assert set(status) == {"mode", "status", "message", "pages_without_text_layer"}


if __name__ == "__main__":
    unittest.main()
