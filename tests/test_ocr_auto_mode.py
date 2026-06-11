import os
import unittest
from pathlib import Path

from ocr_auto_mode import (
    convert_pdf_with_optional_ocr,
    pdf_pages_without_text_layer,
)


class OcrAutoModeTests(unittest.TestCase):
    def test_pdf_pages_without_text_layer_returns_page_numbers(self):
        pages = [
            {"page_number": 1, "has_text_layer": True},
            {"page_number": 2, "has_text_layer": False},
            {"page_number": 3, "has_text_layer": True},
        ]

        self.assertEqual(pdf_pages_without_text_layer(pages), [2])

    def test_skips_ocr_when_all_pages_have_text_layer(self):
        calls = []

        def convert_func(path, page_range=None):
            calls.append(("convert", path, page_range))
            return "markdown"

        def ocr_func(input_path, output_path):
            calls.append(("ocr", input_path, output_path))

        markdown, status = convert_pdf_with_optional_ocr(
            "source.pdf",
            page_range="1-2",
            pages=[{"page_number": 1, "has_text_layer": True}],
            convert_func=convert_func,
            ocr_func=ocr_func,
        )

        self.assertEqual(markdown, "markdown")
        self.assertEqual(calls, [("convert", "source.pdf", "1-2")])
        self.assertEqual(status["status"], "not_needed")
        self.assertIn("текстовый слой найден", status["message"])

    def test_skips_ocr_when_page_range_excludes_image_only_pages(self):
        calls = []

        def convert_func(path, page_range=None):
            calls.append(("convert", path, page_range))
            return "markdown"

        def ocr_func(input_path, output_path):
            calls.append(("ocr", input_path, output_path))

        markdown, status = convert_pdf_with_optional_ocr(
            "source.pdf",
            page_range="4",
            pages=[
                {"page_number": 4, "has_text_layer": True},
                {"page_number": 13, "has_text_layer": False},
            ],
            convert_func=convert_func,
            ocr_func=ocr_func,
        )

        self.assertEqual(markdown, "markdown")
        self.assertEqual(calls, [("convert", "source.pdf", "4")])
        self.assertEqual(status["status"], "not_needed")
        self.assertIn("текстовый слой найден", status["message"])

    def test_applies_ocr_when_page_range_includes_image_only_pages(self):
        calls = []

        def convert_func(path, page_range=None):
            calls.append(("convert", path, page_range))
            return "ocr markdown"

        def ocr_func(input_path, output_path):
            calls.append(("ocr", input_path, output_path))
            Path(output_path).write_bytes(b"%PDF searchable")

        markdown, status = convert_pdf_with_optional_ocr(
            "source.pdf",
            page_range="13",
            pages=[
                {"page_number": 4, "has_text_layer": True},
                {"page_number": 13, "has_text_layer": False},
            ],
            convert_func=convert_func,
            ocr_func=ocr_func,
        )

        self.assertEqual(markdown, "ocr markdown")
        self.assertEqual(calls[0][0], "ocr")
        self.assertEqual(calls[1][0], "convert")
        self.assertEqual(calls[1][2], "13")
        self.assertEqual(status["status"], "applied")
        self.assertIn("OCR применён", status["message"])
        self.assertEqual(status["pages_without_text_layer"], [13])

    def test_applies_ocr_for_pdf_without_text_layer_and_removes_temp_pdf(self):
        converted_paths = []

        def convert_func(path, page_range=None):
            self.assertTrue(os.path.exists(path))
            converted_paths.append(path)
            return "ocr markdown"

        def ocr_func(input_path, output_path):
            self.assertEqual(input_path, "source.pdf")
            Path(output_path).write_bytes(b"%PDF searchable")

        markdown, status = convert_pdf_with_optional_ocr(
            "source.pdf",
            page_range=None,
            pages=[{"page_number": 1, "has_text_layer": False}],
            convert_func=convert_func,
            ocr_func=ocr_func,
        )

        self.assertEqual(markdown, "ocr markdown")
        self.assertEqual(status["status"], "applied")
        self.assertIn("OCR применён", status["message"])
        self.assertEqual(status["pages_without_text_layer"], [1])
        self.assertEqual(len(converted_paths), 1)
        self.assertNotEqual(converted_paths[0], "source.pdf")
        self.assertFalse(os.path.exists(converted_paths[0]))

    def test_wraps_ocr_backend_error(self):
        def convert_func(path, page_range=None):
            return "should not convert"

        def ocr_func(input_path, output_path):
            raise RuntimeError("backend failed")

        with self.assertRaisesRegex(RuntimeError, "OCR auto: ошибка OCR"):
            convert_pdf_with_optional_ocr(
                "source.pdf",
                page_range=None,
                pages=[{"page_number": 1, "has_text_layer": False}],
                convert_func=convert_func,
                ocr_func=ocr_func,
            )


if __name__ == "__main__":
    unittest.main()
