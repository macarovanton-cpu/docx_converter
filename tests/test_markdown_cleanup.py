import unittest

from markdown_cleanup import cleanup_ocr_markdown


class MarkdownCleanupTests(unittest.TestCase):
    def test_replaces_latin_ooo_with_cyrillic_ooo(self):
        self.assertEqual(
            cleanup_ocr_markdown("OOO Ромашка"),
            "ООО Ромашка",
        )

    def test_replaces_ip_rating_artifacts(self):
        self.assertEqual(
            cleanup_ocr_markdown("Корпус [Р68, шкаф [P67"),
            "Корпус IP68, шкаф IP67",
        )

    def test_converts_common_ocr_bullet_marker(self):
        self.assertEqual(cleanup_ocr_markdown("e Первый пункт"), "- Первый пункт")

    def test_collapses_extra_blank_lines(self):
        self.assertEqual(cleanup_ocr_markdown("a\n\n\n\nb"), "a\n\nb")

    def test_removes_form_feed(self):
        self.assertEqual(cleanup_ocr_markdown("a\f\nb"), "a\nb")

    def test_converts_inline_semicolon_ocr_bullet_marker(self):
        self.assertEqual(
            cleanup_ocr_markdown("Раздел;e Первый пункт"),
            "Раздел\n- Первый пункт",
        )

    def test_converts_stuck_e_bullet_after_number(self):
        self.assertEqual(
            cleanup_ocr_markdown("+5.000e Новый пункт"),
            "+5.000\n- Новый пункт",
        )

    def test_normalizes_obvious_no_number_marker(self):
        self.assertEqual(
            cleanup_ocr_markdown("Документ No 12"),
            "Документ № 12",
        )


if __name__ == "__main__":
    unittest.main()
