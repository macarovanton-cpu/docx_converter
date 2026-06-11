import unittest

from markdown_cleanup import cleanup_ocr_markdown


class MarkdownCleanupTests(unittest.TestCase):
    def test_replaces_latin_ooo_with_cyrillic_ooo(self):
        self.assertEqual(
            cleanup_ocr_markdown("Заказчик: OOO Ромашка"),
            "Заказчик: ООО Ромашка",
        )

    def test_replaces_ip_rating_artifacts(self):
        self.assertEqual(
            cleanup_ocr_markdown("Корпус [Р68, шкаф [P67"),
            "Корпус IP68, шкаф IP67",
        )

    def test_converts_common_ocr_bullet_marker(self):
        self.assertEqual(cleanup_ocr_markdown("e Навес"), "- Навес")

    def test_replaces_mixed_language_not_artifact(self):
        self.assertEqual(
            cleanup_ocr_markdown("камера He «слепла»"),
            "камера не «слепла»",
        )

    def test_collapses_extra_blank_lines(self):
        self.assertEqual(cleanup_ocr_markdown("a\n\n\n\nb"), "a\n\nb")

    def test_removes_form_feed(self):
        self.assertEqual(cleanup_ocr_markdown("a\f\nb"), "a\nb")

    def test_converts_inline_semicolon_ocr_bullet_marker(self):
        self.assertEqual(
            cleanup_ocr_markdown("4. Работы;e КЖ часть"),
            "4. Работы\n- КЖ часть",
        )

    def test_adds_line_break_before_glued_technical_heading(self):
        self.assertEqual(
            cleanup_ocr_markdown("текстНавес: закрытого типа"),
            "текст\nНавес: закрытого типа",
        )

    def test_converts_stuck_e_bullet_after_number(self):
        self.assertEqual(
            cleanup_ocr_markdown("+5.000e Помещение управления"),
            "+5.000\n- Помещение управления",
        )

    def test_converts_stuck_e_bullet_after_number_and_letter(self):
        self.assertEqual(
            cleanup_ocr_markdown("1Ce Получение заключения"),
            "1C\n- Получение заключения",
        )


if __name__ == "__main__":
    unittest.main()
