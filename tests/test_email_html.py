import unittest
import email_html
from tests.test_explain import METRICS


class TestEmailHtml(unittest.TestCase):
    def test_html_has_key_blocks(self):
        html = email_html.build_html(METRICS, "Записка прозой.", [])
        self.assertIn("<table", html)
        self.assertIn("Эрел", html)                    # KPI по подразделениям
        self.assertIn("Записка прозой.", html)          # записка
        self.assertIn("отсутствие напряжения", html)    # причины простоя
        self.assertIn("Почасовая", html)                # раздел почасовки
        self.assertIn("2026-06-28", html)

    def test_text_fallback_plain(self):
        txt = email_html.build_text(METRICS, "Записка.", [])
        self.assertNotIn("<table", txt)
        self.assertIn("Эрел", txt)


if __name__ == "__main__":
    unittest.main()
