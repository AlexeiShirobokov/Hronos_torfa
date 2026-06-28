import unittest
from pathlib import Path
import send_email
import email_html
from tests.test_explain import METRICS


class TestBuildMessage(unittest.TestCase):
    def test_html_and_attachment(self):
        env = {"YANDEX_LOGIN": "shirobokov@pskgold.ru"}
        xlsx = Path("/tmp/_se_test.xlsx"); xlsx.write_bytes(b"PK\x03\x04test")
        html = email_html.build_html(METRICS, "Записка.", [])
        text = email_html.build_text(METRICS, "Записка.", [])
        msg = send_email.build_message(env, ["alexeimvc@gmail.com"], xlsx, html, text)
        self.assertEqual(msg["To"], "alexeimvc@gmail.com")
        body = msg.get_body(preferencelist=("html",))
        self.assertIsNotNone(body)
        self.assertIn("Эрел", body.get_content())
        names = [p.get_filename() for p in msg.iter_attachments()]
        self.assertIn("_se_test.xlsx", names)


if __name__ == "__main__":
    unittest.main()
