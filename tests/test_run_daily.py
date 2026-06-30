import json, unittest
from unittest import mock
import run_daily


class TestAlertDedup(unittest.TestCase):
    def setUp(self):
        run_daily.STATE.mkdir(parents=True, exist_ok=True)
        self.sent = run_daily.STATE / "alerts_sent.json"
        if self.sent.exists():
            self.sent.unlink()

    def tearDown(self):
        if self.sent.exists():
            self.sent.unlink()

    def test_alert_sent_once_per_key_per_day(self):
        calls = []
        with mock.patch.object(run_daily, "ALERTS_ENABLED", True), \
             mock.patch.object(run_daily, "_send_alert_email",
                               side_effect=lambda *a, **k: calls.append(a)):
            run_daily.alert({}, "subj", "body", key="idle:Обман")
            run_daily.alert({}, "subj", "body", key="idle:Обман")  # дубль
        self.assertEqual(len(calls), 1)
        data = json.loads(self.sent.read_text(encoding="utf-8"))
        self.assertTrue(any("idle:Обман" in k for k in data))

    def test_alerts_disabled_by_default_suppresses_send(self):
        calls = []
        with mock.patch.object(run_daily, "_send_alert_email",
                               side_effect=lambda *a, **k: calls.append(a)):
            run_daily.alert({}, "subj", "body", key="idle:Обман")
        self.assertEqual(calls, [])  # ALERTS_ENABLED=False → ничего не отправляем
        self.assertFalse(self.sent.exists())  # и состояние анти-спама не пишется


if __name__ == "__main__":
    unittest.main()
