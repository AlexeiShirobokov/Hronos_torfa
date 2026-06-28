import unittest
import explain

METRICS = {
    "report_date": "2026-06-28",
    "period": ["2026-05-15", "2026-06-28"],
    "totals": {"volume": 30000.0, "trips": 1500.0, "m3_per_trip": 20.0,
               "drivers": 10, "trucks": 8, "units": 2},
    "by_unit": [{"unit": "Эрел", "trips": 1000.0, "volume": 20000.0,
                 "m3_per_trip": 20.0, "share_pct": 66.7},
                {"unit": "Обман", "trips": 500.0, "volume": 10000.0,
                 "m3_per_trip": 20.0, "share_pct": 33.3}],
    "by_date": [{"date": "2026-06-27", "volume": 50000.0, "trips": 2500.0},
                {"date": "2026-06-28", "volume": 30000.0, "trips": 1500.0}],
    "truck_util": {"overall_pct": 90.0,
                   "by_mark": [{"mark": "БелАЗ 7547", "util_pct": 90.0}]},
    "idle": [{"unit": "Дражный", "hours": 120, "baseline": 70.0}],
    "abc": {"A": 4, "B": 3, "C": 3},
}


class TestExplain(unittest.TestCase):
    def test_rule_based_note_mentions_key_facts(self):
        note = explain.rule_based_note(METRICS)
        self.assertIn("2026-06-28", note)
        self.assertIn("Эрел", note)
        self.assertTrue(len(note) > 100)

    def test_anomalies_detected(self):
        al = explain.detect_anomalies(METRICS)
        types = {a["type"] for a in al}
        self.assertIn("body_util", types)
        self.assertIn("idle", types)
        self.assertIn("volume_drop", types)

    def test_no_anomalies_when_healthy(self):
        healthy = {**METRICS,
                   "truck_util": {"overall_pct": 99.0, "by_mark": []},
                   "idle": [],
                   "by_date": [{"date": "2026-06-27", "volume": 30000.0, "trips": 1500.0},
                               {"date": "2026-06-28", "volume": 30000.0, "trips": 1500.0}]}
        self.assertEqual(explain.detect_anomalies(healthy), [])


if __name__ == "__main__":
    unittest.main()
