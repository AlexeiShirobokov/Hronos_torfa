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
    "by_unit_day": [
        {"unit": "Эрел", "vol_torf": 3330, "vol_pesok": 1806, "mach_torf": 221, "mach_pesok": 103, "idle_h": 0},
        {"unit": "Обман", "vol_torf": 2100, "vol_pesok": 900, "mach_torf": 175, "mach_pesok": 75, "idle_h": 0},
    ],
    "unit_dev": [
        {"unit": "Эрел", "vol": 5136, "vol_base": 4308, "idle": 0, "idle_base": 0},
        {"unit": "Дражный", "vol": 9372, "vol_base": 11732, "idle": 120, "idle_base": 69},
    ],
    "hourly_unit": [
        {"unit": "Эрел", "hour": "08:00", "machines": 10, "reason": ""},
        {"unit": "Дражный", "hour": "20:00", "machines": 0, "reason": "ЕТО. Пересменка"},
    ],
    "idle_reasons": [
        {"unit": "Дражный", "reason": "отсутствие напряжения ВЛ-35кВ", "kind": "внеплановый", "hours": 45},
    ],
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
