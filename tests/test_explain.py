import unittest
import explain

METRICS = {
    "report_date": "2026-06-28",
    "partial": False,
    "op_hours": ["08:00", "20:00"],
    "last_hour_kpi": {"hour": "13:00", "hour_to": "14:00", "torf": 4, "pesok": 1},
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
    "plan_fact": [
        {"unit": "Эрел", "cur": 1806, "avg7": 3200, "plan": 3300, "expected": 3100,
         "pct_fact": 55.0, "pct_proj": 94.0},
        {"unit": "Обман", "cur": 300, "avg7": 1000, "plan": 1200, "expected": 800,
         "pct_fact": 25.0, "pct_proj": 67.0},
    ],
    "hourly_unit": [
        {"unit": "Эрел", "shift": "1 смена", "hour": "08:00", "torf": 10, "pesok": 2, "reason": ""},
        {"unit": "Дражный", "shift": "2 смена", "hour": "20:00", "torf": 0, "pesok": 0,
         "reason": "ЕТО. Пересменка"},
    ],
    "mach_dynamics": {
        "Торф": {"dates": ["2026-06-27", "2026-06-28"],
                 "rows": [{"hour": "08:00", "2026-06-27": 10, "2026-06-28": 8},
                          {"hour": "20:00", "2026-06-27": 3, "2026-06-28": 0}]},
        "Песок": {"dates": ["2026-06-27", "2026-06-28"],
                  "rows": [{"hour": "08:00", "2026-06-27": 2, "2026-06-28": 1}]},
        "Песок_склад": {"dates": ["2026-06-27", "2026-06-28"],
                        "rows": [{"hour": "10:00", "2026-06-27": 4, "2026-06-28": 3}]},
    },
    "idle_reasons": [
        {"unit": "Дражный", "reason": "отсутствие напряжения ВЛ-35кВ", "kind": "внеплановый", "hours": 45},
    ],
    "pesok_devices": {
        "Эрел": {"devices": ["СБ-2.1 #713"], "rows": [{"Час": "08:00", "СБ-2.1 #713": 2}]},
    },
    "mach_dynamics_unit": {
        "Эрел": {"Торф": {"dates": ["2026-06-27", "2026-06-28"],
                          "rows": [{"hour": "08:00", "2026-06-27": 10, "2026-06-28": 8}]},
                 "Песок": {"dates": ["2026-06-27", "2026-06-28"],
                           "rows": [{"hour": "08:00", "2026-06-27": 2, "2026-06-28": 1}]},
                 "Песок_склад": {"dates": ["2026-06-27", "2026-06-28"],
                                 "rows": [{"hour": "10:00", "2026-06-27": 4, "2026-06-28": 3}]}},
    },
    "otkatka_unit": [{"unit": "Эрел", "torf": 400, "pesok": 380},
                     {"unit": "Обман", "torf": 700, "pesok": None}],
    "otkatka_dyn": {"dates": ["2026-06-27", "2026-06-28"], "units": ["Эрел", "Обман"],
                    "rows": [{"Дата": "2026-06-27", "Эрел": 390, "Обман": 700},
                             {"Дата": "2026-06-28", "Эрел": 400, "Обман": None}]},
    "idle_dyn": [{"date": "2026-06-27", "planned": 20, "unplanned": 5, "total": 25},
                 {"date": "2026-06-28", "planned": 18, "unplanned": 10, "total": 28}],
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
