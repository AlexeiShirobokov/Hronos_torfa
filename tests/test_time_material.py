import unittest
from pathlib import Path
import pandas as pd
import consolidate
import analytics

SCRATCH = Path("tests/_tmp"); SCRATCH.mkdir(parents=True, exist_ok=True)


class TestTimeNorm(unittest.TestCase):
    def test_norm_time_variants(self):
        self.assertEqual(consolidate._norm_time("21:00:00"), "21:00")
        self.assertEqual(consolidate._norm_time("1900-01-01 09:00:00"), "09:00")
        self.assertEqual(consolidate._norm_time("9:30"), "09:30")
        self.assertIsNone(consolidate._norm_time(""))
        self.assertIsNone(consolidate._norm_time(None))


class TestShiftMaterial(unittest.TestCase):
    def test_shift_boundaries_07_20(self):
        self.assertEqual(analytics._shift(7), "1 смена")
        self.assertEqual(analytics._shift(8), "1 смена")
        self.assertEqual(analytics._shift(19), "1 смена")
        self.assertEqual(analytics._shift(20), "2 смена")
        self.assertEqual(analytics._shift(6), "2 смена")
        self.assertEqual(analytics._shift(0), "2 смена")

    def _csv(self, path):
        def row(date, peredel, t, mach):
            return {"Передел": peredel, "Дата. Факт": date, "Подразделение": "Эрел",
                    "Ф.И.О. водителя самосвала": "Иванов И.И.", "Обьем работ, м3": mach * 20.0,
                    "Количство машин, шт": mach, "Обьем кузова,м3": 20,
                    "Инв. № транспортировочной единицы": "1001",
                    "Марка транспортировочной единицы": "HD465-7", "Время": t}
        rows = [
            row("2026-06-28", "Транспортировка торфов", "09:00", 5),   # день
            row("2026-06-28", "Транспортировка торфов", "22:00", 3),   # ночь
            row("2026-06-28", "Транспортировка песков", "10:00", 2),   # день, песок
            row("2026-06-28", "Погрузка торфов", "10:00", 9),          # не транспортировка
        ]
        pd.DataFrame(rows).to_csv(path, index=False)

    def test_machine_aggregates_by_material(self):
        csv = SCRATCH / "tm.csv"; self._csv(csv)
        aggr = analytics.compute(csv, "2026-06-28")
        m = analytics.to_metrics(aggr)
        # материалы разделены
        mats = {r["material"] for r in m["mach_by_day"]}
        self.assertEqual(mats, {"Торф", "Песок"})
        # машины по сменам: торф 1 смена (09:00)=5, 2 смена (22:00)=3
        torf_shift = {(r["shift"]): r["machines"] for r in m["mach_by_shift"]
                      if r["material"] == "Торф"}
        self.assertEqual(torf_shift["1 смена"], 5)
        self.assertEqual(torf_shift["2 смена"], 3)
        # «Погрузка торфов» не попала в транспортировку
        self.assertEqual(sum(r["machines"] for r in m["mach_by_day"]), 10)


if __name__ == "__main__":
    unittest.main()
