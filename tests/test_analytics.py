import json, unittest
from pathlib import Path
import pandas as pd
import analytics

SCRATCH = Path("tests/_tmp"); SCRATCH.mkdir(parents=True, exist_ok=True)


def _make_csv(path: Path):
    rows = []
    base = {"Передел": "Транспортировка торфов"}

    def r(date, unit, drv, vol, kuzov, inv, mark, t="08:00"):
        return {**base, "Дата. Факт": date, "Подразделение": unit,
                "Ф.И.О. водителя самосвала": drv, "Обьем работ, м3": vol,
                "Количство машин, шт": vol / 20.0, "Обьем кузова,м3": kuzov,
                "Инв. № транспортировочной единицы": inv,
                "Марка транспортировочной единицы": mark, "Время": t}

    rows += [r("2026-06-27", "Эрел", "Иванов И.И.", 200, 20, "1001", "HD465-7"),
             r("2026-06-28", "Эрел", "Иванов И.И.", 100, 20, "1001", "HD465-7"),
             r("2026-06-28", "Обман", "Петров П.П.", 60, 12, "2001", "БелАЗ 7547")]
    idle = {"Передел": "простой", "Дата. Факт": "2026-06-28", "Подразделение": "Обман",
            "Ф.И.О. водителя самосвала": "", "Обьем работ, м3": "",
            "Количство машин, шт": "", "Обьем кузова,м3": "",
            "Инв. № транспортировочной единицы": "", "Марка транспортировочной единицы": "",
            "Время": "10:00"}
    rows.append(idle)
    pd.DataFrame(rows).to_csv(path, index=False)


class TestMetrics(unittest.TestCase):
    def setUp(self):
        self.csv = SCRATCH / "consolidated.csv"
        _make_csv(self.csv)

    def test_metrics_shape(self):
        aggr = analytics.compute(self.csv, report_date="2026-06-28")
        m = analytics.to_metrics(aggr)
        self.assertEqual(m["report_date"], "2026-06-28")
        self.assertEqual(m["totals"]["units"], 2)
        self.assertTrue(any(u["unit"] == "Эрел" for u in m["by_unit"]))
        dates = {d["date"] for d in m["by_date"]}
        self.assertIn("2026-06-27", dates)
        self.assertIn("2026-06-28", dates)
        self.assertTrue(any(i["unit"] == "Обман" and i["hours"] >= 1 for i in m["idle"]))

    def test_dump_writes_json(self):
        aggr = analytics.compute(self.csv, report_date="2026-06-28")
        out = SCRATCH / "metrics.json"
        analytics.dump_metrics(aggr, out)
        data = json.loads(out.read_text(encoding="utf-8"))
        self.assertIn("by_unit", data)


if __name__ == "__main__":
    unittest.main()
