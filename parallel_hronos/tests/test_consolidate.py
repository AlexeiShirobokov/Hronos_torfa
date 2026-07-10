from __future__ import annotations

from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import pandas as pd

from hronos.consolidate import build_freshness_check, build_pivot_expected, build_volume_check, discover_latest_files, normalize_date_series, normalize_registry, norm_time
from hronos.consolidate import FileReport


class TestConsolidate(unittest.TestCase):
    def test_norm_time_variants(self):
        self.assertEqual(norm_time("1900-01-01 09:30:00"), "09:30")
        self.assertEqual(norm_time("7:05:00"), "07:05")
        self.assertEqual(norm_time(0.5), "12:00")

    def test_normalize_date_excel_serial(self):
        normalized = normalize_date_series(pd.Series([46206]))
        self.assertEqual(normalized.iloc[0].date(), date(2026, 7, 3))

    def test_volume_check_by_unit(self):
        df = pd.DataFrame(
            {
                "Подразделение": ["Дражный", "Дражный", "Обман"],
                "Дата. Факт": ["2026-07-02"] * 3,
                "Дата выдачи наряд-задания": ["2026-07-02"] * 3,
                "Время": ["08:00", "09:00", "08:00"],
                "Передел": ["Транспортировка песков"] * 3,
                "Обьем работ, м3": [10, 15, 7],
            }
        )
        normalized = normalize_registry(df, "Хронометраж (Дражный).xlsx")
        reports = [
            FileReport(
                path="a.xlsx",
                base_name="a.xlsx",
                rows=3,
                volume=32,
                volumes_by_unit={"Дражный": 25, "Обман": 7},
            )
        ]
        check = build_volume_check(reports, normalized)
        self.assertEqual(set(check["Статус"]), {"OK"})
        self.assertEqual(float(check.loc[check["Подразделение"] == "Дражный", "Объем_реестр"].iloc[0]), 25)

    def test_pivot_expected_today_sand(self):
        df = pd.DataFrame(
            {
                "Подразделение": ["Дражный", "Дражный", "Обман"],
                "Дата. Факт": ["2026-07-02"] * 3,
                "Дата выдачи наряд-задания": ["2026-07-02", "2026-07-01", "2026-07-02"],
                "Время": ["08:00", "09:00", "08:00"],
                "Передел": ["Транспортировка песков", "Транспортировка песков", "Транспортировка торфов"],
                "Обьем работ, м3": [10, 15, 7],
            }
        )
        normalized = normalize_registry(df, "Хронометраж (Дражный).xlsx")
        expected = build_pivot_expected(normalized, date(2026, 7, 2), ["Транспортировка песков"])
        self.assertEqual(
            expected.to_dict(orient="records"),
            [{"Подразделение": "Дражный", "Ожидаемый_объем": 10, "Переделы": "Транспортировка песков"}],
        )

    def test_peredel_canonicalization(self):
        df = pd.DataFrame(
            {
                "Подразделение": ["Дражный"] * 4,
                "Дата. Факт": ["2026-07-02"] * 4,
                "Дата выдачи наряд-задания": ["2026-07-02"] * 4,
                "Время": ["08:00"] * 4,
                "Передел": ["Транспортировка песков   ", "Подача песков", "простой", "ГПР "],
                "Обьем работ, м3": [10, 20, 30, 40],
            }
        )
        normalized = normalize_registry(df, "Хронометраж (Дражный).xlsx")
        self.assertEqual(
            normalized["Передел"].tolist(),
            ["Транспортировка песков", "Подача песков", "Простой", "ГПР"],
        )

    def test_sand_to_warehouse_remark_rewrites_peredel(self):
        df = pd.DataFrame(
            {
                "Подразделение": ["Дражный", "Дражный"],
                "Дата. Факт": ["2026-07-02", "2026-07-02"],
                "Дата выдачи наряд-задания": ["2026-07-02", "2026-07-02"],
                "Время": ["08:00", "09:00"],
                "Передел": ["Транспортировка песков", "Транспортировка песков"],
                "Примечание": ["Склад", ""],
                "Обьем работ, м3": [10, 20],
            }
        )
        normalized = normalize_registry(df, "Хронометраж (Дражный).xlsx")
        self.assertEqual(
            normalized["Передел"].tolist(),
            ["Транспортировка песков на склад", "Транспортировка песков"],
        )

    def test_freshness_check_marks_stale_required_unit(self):
        reports = [
            FileReport(
                path="Дражный.xlsx",
                base_name="Дражный.xlsx",
                rows=10,
                volumes_by_unit={"Дражный": 100},
                max_fact_date="2026-07-01",
            ),
            FileReport(
                path="Обман.xlsx",
                base_name="Обман.xlsx",
                rows=10,
                volumes_by_unit={"Обман": 100},
                max_fact_date="2026-07-03",
            ),
        ]
        check = build_freshness_check(reports, date(2026, 7, 3), ["Дражный", "Обман"])
        statuses = dict(zip(check["Подразделение"], check["Статус"]))
        self.assertEqual(statuses["Дражный"], "НЕТ СВЕЖЕГО ФАЙЛА")
        self.assertEqual(statuses["Обман"], "OK")

    def test_discover_latest_groups_xlsx_and_xlsb_same_source(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            old = root / "Хронометраж транспортировки торфов и песков (Дражный).xlsx"
            new = root / "Хронометраж транспортировки торфов и песков (Дражный).xlsb"
            old.write_bytes(b"old")
            new.write_bytes(b"new")
            files = discover_latest_files(root)
            self.assertEqual(files, [new])


if __name__ == "__main__":
    unittest.main()
