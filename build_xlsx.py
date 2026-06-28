"""Дописывает к книге консолидированного реестра листы аналитики (замена графикам PDF).
Запуск: python3 build_xlsx.py  — берёт последнюю книгу из output/ и logs/consolidated.csv,
а также пишет state/last_metrics.json для explain.py / письма / алертов.
"""
from __future__ import annotations
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd
import analytics

BASE = Path(__file__).resolve().parent
OUT = BASE / "output"
LOGS = BASE / "logs"
CSV = LOGS / "consolidated.csv"

SHEETS = {
    "Свод_подразделения": "by_unit",
    "Свод_даты": "by_date",
    "ABC_водители": "drivers",
    "Парк_марки": "truck_mark",
    "Парк_инв": "truck_inv",
}


def write_sheets(aggr: dict, book_path: Path) -> None:
    with pd.ExcelWriter(book_path, engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as xw:
        for sheet, key in SHEETS.items():
            aggr[key].to_excel(xw, sheet_name=sheet, index=False)


def _report_date() -> str:
    rd = datetime.now().strftime("%Y-%m-%d")
    meta = LOGS / "last_fetch.meta"
    if meta.exists():
        for line in meta.read_text(encoding="utf-8").splitlines():
            if line.startswith("date="):
                rd = line.split("=", 1)[1].strip()
    return rd


def _latest_book() -> Path | None:
    books = sorted(OUT.glob("Хронометраж_транспортировки_торфов_*.xlsx"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return books[0] if books else None


def main() -> int:
    if not CSV.exists():
        print(f"[ERR] нет {CSV}"); return 2
    book = _latest_book()
    if not book:
        print("[ERR] нет книги реестра в output/"); return 3
    aggr = analytics.compute(CSV, _report_date())
    write_sheets(aggr, book)
    # метрики для explain.py / email_html / алертов
    analytics.dump_metrics(aggr, BASE / "state" / "last_metrics.json")
    print(f"[OK] листы аналитики добавлены: {book}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
