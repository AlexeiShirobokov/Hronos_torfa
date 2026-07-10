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

# короткий формат дат в «Сводном_Реестре» («18 май» вместо «2026-05-18 0:00:00»)
REESTR_SHEET = "Сводный_Реестр"
REESTR_DATE_COLS = ("Дата. Факт", "Дата выдачи наряд-задания")
DATE_FMT = "[$-419]d mmm"        # «18 май» независимо от локали Excel

SHEETS = {
    "KPI_сутки": "by_unit_day",
    "Данные_пески": "pesok_source",          # источник для native-сводной (скрытый)
    "План_факт_пески": "plan_fact",
    "Отклонения": "unit_dev",
    "Причины_простоя": "idle_reasons",
    "Почасовка_сутки": "hourly_unit",
    "Свод_подразделения": "by_unit",
    "Свод_даты": "by_date",
    "ABC_водители": "drivers",
    "Парк_марки": "truck_mark",
    "Парк_инв": "truck_inv",
    "Машины_дни": "mach_by_day",
    "Машины_смены": "mach_by_shift",
}


def write_sheets(aggr: dict, book_path: Path) -> None:
    with pd.ExcelWriter(book_path, engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as xw:
        for sheet, key in SHEETS.items():
            aggr[key].to_excel(xw, sheet_name=sheet, index=False)
        # пивоты «час × дата» по материалам (как ручная сводка, но чисто)
        for mat, piv in aggr.get("mach_hour_pivot", {}).items():
            sheet = f"Машины_час_{mat}"[:31]
            piv.to_excel(xw, sheet_name=sheet, index=False)
        # песок в разрезе промывочных приборов (час × прибор) по подразделениям
        for unit, piv in aggr.get("pesok_devices", {}).items():
            sheet = f"Песок_{unit}"[:31]
            piv.to_excel(xw, sheet_name=sheet, index=False)
        # «Данные_пески» (источник сводной): даты «18 май» — формат наследует сводная
        if "Данные_пески" in xw.book.sheetnames:
            src = xw.book["Данные_пески"]
            hdr = {c.value: c.column_letter for c in src[1]}
            for name in ("Дата выдачи наряд-задания", "Дата. Факт"):
                letter = hdr.get(name)
                if not letter:
                    continue
                for cell in src[letter]:
                    if cell.row > 1 and cell.value is not None:
                        cell.number_format = DATE_FMT
        # даты в «Сводном_Реестре» — короткий формат «18 май» + узкие столбцы
        _format_reestr_dates(xw.book)
        # оставляем видимым только «Сводный_Реестр» (сводную «Сводная_пески» первым листом
        # добавит PivotBuilder), остальные листы прячем — по просьбе Алексея (как в шаблоне)
        for ws in xw.book.worksheets:
            ws.sheet_state = "visible" if ws.title == REESTR_SHEET else "hidden"
        if REESTR_SHEET in xw.book.sheetnames:
            xw.book.active = xw.book.sheetnames.index(REESTR_SHEET)


def _format_reestr_dates(book) -> None:
    """Лист «Сводный_Реестр»: даты «Дата. Факт»/«Дата выдачи наряд-задания» —
    короткий формат «18 май» и узкие столбцы (по просьбе Алексея)."""
    from openpyxl.styles import Alignment
    if REESTR_SHEET not in book.sheetnames:
        return
    ws = book[REESTR_SHEET]
    header = {c.value: c.column_letter for c in ws[1]}
    for name, width in zip(REESTR_DATE_COLS, (11, 13)):
        letter = header.get(name)
        if not letter:
            continue
        ws.column_dimensions[letter].width = width
        for cell in ws[letter]:
            if cell.row == 1:
                cell.alignment = Alignment(wrap_text=True, vertical="center",
                                           horizontal="center")
            elif cell.value is not None:
                cell.number_format = DATE_FMT


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
    aggr = analytics.compute(CSV, _report_date())
    # листы аналитики в книгу больше НЕ пишем: финальную книгу (сводная + реестр)
    # собирает BookBuilder из шаблона. Здесь нужны только метрики для письма/explain.
    analytics.dump_metrics(aggr, BASE / "state" / "last_metrics.json")
    print("[OK] метрики обновлены (state/last_metrics.json)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
