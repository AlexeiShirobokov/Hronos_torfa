"""Расчёт агрегатов по консолидированному реестру + плоский metrics-словарь.
Без matplotlib/сети. Используется build_xlsx.py (DataFrame'ы) и
explain.py/email_html.py (metrics).
"""
from __future__ import annotations
import json
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd


def _abc_class(p: float) -> str:
    if p <= 80.0:
        return "A"
    if p <= 95.0:
        return "B"
    return "C"


# материалы транспортировки (передел → краткое имя)
TRANSPORT_PEREDELY = {
    "Транспортировка торфов": "Торф",
    "Транспортировка песков": "Песок",
}
DAY_SHIFT_START = 8    # дневная смена 08:00–19:59
NIGHT_SHIFT_START = 20  # ночная смена 20:00–07:59


def _shift(hour) -> str | None:
    if hour is None or (isinstance(hour, float) and hour != hour):
        return None
    h = int(hour)
    return "Дневная" if DAY_SHIFT_START <= h < NIGHT_SHIFT_START else "Ночная"


def _to_hour(x):
    if pd.isna(x) or x in (None, "", " "):
        return None
    s = str(x).strip()
    if " " in s and ":" in s.split(" ", 1)[1]:
        s = s.split(" ", 1)[1]
    if ":" in s:
        try:
            return int(s.split(":", 1)[0])
        except ValueError:
            return None
    return None


def compute(csv_path: Path, report_date: str | None = None) -> dict:
    df = pd.read_csv(csv_path, low_memory=False)
    if report_date is None:
        report_date = datetime.now().strftime("%Y-%m-%d")

    df["Дата. Факт"] = pd.to_datetime(df["Дата. Факт"], errors="coerce")
    for c in ("Количство машин, шт", "Обьем работ, м3", "Обьем кузова,м3"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Подразделение"] = df["Подразделение"].astype(str).str.strip()
    df["Час"] = df["Время"].apply(_to_hour)

    trans = df[df["Передел"].astype(str).str.strip() == "Транспортировка торфов"].copy()
    trans["Водитель"] = (trans["Ф.И.О. водителя самосвала"].astype(str).str.strip()
                         .replace({"nan": np.nan, "": np.nan}))
    trans = trans[trans["Водитель"].notna()]

    totals = {
        "volume": float(trans["Обьем работ, м3"].sum()),
        "trips": float(trans["Количство машин, шт"].sum()),
        "drivers": int(trans["Водитель"].nunique()),
        "trucks": int(trans["Инв. № транспортировочной единицы"].dropna().nunique()),
        "units": int(trans["Подразделение"].nunique()),
    }
    totals["m3_per_trip"] = round(totals["volume"] / totals["trips"], 1) if totals["trips"] else None

    by_unit = (trans.groupby("Подразделение")
               .agg(Рейсы=("Количство машин, шт", "sum"),
                    Объем_м3=("Обьем работ, м3", "sum"),
                    Водителей=("Водитель", "nunique"),
                    Самосвалов=("Инв. № транспортировочной единицы",
                                lambda s: s.dropna().nunique()),
                    Дней=("Дата. Факт", "nunique")).reset_index())
    by_unit["м3/рейс"] = by_unit["Объем_м3"] / by_unit["Рейсы"].replace(0, np.nan)
    by_unit["Доля_%"] = by_unit["Объем_м3"] / by_unit["Объем_м3"].sum() * 100
    by_unit = by_unit.sort_values("Объем_м3", ascending=False).reset_index(drop=True)

    by_date = (trans.groupby(trans["Дата. Факт"].dt.date)
               .agg(Рейсы=("Количство машин, шт", "sum"),
                    Объем_м3=("Обьем работ, м3", "sum"),
                    Записей=("Водитель", "count")).reset_index()
               .rename(columns={"Дата. Факт": "Дата"}))

    drv = (trans.groupby("Водитель")
           .agg(Рейсы=("Количство машин, шт", "sum"),
                Объем_м3=("Обьем работ, м3", "sum"),
                Дней=("Дата. Факт", "nunique"),
                Подразделение=("Подразделение",
                               lambda s: s.mode().iat[0] if len(s.mode()) else "—"))
           .reset_index())
    drv["м3/рейс"] = drv["Объем_м3"] / drv["Рейсы"].replace(0, np.nan)
    drv = drv.sort_values("Объем_м3", ascending=False).reset_index(drop=True)
    tv = drv["Объем_м3"].sum()
    drv["Доля_%"] = drv["Объем_м3"] / tv * 100 if tv else 0
    drv["Накопл_%"] = drv["Доля_%"].cumsum()
    drv["ABC"] = drv["Накопл_%"].apply(_abc_class)

    truck_mark = (trans.groupby("Марка транспортировочной единицы")
                  .agg(Кузов_м3_сред=("Обьем кузова,м3", "mean"),
                       Рейсы=("Количство машин, шт", "sum"),
                       Объем_м3=("Обьем работ, м3", "sum"),
                       Самосвалов=("Инв. № транспортировочной единицы",
                                   lambda s: s.dropna().nunique())).reset_index())
    truck_mark["м3/рейс"] = truck_mark["Объем_м3"] / truck_mark["Рейсы"].replace(0, np.nan)
    truck_mark["Использование_%"] = truck_mark["м3/рейс"] / truck_mark["Кузов_м3_сред"] * 100
    truck_mark = truck_mark.sort_values("Объем_м3", ascending=False).reset_index(drop=True)

    truck_inv = (trans.dropna(subset=["Инв. № транспортировочной единицы"])
                 .groupby(["Марка транспортировочной единицы",
                           "Инв. № транспортировочной единицы", "Подразделение"])
                 .agg(Кузов_м3=("Обьем кузова,м3", "mean"),
                      Рейсы=("Количство машин, шт", "sum"),
                      Объем_м3=("Обьем работ, м3", "sum"),
                      Дней=("Дата. Факт", "nunique")).reset_index())
    truck_inv["м3/рейс"] = truck_inv["Объем_м3"] / truck_inv["Рейсы"].replace(0, np.nan)
    truck_inv["Исп_%"] = truck_inv["м3/рейс"] / truck_inv["Кузов_м3"] * 100
    truck_inv = truck_inv.sort_values(["Подразделение", "Объем_м3"],
                                      ascending=[True, False]).reset_index(drop=True)

    by_hour = (trans.groupby("Час").agg(Рейсы=("Количство машин, шт", "sum"),
                                        Объем_м3=("Обьем работ, м3", "sum"))
               .reindex(range(24), fill_value=0).reset_index())

    # простои за отчётную дату (чтобы порог «ч/смену» был осмысленным,
    # а не кумулятивом за весь период)
    rd_ts = pd.to_datetime(report_date, errors="coerce")
    idle_src = df[df["Передел"].astype(str).str.strip() == "простой"]
    idle = (idle_src[idle_src["Дата. Факт"] == rd_ts]
            .groupby(["Подразделение"]).size().reset_index(name="Часов простоя"))

    # ── Транспортировка торф+песок: машины по дням / часам / сменам ──
    transport = df[df["Передел"].astype(str).str.strip().isin(TRANSPORT_PEREDELY)].copy()
    transport["Материал"] = transport["Передел"].astype(str).str.strip().map(TRANSPORT_PEREDELY)
    transport["Смена"] = transport["Час"].apply(_shift)

    mach_by_day = (transport.groupby([transport["Дата. Факт"].dt.date, "Материал"])
                   .agg(Машины=("Количство машин, шт", "sum")).reset_index()
                   .rename(columns={"Дата. Факт": "Дата"}))
    mach_by_hour = (transport.dropna(subset=["Час"])
                    .groupby(["Час", "Материал"])
                    .agg(Машины=("Количство машин, шт", "sum")).reset_index())
    mach_by_shift = (transport.dropna(subset=["Смена"])
                     .groupby([transport["Дата. Факт"].dt.date, "Смена", "Материал"])
                     .agg(Машины=("Количство машин, шт", "sum")).reset_index()
                     .rename(columns={"Дата. Факт": "Дата"}))
    # пивоты «час × дата» по материалам (как ручная сводка, но чисто)
    mach_hour_pivot = {}
    for mat in ("Торф", "Песок"):
        sub = transport[(transport["Материал"] == mat) & transport["Час"].notna()]
        if len(sub):
            piv = (sub.pivot_table(index="Час", columns=sub["Дата. Факт"].dt.date,
                                   values="Количство машин, шт", aggfunc="sum")
                   .reindex(range(24)))
            piv.index = [f"{h:02d}:00" for h in piv.index]
            mach_hour_pivot[mat] = piv.reset_index().rename(columns={"index": "Час"})

    return {
        "report_date": report_date,
        "date_min": df["Дата. Факт"].min(),
        "date_max": df["Дата. Факт"].max(),
        "totals": totals,
        "by_unit": by_unit, "by_date": by_date, "by_hour": by_hour,
        "drivers": drv, "truck_mark": truck_mark, "truck_inv": truck_inv,
        "idle": idle, "units": sorted(trans["Подразделение"].unique().tolist()),
        "n_rows": int(len(df)), "n_trans": int(len(trans)),
        "mach_by_day": mach_by_day, "mach_by_hour": mach_by_hour,
        "mach_by_shift": mach_by_shift, "mach_hour_pivot": mach_hour_pivot,
    }


def to_metrics(aggr: dict) -> dict:
    bu = aggr["by_unit"]
    bd = aggr["by_date"].sort_values("Дата")
    tm = aggr["truck_mark"]
    drv = aggr["drivers"]
    idle = aggr["idle"]

    def _d(x):
        if x is None or (hasattr(x, "__class__") and pd.isna(x)):
            return None
        if hasattr(x, "date"):        # pandas Timestamp / datetime → только дата
            return x.date().isoformat()
        if hasattr(x, "isoformat"):   # datetime.date
            return x.isoformat()
        return str(x)

    denom = (tm["Рейсы"] * tm["Кузов_м3_сред"]).sum()
    overall_util = (tm["Объем_м3"].sum() / denom * 100) if denom else None

    abc_counts = drv["ABC"].value_counts().to_dict()

    return {
        "report_date": aggr["report_date"],
        "period": [_d(aggr["date_min"]), _d(aggr["date_max"])],
        "n_rows": aggr["n_rows"], "n_trans": aggr["n_trans"],
        "totals": aggr["totals"],
        "by_unit": [{"unit": r["Подразделение"], "trips": round(float(r["Рейсы"]), 1),
                     "volume": round(float(r["Объем_м3"]), 1),
                     "m3_per_trip": round(float(r["м3/рейс"]), 1) if pd.notna(r["м3/рейс"]) else None,
                     "share_pct": round(float(r["Доля_%"]), 1)}
                    for _, r in bu.iterrows()],
        "by_date": [{"date": _d(r["Дата"]), "volume": round(float(r["Объем_м3"]), 1),
                     "trips": round(float(r["Рейсы"]), 1)}
                    for _, r in bd.iterrows()],
        "truck_util": {
            "overall_pct": round(float(overall_util), 1) if overall_util is not None else None,
            "by_mark": [{"mark": r["Марка транспортировочной единицы"],
                         "util_pct": round(float(r["Использование_%"]), 1) if pd.notna(r["Использование_%"]) else None}
                        for _, r in tm.iterrows()],
        },
        "idle": [{"unit": r["Подразделение"], "hours": int(r["Часов простоя"])}
                 for _, r in idle.iterrows()],
        "abc": {k: int(abc_counts.get(k, 0)) for k in ("A", "B", "C")},
        "mach_by_day": [{"date": _d(r["Дата"]), "material": r["Материал"],
                         "machines": int(round(float(r["Машины"])))}
                        for _, r in aggr["mach_by_day"].iterrows()],
        "mach_by_hour": [{"hour": f"{int(r['Час']):02d}:00", "material": r["Материал"],
                          "machines": int(round(float(r["Машины"])))}
                         for _, r in aggr["mach_by_hour"].iterrows()],
        "mach_by_shift": [{"date": _d(r["Дата"]), "shift": r["Смена"],
                           "material": r["Материал"],
                           "machines": int(round(float(r["Машины"])))}
                          for _, r in aggr["mach_by_shift"].iterrows()],
    }


def dump_metrics(aggr: dict, path: Path) -> dict:
    m = to_metrics(aggr)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")
    return m
