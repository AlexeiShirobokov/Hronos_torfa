"""Расчёт агрегатов по консолидированному реестру + плоский metrics-словарь.
Без matplotlib/сети. Используется build_xlsx.py (DataFrame'ы) и
explain.py/email_html.py (metrics).
"""
from __future__ import annotations
import json
import re
from pathlib import Path
from datetime import datetime, date
import numpy as np
import pandas as pd


def _device(mark, inv) -> str:
    """Метка промывочного прибора: «Марка #Инв» с нормализацией (СБ2.1→СБ-2.1, без .0)."""
    m = re.sub(r"\s+", " ", str(mark).strip())
    m = re.sub(r"СБ\s*-?\s*", "СБ-", m)        # СБ2.1 / СБ 2.1 → СБ-2.1
    if m.lower() in ("", "nan"):
        m = "—"
    iv = str(inv).strip()
    if iv.endswith(".0"):
        iv = iv[:-2]
    return f"{m} #{iv}" if iv.lower() not in ("", "nan") else m


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


def _idle_kind(note) -> str:
    """Плановый простой (обед, пересменка, ЕТО) vs внеплановый (ремонт/поломка/прочее)."""
    s = str(note).strip().lower()
    if "обед" in s or "пересмен" in s or s.startswith("ето"):
        return "плановый"
    return "внеплановый"


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
    if "Примечание" not in df.columns:
        df["Примечание"] = ""
    for _col in ("Марка промывочного прибора", "Инв. № промывочного прибора"):
        if _col not in df.columns:
            df[_col] = ""

    # отчёт «за сутки» = последний ПОЛНЫЙ день: если отчётная дата = сегодня
    # (неполные сутки), отступаем на последний завершённый день в данных
    today = date.today()
    rd_ts = pd.to_datetime(report_date, errors="coerce")
    rd_date = rd_ts.date() if pd.notna(rd_ts) else today
    if rd_date >= today:
        earlier = [d for d in df["Дата. Факт"].dt.date.dropna().unique() if d < today]
        if earlier:
            rd_date = max(earlier)
    report_date = rd_date.isoformat()
    rd_ts = pd.Timestamp(rd_date)

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

    # простои: за отчётную дату + базовая линия (медиана до 7 дней до неё),
    # чтобы алертить относительный скачок, а не абсолют (rd_date — последний полный день)
    idle_src = df[df["Передел"].astype(str).str.strip() == "простой"].copy()
    idle_src["Дата"] = idle_src["Дата. Факт"].dt.date
    idle_daily = (idle_src.groupby(["Подразделение", "Дата"]).size()
                  .reset_index(name="Часов простоя"))
    idle_rows = []
    for unit, g in idle_daily.groupby("Подразделение"):
        today_h = int(g.loc[g["Дата"] == rd_date, "Часов простоя"].sum())
        prior = g.loc[g["Дата"] < rd_date].sort_values("Дата")["Часов простоя"].tail(7)
        baseline = float(prior.median()) if len(prior) else 0.0
        if today_h > 0 or baseline > 0:
            idle_rows.append({"Подразделение": unit, "Часов простоя": today_h,
                              "База_медиана": round(baseline, 1)})
    idle = pd.DataFrame(idle_rows, columns=["Подразделение", "Часов простоя", "База_медиана"])

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

    # ── Операционный срез за отчётную дату по подразделениям ──
    transport["Дата"] = transport["Дата. Факт"].dt.date
    day_tr = transport[transport["Дата"] == rd_date]

    def _mat_piv(frame, value):
        p = frame.pivot_table(index="Подразделение", columns="Материал",
                              values=value, aggfunc="sum", fill_value=0)
        for mat in ("Торф", "Песок"):
            if mat not in p.columns:
                p[mat] = 0
        return p[["Торф", "Песок"]]

    vol_p = _mat_piv(day_tr, "Обьем работ, м3")
    mach_p = _mat_piv(day_tr, "Количство машин, шт")
    by_unit_day = pd.DataFrame({
        "Подразделение": list(vol_p.index),
        "Объём_торф_м3": vol_p["Торф"].round().astype(int).values,
        "Объём_песок_м3": vol_p["Песок"].round().astype(int).values,
        "Машин_торф": mach_p["Торф"].round().astype(int).values,
        "Машин_песок": mach_p["Песок"].round().astype(int).values,
    })
    idle_today = idle_daily[idle_daily["Дата"] == rd_date].set_index("Подразделение")["Часов простоя"]
    by_unit_day["Простои_ч"] = by_unit_day["Подразделение"].map(idle_today).fillna(0).astype(int)
    by_unit_day = by_unit_day.sort_values("Объём_торф_м3", ascending=False).reset_index(drop=True)

    # отклонения по подразделениям: объём (торф+песок) и простои vs медиана 7 дней
    unit_day_vol = transport.groupby(["Подразделение", "Дата"])["Обьем работ, м3"].sum().reset_index()
    dev_rows = []
    for unit in sorted(transport["Подразделение"].unique()):
        g = unit_day_vol[unit_day_vol["Подразделение"] == unit]
        today_v = float(g.loc[g["Дата"] == rd_date, "Обьем работ, м3"].sum())
        prior_v = g.loc[g["Дата"] < rd_date].sort_values("Дата")["Обьем работ, м3"].tail(7)
        base_v = float(prior_v.median()) if len(prior_v) else 0.0
        gi = idle_daily[idle_daily["Подразделение"] == unit]
        today_i = int(gi.loc[gi["Дата"] == rd_date, "Часов простоя"].sum())
        prior_i = gi.loc[gi["Дата"] < rd_date].sort_values("Дата")["Часов простоя"].tail(7)
        base_i = float(prior_i.median()) if len(prior_i) else 0.0
        dev_rows.append({"Подразделение": unit, "Объём": round(today_v), "Объём_база": round(base_v),
                         "Простои": today_i, "Простои_база": round(base_i)})
    unit_dev = pd.DataFrame(dev_rows)

    # почасовка по подразделениям: машины раздельно торф/песок + причина простоя
    day_all = df[df["Дата. Факт"].dt.date == rd_date].copy()
    day_all["_per"] = day_all["Передел"].astype(str).str.strip()
    day_tr_all = day_all[day_all["_per"].isin(TRANSPORT_PEREDELY)].copy()
    day_tr_all["Материал"] = day_tr_all["_per"].map(TRANSPORT_PEREDELY)
    day_idle = day_all[day_all["_per"] == "простой"].dropna(subset=["Час"])
    mt_h = (day_tr_all[day_tr_all["Материал"] == "Торф"]
            .groupby(["Подразделение", "Час"])["Количство машин, шт"].sum())
    mp_h = (day_tr_all[day_tr_all["Материал"] == "Песок"]
            .groupby(["Подразделение", "Час"])["Количство машин, шт"].sum())
    reason_h = (day_idle.groupby(["Подразделение", "Час"])["Примечание"]
                .agg(lambda s: "; ".join(dict.fromkeys(
                    t for t in (str(x).strip() for x in s) if t and t != "nan"))))
    hourly_rows = []
    for unit in sorted(day_tr_all["Подразделение"].unique()):
        for h in range(24):
            hourly_rows.append({
                "Подразделение": unit, "Час": f"{h:02d}:00",
                "Машин_торф": int(round(float(mt_h.get((unit, h), 0)))),
                "Машин_песок": int(round(float(mp_h.get((unit, h), 0)))),
                "Причина": reason_h.get((unit, h), ""),
            })
    hourly_unit = pd.DataFrame(
        hourly_rows, columns=["Подразделение", "Час", "Машин_торф", "Машин_песок", "Причина"])

    # песок в разрезе промывочных приборов: пивот час × прибор по подразделениям
    pes = day_tr_all[day_tr_all["Материал"] == "Песок"].dropna(subset=["Час"]).copy()
    pes["Прибор"] = [_device(mk, iv) for mk, iv in
                     zip(pes["Марка промывочного прибора"], pes["Инв. № промывочного прибора"])]
    pesok_devices = {}
    for unit in sorted(pes["Подразделение"].unique()):
        sub = pes[pes["Подразделение"] == unit]
        piv = (sub.pivot_table(index="Час", columns="Прибор",
                               values="Количство машин, шт", aggfunc="sum", fill_value=0)
               .reindex(range(24), fill_value=0))
        piv.index = [f"{h:02d}:00" for h in piv.index]
        piv = piv.reset_index().rename(columns={"index": "Час"})
        piv.columns.name = None
        pesok_devices[unit] = piv

    # аналитика причин простоя за сутки (плановые/внеплановые)
    di = day_all[day_all["_per"] == "простой"].copy()
    di["Причина"] = di["Примечание"].astype(str).str.strip().replace({"nan": "—", "": "—"})
    di["Тип"] = di["Примечание"].apply(_idle_kind)
    idle_reasons = (di.groupby(["Подразделение", "Причина", "Тип"]).size()
                    .reset_index(name="Часов").sort_values("Часов", ascending=False)
                    .reset_index(drop=True))

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
        "by_unit_day": by_unit_day, "unit_dev": unit_dev,
        "hourly_unit": hourly_unit, "idle_reasons": idle_reasons,
        "pesok_devices": pesok_devices,
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
        "idle": [{"unit": r["Подразделение"], "hours": int(r["Часов простоя"]),
                  "baseline": float(r["База_медиана"])}
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
        "by_unit_day": [{"unit": r["Подразделение"],
                         "vol_torf": int(r["Объём_торф_м3"]), "vol_pesok": int(r["Объём_песок_м3"]),
                         "mach_torf": int(r["Машин_торф"]), "mach_pesok": int(r["Машин_песок"]),
                         "idle_h": int(r["Простои_ч"])}
                        for _, r in aggr["by_unit_day"].iterrows()],
        "unit_dev": [{"unit": r["Подразделение"], "vol": int(r["Объём"]), "vol_base": int(r["Объём_база"]),
                      "idle": int(r["Простои"]), "idle_base": int(r["Простои_база"])}
                     for _, r in aggr["unit_dev"].iterrows()],
        "hourly_unit": [{"unit": r["Подразделение"], "hour": r["Час"],
                         "torf": int(r["Машин_торф"]), "pesok": int(r["Машин_песок"]),
                         "reason": r["Причина"]}
                        for _, r in aggr["hourly_unit"].iterrows()],
        "idle_reasons": [{"unit": r["Подразделение"], "reason": r["Причина"],
                          "kind": r["Тип"], "hours": int(r["Часов"])}
                         for _, r in aggr["idle_reasons"].iterrows()],
        "pesok_devices": {unit: {"devices": [c for c in piv.columns if c != "Час"],
                                 "rows": piv.to_dict("records")}
                          for unit, piv in aggr["pesok_devices"].items()},
    }


def dump_metrics(aggr: dict, path: Path) -> dict:
    m = to_metrics(aggr)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")
    return m
