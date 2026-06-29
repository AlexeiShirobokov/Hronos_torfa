"""Расчёт агрегатов по консолидированному реестру + плоский metrics-словарь.
Операционная модель суток: граница 07:00. 1 смена 07:00–19:59, 2 смена 20:00–06:59
(ночная смена относится к суткам, в которые НАЧАЛАСЬ). Отчёт «за сутки» = текущие
операционные сутки нарастающим итогом + выделенный прошедший час.
Без matplotlib/сети. Используется build_xlsx.py (DataFrame'ы) и
explain.py/email_html.py (metrics).
"""
from __future__ import annotations
import json
import re
from pathlib import Path
from datetime import datetime, date, timedelta
import numpy as np
import pandas as pd


def _device(mark, inv) -> str:
    """Метка промывочного прибора: «Марка #Инв» с нормализацией (СБ2.1→СБ-2.1, без .0)."""
    m = re.sub(r"\s+", " ", str(mark).strip())
    m = re.sub(r"СБ\s*-?\s*", "СБ-", m)        # СБ2.1 / СБ 2.1 → СБ-2.1
    if m.lower() in ("", "nan"):
        return "Без прибора"                    # в источнике прибор не указан
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

# Операционные сутки: граница 07:00. 1 смена 07:00–19:59, 2 смена 20:00–06:59.
DAY_START = 7       # начало 1 смены и операционных суток
SHIFT2_START = 20   # начало 2 смены
OP_ORDER = list(range(7, 24)) + list(range(0, 7))   # порядок часов внутри опер. суток


def _shift(hour) -> str | None:
    if hour is None or (isinstance(hour, float) and hour != hour):
        return None
    h = int(hour)
    return "1 смена" if DAY_START <= h < SHIFT2_START else "2 смена"


def _idle_kind(note) -> str:
    """Плановый простой (обед, пересменка, ЕТО) vs внеплановый (ремонт/поломка/прочее)."""
    s = str(note).strip().lower()
    if "обед" in s or "пересмен" in s or s.startswith("ето"):
        return "плановый"
    return "внеплановый"


def _wavg_otk(frame) -> float | None:
    """Средняя откатка (м), взвешенная по количеству машин."""
    w = pd.to_numeric(frame["Количство машин, шт"], errors="coerce")
    v = pd.to_numeric(frame["_otk"], errors="coerce")
    mask = v.notna() & w.notna() & (w > 0)
    sw = float(w[mask].sum())
    return round(float((v[mask] * w[mask]).sum() / sw), 0) if sw > 0 else None


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

    df["Дата. Факт"] = pd.to_datetime(df["Дата. Факт"], errors="coerce")
    for c in ("Количство машин, шт", "Обьем работ, м3", "Обьем кузова,м3"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Подразделение"] = df["Подразделение"].astype(str).str.strip()
    if "Откатка, м" not in df.columns:
        df["Откатка, м"] = np.nan
    df["_otk"] = pd.to_numeric(df["Откатка, м"], errors="coerce")
    df["Час"] = df["Время"].apply(_to_hour)
    df["_hour"] = pd.to_numeric(df["Час"], errors="coerce")
    if "Примечание" not in df.columns:
        df["Примечание"] = ""
    for _col in ("Марка промывочного прибора", "Инв. № промывочного прибора"):
        if _col not in df.columns:
            df[_col] = ""

    # операционная дата строки: до 07:00 относится к предыдущим суткам (ночная смена)
    back = (df["_hour"] < DAY_START).fillna(False).astype(int)
    df["ОперДата"] = (df["Дата. Факт"].dt.normalize()
                      - pd.to_timedelta(back, unit="D")).dt.date

    # текущие операционные сутки (по часам, граница 07:00) + истёкшие часы
    now = datetime.now()
    op_today = (now.date() - timedelta(days=1)) if now.hour < DAY_START else now.date()
    rd_ts = pd.to_datetime(report_date, errors="coerce")
    if pd.isna(rd_ts):
        rd_date = op_today
    else:
        rd_date = rd_ts.date()
        if rd_date >= op_today:          # «сегодня/будущее» → текущие операционные сутки
            rd_date = op_today
    partial = (rd_date == op_today)      # отчётные сутки ещё идут (неполные)
    report_date = rd_date.isoformat()

    if partial:
        if now.hour >= DAY_START:
            elapsed = set(range(DAY_START, now.hour + 1))
        else:
            elapsed = set(range(DAY_START, 24)) | set(range(0, now.hour + 1))
    else:
        elapsed = set(range(24))
    op_hours = [h for h in OP_ORDER if h in elapsed]   # истёкшие часы в опер. порядке

    # прошедший (последний завершённый) час и его операционные сутки
    last_dt = now - timedelta(hours=1)
    last_h = last_dt.hour
    last_op = (last_dt.date() - timedelta(days=1)) if last_h < DAY_START else last_dt.date()

    # ── Период (только торф, водители — для ABC/парка/сводов) ──
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
                    Дней=("ОперДата", "nunique")).reset_index())
    by_unit["м3/рейс"] = by_unit["Объем_м3"] / by_unit["Рейсы"].replace(0, np.nan)
    by_unit["Доля_%"] = by_unit["Объем_м3"] / by_unit["Объем_м3"].sum() * 100
    by_unit = by_unit.sort_values("Объем_м3", ascending=False).reset_index(drop=True)

    by_date = (trans.groupby("ОперДата")
               .agg(Рейсы=("Количство машин, шт", "sum"),
                    Объем_м3=("Обьем работ, м3", "sum"),
                    Записей=("Водитель", "count")).reset_index()
               .rename(columns={"ОперДата": "Дата"}))

    drv = (trans.groupby("Водитель")
           .agg(Рейсы=("Количство машин, шт", "sum"),
                Объем_м3=("Обьем работ, м3", "sum"),
                Дней=("ОперДата", "nunique"),
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
                      Дней=("ОперДата", "nunique")).reset_index())
    truck_inv["м3/рейс"] = truck_inv["Объем_м3"] / truck_inv["Рейсы"].replace(0, np.nan)
    truck_inv["Исп_%"] = truck_inv["м3/рейс"] / truck_inv["Кузов_м3"] * 100
    truck_inv = truck_inv.sort_values(["Подразделение", "Объем_м3"],
                                      ascending=[True, False]).reset_index(drop=True)

    by_hour = (trans.groupby("Час").agg(Рейсы=("Количство машин, шт", "sum"),
                                        Объем_м3=("Обьем работ, м3", "sum"))
               .reindex(range(24), fill_value=0).reset_index())

    # ── Транспортировка торф+песок: материал/смена + машины по дням/часам/сменам ──
    transport = df[df["Передел"].astype(str).str.strip().isin(TRANSPORT_PEREDELY)].copy()
    transport["Материал"] = transport["Передел"].astype(str).str.strip().map(TRANSPORT_PEREDELY)
    transport["Смена"] = transport["Час"].apply(_shift)

    mach_by_day = (transport.groupby(["ОперДата", "Материал"])
                   .agg(Машины=("Количство машин, шт", "sum")).reset_index()
                   .rename(columns={"ОперДата": "Дата"}))
    mach_by_hour = (transport.dropna(subset=["Час"])
                    .groupby(["Час", "Материал"])
                    .agg(Машины=("Количство машин, шт", "sum")).reset_index())
    mach_by_shift = (transport.dropna(subset=["Смена"])
                     .groupby(["ОперДата", "Смена", "Материал"])
                     .agg(Машины=("Количство машин, шт", "sum")).reset_index()
                     .rename(columns={"ОперДата": "Дата"}))

    # пивоты «час × опер-дата» по материалам (операционный порядок часов)
    mach_hour_pivot = {}
    for mat in ("Торф", "Песок"):
        sub = transport[(transport["Материал"] == mat) & transport["_hour"].notna()]
        if len(sub):
            piv = (sub.pivot_table(index="_hour", columns="ОперДата",
                                   values="Количство машин, шт", aggfunc="sum")
                   .reindex(OP_ORDER))
            piv.index = [f"{h:02d}:00" for h in piv.index]
            piv.index.name = "Час"
            piv = piv.reset_index()
            piv.columns = ["Час"] + [c.isoformat() if hasattr(c, "isoformat") else str(c)
                                     for c in piv.columns[1:]]
            mach_hour_pivot[mat] = piv

    # те же пивоты «час × опер-дата», но в разрезе подразделений (для письма)
    mach_hour_pivot_unit = {}
    for unit in sorted(transport["Подразделение"].unique()):
        per_mat = {}
        for mat in ("Торф", "Песок"):
            sub = transport[(transport["Подразделение"] == unit)
                            & (transport["Материал"] == mat) & transport["_hour"].notna()]
            if len(sub):
                piv = (sub.pivot_table(index="_hour", columns="ОперДата",
                                       values="Количство машин, шт", aggfunc="sum")
                       .reindex(OP_ORDER))
                piv.index = [f"{h:02d}:00" for h in piv.index]
                piv.index.name = "Час"
                piv = piv.reset_index()
                piv.columns = ["Час"] + [c.isoformat() if hasattr(c, "isoformat") else str(c)
                                         for c in piv.columns[1:]]
                per_mat[mat] = piv
        if per_mat:
            mach_hour_pivot_unit[unit] = per_mat

    # ── Откатка (м), взвешенная по машинам: за сутки по подразделениям + динамика ──
    day_otk = transport[(transport["ОперДата"] == rd_date)
                        & (transport["_hour"].isin(elapsed) | transport["_hour"].isna())]
    otk_rows = []
    for unit in sorted(day_otk["Подразделение"].unique()):
        g = day_otk[day_otk["Подразделение"] == unit]
        otk_rows.append({"Подразделение": unit,
                         "Откатка_торф": _wavg_otk(g[g["Материал"] == "Торф"]),
                         "Откатка_песок": _wavg_otk(g[g["Материал"] == "Песок"])})
    otkatka_unit = pd.DataFrame(otk_rows, columns=["Подразделение", "Откатка_торф", "Откатка_песок"])
    otk_dates = sorted([d for d in transport["ОперДата"].dropna().unique()])[-7:]
    otk_units = sorted(transport["Подразделение"].unique())
    otk_dyn_rows = []
    for d in otk_dates:
        row = {"Дата": d.isoformat()}
        for unit in otk_units:
            row[unit] = _wavg_otk(transport[(transport["ОперДата"] == d)
                                            & (transport["Подразделение"] == unit)])
        otk_dyn_rows.append(row)
    otkatka_dyn = {"dates": [d.isoformat() for d in otk_dates],
                   "units": otk_units, "rows": otk_dyn_rows}

    # ── Операционный срез за отчётные сутки по подразделениям ──
    day_tr = transport[(transport["ОперДата"] == rd_date)
                       & (transport["_hour"].isin(elapsed) | transport["_hour"].isna())]

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
        "Объём торф, м³": vol_p["Торф"].round().astype(int).values,
        "Объём пески, м³": vol_p["Песок"].round().astype(int).values,
        "Кол-во машин торф, шт": mach_p["Торф"].round().astype(int).values,
        "Кол-во машин пески, шт": mach_p["Песок"].round().astype(int).values,
    })

    # простои по опер-дням в окне истёкших часов (для честной базы на частичных сутках)
    idle_src = df[df["Передел"].astype(str).str.strip() == "простой"].copy()
    idle_src["Тип"] = idle_src["Примечание"].apply(_idle_kind)
    idle_win = idle_src[idle_src["_hour"].isin(elapsed)]
    idle_df = (idle_win.groupby(["Подразделение", "ОперДата"]).size()
               .reset_index(name="Часов"))

    # динамика простоев по опер-дням (плановые/внеплановые, по предприятию, 7 дней)
    idle_dyn_dates = sorted([d for d in idle_src["ОперДата"].dropna().unique()])[-7:]
    idle_kind_day = idle_src.groupby(["ОперДата", "Тип"]).size()
    idle_dyn = []
    for d in idle_dyn_dates:
        pl = int(idle_kind_day.get((d, "плановый"), 0))
        un = int(idle_kind_day.get((d, "внеплановый"), 0))
        idle_dyn.append({"date": d.isoformat(), "planned": pl,
                         "unplanned": un, "total": pl + un})
    idle_today = idle_df[idle_df["ОперДата"] == rd_date].set_index("Подразделение")["Часов"]
    by_unit_day["Простои, ч"] = by_unit_day["Подразделение"].map(idle_today).fillna(0).astype(int)
    by_unit_day = by_unit_day.sort_values("Объём торф, м³", ascending=False).reset_index(drop=True)

    # для алертов: простои за сутки + база-медиана (то же окно, до 7 опер-дней назад)
    idle_rows = []
    for unit, g in idle_df.groupby("Подразделение"):
        today_h = int(g.loc[g["ОперДата"] == rd_date, "Часов"].sum())
        prior = g.loc[g["ОперДата"] < rd_date].sort_values("ОперДата")["Часов"].tail(7)
        baseline = float(prior.median()) if len(prior) else 0.0
        if today_h > 0 or baseline > 0:
            idle_rows.append({"Подразделение": unit, "Часов простоя": today_h,
                              "База_медиана": round(baseline, 1)})
    idle = pd.DataFrame(idle_rows, columns=["Подразделение", "Часов простоя", "База_медиана"])

    # отклонения: объём (торф+песок) и простои vs медиана за ТО ЖЕ окно суток (7 дней)
    tr_win = transport[transport["_hour"].isin(elapsed)]
    vol_df = (tr_win.groupby(["Подразделение", "ОперДата"])["Обьем работ, м3"].sum()
              .reset_index(name="Объём"))
    dev_rows = []
    for unit in sorted(transport["Подразделение"].unique()):
        gv = vol_df[vol_df["Подразделение"] == unit]
        today_v = float(gv.loc[gv["ОперДата"] == rd_date, "Объём"].sum())
        prior_v = gv.loc[gv["ОперДата"] < rd_date].sort_values("ОперДата")["Объём"].tail(7)
        base_v = float(prior_v.median()) if len(prior_v) else 0.0
        gi = idle_df[idle_df["Подразделение"] == unit]
        today_i = int(gi.loc[gi["ОперДата"] == rd_date, "Часов"].sum())
        prior_i = gi.loc[gi["ОперДата"] < rd_date].sort_values("ОперДата")["Часов"].tail(7)
        base_i = float(prior_i.median()) if len(prior_i) else 0.0
        dev_rows.append({"Подразделение": unit, "Объём": round(today_v), "Объём_база": round(base_v),
                         "Простои": today_i, "Простои_база": round(base_i)})
    unit_dev = pd.DataFrame(dev_rows)

    # ── Почасовка за отчётные сутки (истёкшие часы, опер. порядок, со сменой) ──
    day_all = df[df["ОперДата"] == rd_date].copy()
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
        for h in op_hours:
            hourly_rows.append({
                "Подразделение": unit, "Смена": _shift(h), "Час": f"{h:02d}:00",
                "Машин_торф": int(round(float(mt_h.get((unit, h), 0)))),
                "Машин_песок": int(round(float(mp_h.get((unit, h), 0)))),
                "Причина": reason_h.get((unit, h), ""),
            })
    hourly_unit = pd.DataFrame(
        hourly_rows,
        columns=["Подразделение", "Смена", "Час", "Машин_торф", "Машин_песок", "Причина"])

    # песок в разрезе промывочных приборов: пивот час × прибор (истёкшие часы)
    pes = day_tr_all[day_tr_all["Материал"] == "Песок"].dropna(subset=["Час"]).copy()
    pes["Прибор"] = [_device(mk, iv) for mk, iv in
                     zip(pes["Марка промывочного прибора"], pes["Инв. № промывочного прибора"])]
    pesok_devices = {}
    for unit in sorted(pes["Подразделение"].unique()):
        sub = pes[pes["Подразделение"] == unit]
        piv = (sub.pivot_table(index="Час", columns="Прибор",
                               values="Количство машин, шт", aggfunc="sum", fill_value=0)
               .reindex([h for h in op_hours], fill_value=0))
        piv.index = [f"{h:02d}:00" for h in piv.index]
        piv = piv.reset_index().rename(columns={"index": "Час"})
        piv.columns.name = None
        pesok_devices[unit] = piv

    # аналитика причин простоя за сутки (плановые/внеплановые, окно истёкших часов)
    di = day_all[(day_all["_per"] == "простой") & day_all["_hour"].isin(elapsed)].copy()
    di["Причина"] = di["Примечание"].astype(str).str.strip().replace({"nan": "—", "": "—"})
    di["Тип"] = di["Примечание"].apply(_idle_kind)
    idle_reasons = (di.groupby(["Подразделение", "Причина", "Тип"]).size()
                    .reset_index(name="Часов").sort_values("Часов", ascending=False)
                    .reset_index(drop=True))

    # KPI прошедшего часа (по предприятию, машины торф/пески)
    lh = transport[(transport["ОперДата"] == last_op) & (transport["_hour"] == last_h)]
    last_hour_kpi = {
        "hour": f"{last_h:02d}:00",
        "hour_to": f"{(last_h + 1) % 24:02d}:00",
        "torf": int(round(float(lh.loc[lh["Материал"] == "Торф", "Количство машин, шт"].sum()))),
        "pesok": int(round(float(lh.loc[lh["Материал"] == "Песок", "Количство машин, шт"].sum()))),
    }

    return {
        "report_date": report_date,
        "partial": bool(partial),
        "op_hours": [f"{h:02d}:00" for h in op_hours],
        "last_hour_kpi": last_hour_kpi,
        "date_min": df["Дата. Факт"].min(),
        "date_max": df["Дата. Факт"].max(),
        "totals": totals,
        "by_unit": by_unit, "by_date": by_date, "by_hour": by_hour,
        "drivers": drv, "truck_mark": truck_mark, "truck_inv": truck_inv,
        "idle": idle, "units": sorted(trans["Подразделение"].unique().tolist()),
        "n_rows": int(len(df)), "n_trans": int(len(trans)),
        "mach_by_day": mach_by_day, "mach_by_hour": mach_by_hour,
        "mach_by_shift": mach_by_shift, "mach_hour_pivot": mach_hour_pivot,
        "mach_hour_pivot_unit": mach_hour_pivot_unit,
        "by_unit_day": by_unit_day, "unit_dev": unit_dev,
        "hourly_unit": hourly_unit, "idle_reasons": idle_reasons,
        "idle_dyn": idle_dyn, "pesok_devices": pesok_devices,
        "otkatka_unit": otkatka_unit, "otkatka_dyn": otkatka_dyn,
    }


def _dynamics(mach_hour_pivot: dict, days: int = 7) -> dict:
    """Динамика машин «час × дата» по материалам — последние `days` дат для письма."""
    out = {}
    for mat, piv in mach_hour_pivot.items():
        date_cols = [c for c in piv.columns if c != "Час"][-days:]
        rows = []
        for _, r in piv.iterrows():
            row = {"hour": r["Час"]}
            for c in date_cols:
                v = r[c]
                row[c] = int(round(float(v))) if pd.notna(v) else 0
            rows.append(row)
        out[mat] = {"dates": date_cols, "rows": rows}
    return out


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
        "partial": aggr.get("partial", False),
        "op_hours": aggr.get("op_hours", []),
        "last_hour_kpi": aggr.get("last_hour_kpi", {}),
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
        "mach_dynamics": _dynamics(aggr.get("mach_hour_pivot", {})),
        "mach_dynamics_unit": {unit: _dynamics(pm)
                               for unit, pm in aggr.get("mach_hour_pivot_unit", {}).items()},
        "otkatka_unit": [{"unit": r["Подразделение"],
                          "torf": None if pd.isna(r["Откатка_торф"]) else int(r["Откатка_торф"]),
                          "pesok": None if pd.isna(r["Откатка_песок"]) else int(r["Откатка_песок"])}
                         for _, r in aggr["otkatka_unit"].iterrows()],
        "otkatka_dyn": {
            "dates": aggr["otkatka_dyn"]["dates"],
            "units": aggr["otkatka_dyn"]["units"],
            "rows": [{k: (None if v is None or (isinstance(v, float) and pd.isna(v))
                          else (v if k == "Дата" else int(round(float(v)))))
                      for k, v in row.items()}
                     for row in aggr["otkatka_dyn"]["rows"]],
        },
        "idle_dyn": aggr.get("idle_dyn", []),
        "by_unit_day": [{"unit": r["Подразделение"],
                         "vol_torf": int(r["Объём торф, м³"]), "vol_pesok": int(r["Объём пески, м³"]),
                         "mach_torf": int(r["Кол-во машин торф, шт"]),
                         "mach_pesok": int(r["Кол-во машин пески, шт"]),
                         "idle_h": int(r["Простои, ч"])}
                        for _, r in aggr["by_unit_day"].iterrows()],
        "unit_dev": [{"unit": r["Подразделение"], "vol": int(r["Объём"]), "vol_base": int(r["Объём_база"]),
                      "idle": int(r["Простои"]), "idle_base": int(r["Простои_база"])}
                     for _, r in aggr["unit_dev"].iterrows()],
        "hourly_unit": [{"unit": r["Подразделение"], "shift": r["Смена"], "hour": r["Час"],
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
