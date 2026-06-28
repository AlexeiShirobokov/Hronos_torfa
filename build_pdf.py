"""PDF-аналитика по хронометражу транспортировки торфов.

Читает logs/consolidated.csv (создаётся consolidate.py), сохраняет
output/Аналитика_хронометраж_торфов_<дата>.pdf.

Содержит: KPI, подразделения, динамику по датам/часам, грузоподъёмность,
ABC-анализ (общий и по подразделениям), водителей по подразделениям,
выводы/рекомендации и чек-лист начальника участка.
"""
# DEPRECATED (Фаза 1): не используется в пайплайне — отчёт теперь HTML-письмо +
# Excel с листами аналитики (analytics.py / build_xlsx.py / email_html.py).
# Оставлено для истории; удалить после Фазы 2.
from __future__ import annotations
import json, math, os, sys, textwrap
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.font_manager as fm
from matplotlib.patches import Patch

for f in fm.findSystemFonts(fontpaths=None, fontext="ttf"):
    if "DejaVuSans" in f:
        fm.fontManager.addfont(f)
plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["font.size"] = 9
plt.rcParams["axes.unicode_minus"] = False

BASE = Path(__file__).resolve().parent
OUT = BASE / "output"
LOGS = BASE / "logs"
CSV = LOGS / "consolidated.csv"

UNIT_COLORS = {"Эрел": "#2E75B6", "Дражный": "#70AD47",
               "Обман": "#C00000", "Сайлык": "#F4B400"}
ABC_COLORS = {"A": "#2E7D32", "B": "#F9A825", "C": "#C62828"}


def fmt_int(x):
    if pd.isna(x): return "—"
    try: return f"{int(round(float(x))):,}".replace(",", " ")
    except (ValueError, TypeError): return str(x)


def fmt_f(x, d=1):
    if pd.isna(x): return "—"
    return f"{x:,.{d}f}".replace(",", " ")


def to_hour(x):
    if pd.isna(x) or x in (None, "", " "): return None
    s = str(x).strip()
    if " " in s and ":" in s.split(" ", 1)[1]: s = s.split(" ", 1)[1]
    if ":" in s:
        try: return int(s.split(":", 1)[0])
        except ValueError: return None
    return None


def color_for_unit(u): return UNIT_COLORS.get(u, "#444444")


def draw_table(ax, df_, title=None, col_widths=None, fontsize=8,
               highlight_col=None, highlight_map=None):
    ax.axis("off")
    if title: ax.set_title(title, fontsize=11, weight="bold", loc="left", pad=8)
    rows = df_.values.tolist()
    cols = list(df_.columns)
    t = ax.table(cellText=rows, colLabels=cols, loc="upper left",
                 cellLoc="center", colWidths=col_widths)
    t.auto_set_font_size(False); t.set_fontsize(fontsize); t.scale(1, 1.25)
    for c in range(len(cols)):
        cell = t[0, c]
        cell.set_facecolor("#305496"); cell.set_text_props(color="white", weight="bold")
    for r in range(1, len(rows)+1):
        for c in range(len(cols)):
            t[r, c].set_facecolor("#F2F2F2" if r % 2 == 0 else "#FFFFFF")
    if highlight_col is not None and highlight_map:
        ci = cols.index(highlight_col)
        for r in range(1, len(rows)+1):
            v = df_.iloc[r-1][highlight_col]
            c = highlight_map.get(v)
            if c:
                t[r, ci].set_facecolor(c)
                t[r, ci].set_text_props(weight="bold", color="white")


def abc_class(p):
    if p <= 80.0: return "A"
    if p <= 95.0: return "B"
    return "C"


def main() -> int:
    if not CSV.exists():
        print(f"[ERR] нет {CSV} (запусти consolidate.py)"); return 2
    df = pd.read_csv(CSV)

    report_date = datetime.now().strftime("%Y-%m-%d")
    meta = LOGS / "last_fetch.meta"
    if meta.exists():
        for line in meta.read_text(encoding="utf-8").splitlines():
            if line.startswith("date="):
                report_date = line.split("=", 1)[1].strip()

    df["Дата. Факт"] = pd.to_datetime(df["Дата. Факт"], errors="coerce")
    for c in ("Количство машин, шт", "Обьем работ, м3", "Обьем кузова,м3"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Подразделение"] = df["Подразделение"].astype(str).str.strip()
    df["Час"] = df["Время"].apply(to_hour)

    trans = df[df["Передел"].astype(str).str.strip() == "Транспортировка торфов"].copy()
    trans["Водитель"] = trans["Ф.И.О. водителя самосвала"].astype(str).str.strip().replace({"nan": np.nan, "": np.nan})
    trans = trans[trans["Водитель"].notna()]

    total_volume = trans["Обьем работ, м3"].sum()
    total_trips = trans["Количство машин, шт"].sum()
    n_drivers = trans["Водитель"].nunique()
    n_trucks = trans["Инв. № транспортировочной единицы"].dropna().nunique()
    n_units = trans["Подразделение"].nunique()
    date_min = df["Дата. Факт"].min(); date_max = df["Дата. Факт"].max()
    UNITS = sorted(trans["Подразделение"].unique())

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
                    Записей=("Водитель", "count")).reset_index().rename(columns={"Дата. Факт": "Дата"}))
    by_date_unit = (trans.groupby([trans["Дата. Факт"].dt.date, "Подразделение"])
                    .agg(Объем_м3=("Обьем работ, м3", "sum"),
                         Рейсы=("Количство машин, шт", "sum"))
                    .reset_index().rename(columns={"Дата. Факт": "Дата"}))

    drv = (trans.groupby("Водитель")
           .agg(Рейсы=("Количство машин, шт", "sum"),
                Объем_м3=("Обьем работ, м3", "sum"),
                Дней=("Дата. Факт", "nunique"),
                Подразделение=("Подразделение",
                               lambda s: s.mode().iat[0] if len(s.mode()) else "—"))
           .reset_index())
    drv["м3/рейс"] = drv["Объем_м3"] / drv["Рейсы"].replace(0, np.nan)
    drv = drv.sort_values("Объем_м3", ascending=False).reset_index(drop=True)
    total_v = drv["Объем_м3"].sum()
    drv["Доля_%"] = drv["Объем_м3"] / total_v * 100
    drv["Накопл_%"] = drv["Доля_%"].cumsum()
    drv["ABC"] = drv["Накопл_%"].apply(abc_class)

    drv_by_unit = {}
    for u in UNITS:
        sub = (trans[trans["Подразделение"] == u].groupby("Водитель")
               .agg(Рейсы=("Количство машин, шт", "sum"),
                    Объем_м3=("Обьем работ, м3", "sum"),
                    Дней=("Дата. Факт", "nunique")).reset_index())
        sub["м3/рейс"] = sub["Объем_м3"] / sub["Рейсы"].replace(0, np.nan)
        sub = sub.sort_values("Объем_м3", ascending=False).reset_index(drop=True)
        tv = sub["Объем_м3"].sum()
        sub["Доля_%"] = sub["Объем_м3"] / tv * 100 if tv else 0
        sub["Накопл_%"] = sub["Доля_%"].cumsum()
        sub["ABC"] = sub["Накопл_%"].apply(abc_class)
        drv_by_unit[u] = sub

    truck_mark = (trans.groupby("Марка транспортировочной единицы")
                  .agg(Кузов_м3_сред=("Обьем кузова,м3", "mean"),
                       Рейсы=("Количство машин, шт", "sum"),
                       Объем_м3=("Обьем работ, м3", "sum"),
                       Самосвалов=("Инв. № транспортировочной единицы",
                                   lambda s: s.dropna().nunique())).reset_index())
    truck_mark["м3/рейс"] = truck_mark["Объем_м3"] / truck_mark["Рейсы"].replace(0, np.nan)
    truck_mark["Использование_кузова_%"] = truck_mark["м3/рейс"] / truck_mark["Кузов_м3_сред"] * 100
    truck_mark = truck_mark.sort_values("Объем_м3", ascending=False).reset_index(drop=True)

    truck_inv = (trans.dropna(subset=["Инв. № транспортировочной единицы"])
                 .groupby(["Марка транспортировочной единицы",
                           "Инв. № транспортировочной единицы", "Подразделение"])
                 .agg(Кузов_м3=("Обьем кузова,м3", "mean"),
                      Рейсы=("Количство машин, шт", "sum"),
                      Объем_м3=("Обьем работ, м3", "sum"),
                      Дней=("Дата. Факт", "nunique")).reset_index())
    truck_inv["м3/рейс"] = truck_inv["Объем_м3"] / truck_inv["Рейсы"].replace(0, np.nan)
    truck_inv["Исп_кузова_%"] = truck_inv["м3/рейс"] / truck_inv["Кузов_м3"] * 100
    truck_inv = truck_inv.sort_values(["Подразделение", "Объем_м3"],
                                      ascending=[True, False]).reset_index(drop=True)

    rd = pd.to_datetime(report_date)
    trans_day = trans[trans["Дата. Факт"] == rd]
    hour_dyn = (trans_day.groupby("Час").agg(Рейсы=("Количство машин, шт", "sum"),
                                              Объем_м3=("Обьем работ, м3", "sum"))
                .reindex(range(24), fill_value=0).reset_index())

    idle = (df[df["Передел"].astype(str).str.strip() == "простой"]
            .groupby(["Подразделение", "Дата. Факт"]).size().reset_index(name="Часов простоя"))

    OUT.mkdir(parents=True, exist_ok=True)
    pdf_path = OUT / f"Аналитика_хронометраж_торфов_{report_date}.pdf"

    with PdfPages(pdf_path) as pdf:
        # ===== Стр. 1 Титул =====
        fig = plt.figure(figsize=(8.27, 11.69))
        ax = fig.add_axes([0.07, 0.7, 0.86, 0.27]); ax.axis("off")
        ax.text(0, 1.0, "Аналитика по хронометражу", fontsize=22, weight="bold")
        ax.text(0, 0.85, "транспортировки торфов", fontsize=22, weight="bold")
        ax.text(0, 0.65, f"Отчётная дата: {report_date}", fontsize=12)
        ax.text(0, 0.55, f"Период данных: {date_min.date() if pd.notna(date_min) else '—'} … {date_max.date() if pd.notna(date_max) else '—'}", fontsize=10)
        ax.text(0, 0.45, f"Сформировано: {datetime.now():%Y-%m-%d %H:%M}", fontsize=9, color="#555")
        ax.text(0, 0.30, f"Источник: консолидированный реестр (4 подразделения, {len(df)} строк, передел «Транспортировка торфов»: {len(trans)} записей)", fontsize=9, color="#555")
        ax_k = fig.add_axes([0.07, 0.36, 0.86, 0.30]); ax_k.axis("off")
        tiles = [("Общий объём", f"{fmt_int(total_volume)} м³"),
                 ("Рейсов", f"{fmt_int(total_trips)}"),
                 ("Водителей", f"{n_drivers}"),
                 ("Самосвалов", f"{n_trucks}"),
                 ("Подразделений", f"{n_units}"),
                 ("м³ / рейс", f"{fmt_f(total_volume/total_trips,1) if total_trips else '—'}")]
        for i,(k,v) in enumerate(tiles):
            col, row = i % 3, i // 3
            x0 = col * 0.34; y0 = 0.55 - row * 0.30
            ax_k.add_patch(plt.Rectangle((x0, y0), 0.30, 0.22, facecolor="#EEF3FA", edgecolor="#305496"))
            ax_k.text(x0+0.015, y0+0.16, k, fontsize=8, color="#305496")
            ax_k.text(x0+0.015, y0+0.05, v, fontsize=13, weight="bold")
        ax_k.set_xlim(0, 1); ax_k.set_ylim(0, 1)
        ax_t = fig.add_axes([0.07, 0.04, 0.86, 0.28])
        td = pd.DataFrame({"Подразделение": by_unit["Подразделение"],
                           "Рейсы": by_unit["Рейсы"].apply(fmt_int),
                           "Объём, м³": by_unit["Объем_м3"].apply(fmt_int),
                           "Доля, %": by_unit["Доля_%"].apply(lambda v: fmt_f(v,1)),
                           "Водителей": by_unit["Водителей"].apply(fmt_int),
                           "Самосвалов": by_unit["Самосвалов"].apply(fmt_int),
                           "Дней": by_unit["Дней"].apply(fmt_int),
                           "м³/рейс": by_unit["м3/рейс"].apply(lambda v: fmt_f(v,1))})
        draw_table(ax_t, td, title="Распределение по подразделениям",
                   col_widths=[0.18,0.10,0.14,0.10,0.11,0.11,0.08,0.10])
        pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        # ===== Стр. 2 Сравнение =====
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.suptitle("Сравнение подразделений", fontsize=14, weight="bold", x=0.07, ha="left")
        for i, (col, title, fmt) in enumerate([
            ("Объем_м3","Объём транспортировки, м³", fmt_int),
            ("Рейсы","Рейсы, шт", fmt_int),
            ("м3/рейс","Средний м³ за рейс", lambda v: fmt_f(v,1)),
            ("Водителей","Уникальных водителей", fmt_int),
        ]):
            ax = fig.add_subplot(2,2,i+1)
            cols = [color_for_unit(u) for u in by_unit["Подразделение"]]
            ax.bar(by_unit["Подразделение"], by_unit[col], color=cols)
            ax.set_title(title)
            for x,v in zip(by_unit["Подразделение"], by_unit[col]):
                ax.text(x, v, fmt(v), ha="center", va="bottom", fontsize=8)
            ax.grid(axis="y", linestyle=":", alpha=0.5)
        fig.tight_layout(rect=[0,0,1,0.96])
        pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        # ===== Стр. 3 Динамика =====
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.suptitle("Динамика по датам", fontsize=14, weight="bold", x=0.07, ha="left")
        ax = fig.add_subplot(2,1,1)
        ax.bar([str(d) for d in by_date["Дата"]], by_date["Объем_м3"], color="#305496")
        for x,v in zip(by_date["Дата"], by_date["Объем_м3"]):
            ax.text(str(x), v, fmt_int(v), ha="center", va="bottom", fontsize=8)
        ax.set_ylabel("Объём, м³"); ax.set_title("Объём по датам"); ax.grid(axis="y", linestyle=":", alpha=0.5)
        ax = fig.add_subplot(2,1,2)
        pivot = by_date_unit.pivot(index="Дата", columns="Подразделение", values="Объем_м3").fillna(0).sort_index()
        bottom = np.zeros(len(pivot)); xs = [str(d) for d in pivot.index]
        for u in pivot.columns:
            ax.bar(xs, pivot[u].values, bottom=bottom, label=u, color=color_for_unit(u))
            for i, v in enumerate(pivot[u].values):
                if v > 200: ax.text(xs[i], bottom[i]+v/2, fmt_int(v), ha="center", va="center", fontsize=7, color="white")
            bottom += pivot[u].values
        ax.legend(loc="upper left", fontsize=8, frameon=False)
        ax.set_ylabel("Объём, м³"); ax.set_title("Объём по подразделениям × дата"); ax.grid(axis="y", linestyle=":", alpha=0.5)
        fig.tight_layout(rect=[0,0,1,0.96])
        pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        # ===== Стр. 4 По датам таблица + час-пик =====
        fig = plt.figure(figsize=(8.27, 11.69))
        ax = fig.add_axes([0.06, 0.55, 0.88, 0.40])
        tbl = pd.DataFrame({"Дата":[str(d) for d in by_date["Дата"]],
                            "Рейсы": by_date["Рейсы"].apply(fmt_int),
                            "Объём, м³": by_date["Объем_м3"].apply(fmt_int),
                            "м³/рейс":(by_date["Объем_м3"]/by_date["Рейсы"].replace(0,np.nan)).apply(lambda v: fmt_f(v,1)),
                            "Записей": by_date["Записей"].apply(fmt_int)})
        draw_table(ax, tbl, title="Сводка по датам", col_widths=[0.20,0.16,0.20,0.16,0.16])
        ax = fig.add_axes([0.06, 0.06, 0.88, 0.42])
        ax.bar(hour_dyn["Час"], hour_dyn["Рейсы"], color="#2E75B6", label="Рейсы")
        ax2 = ax.twinx(); ax2.plot(hour_dyn["Час"], hour_dyn["Объем_м3"], color="#C00000", marker="o", linewidth=2)
        ax.set_xticks(range(24)); ax.set_xlabel("Час суток"); ax.set_ylabel("Рейсы"); ax2.set_ylabel("Объём, м³")
        ax.set_title(f"Часовая динамика · {report_date}", fontsize=11, weight="bold", loc="left")
        ax.grid(axis="y", linestyle=":", alpha=0.5)
        pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        # ===== Стр. 5 Грузоподъёмность =====
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.suptitle("Грузоподъёмность и использование самосвалов", fontsize=14, weight="bold", x=0.07, ha="left")
        fig.text(0.07, 0.93,
                 "Использование = средний м³/рейс ÷ номинал кузова. В реестре объём считается как «кузов × рейсы», что даёт ~100% — нужен фактический замер.",
                 fontsize=9, color="#555")
        ax = fig.add_subplot(2,1,1)
        x = np.arange(len(truck_mark)); w = 0.35
        ax.bar(x-w/2, truck_mark["Кузов_м3_сред"], w, label="Номинал кузова, м³", color="#A6A6A6")
        ax.bar(x+w/2, truck_mark["м3/рейс"], w, label="Фактический м³/рейс", color="#305496")
        ax.set_xticks(x); ax.set_xticklabels(truck_mark["Марка транспортировочной единицы"], rotation=10)
        ax.set_ylabel("м³"); ax.legend(frameon=False, fontsize=8)
        ax.set_title("Кузов vs средний рейс"); ax.grid(axis="y", linestyle=":", alpha=0.5)
        ax = fig.add_subplot(2,1,2)
        util_colors = ["#2E7D32" if v >= 95 else ("#F9A825" if v >= 80 else "#C62828")
                       for v in truck_mark["Использование_кузова_%"]]
        ax.barh(truck_mark["Марка транспортировочной единицы"], truck_mark["Использование_кузова_%"], color=util_colors)
        ax.set_xlabel("Использование, %"); ax.set_xlim(0, 110)
        ax.axvline(80, linestyle="--", color="#F9A825", linewidth=1)
        ax.axvline(95, linestyle="--", color="#2E7D32", linewidth=1)
        for y, v in enumerate(truck_mark["Использование_кузова_%"]):
            ax.text(v+1, y, fmt_f(v,1)+" %", va="center", fontsize=8)
        ax.set_title("Степень загрузки кузова (норма ≥ 95%)"); ax.grid(axis="x", linestyle=":", alpha=0.5)
        fig.tight_layout(rect=[0,0,1,0.92])
        pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        # ===== Стр. 6 Таблицы по технике =====
        fig = plt.figure(figsize=(8.27, 11.69))
        ax = fig.add_axes([0.04, 0.55, 0.92, 0.41])
        tm = pd.DataFrame({"Марка": truck_mark["Марка транспортировочной единицы"],
                           "Самосвалов": truck_mark["Самосвалов"].apply(fmt_int),
                           "Кузов, м³": truck_mark["Кузов_м3_сред"].apply(lambda v: fmt_f(v,1)),
                           "Рейсы": truck_mark["Рейсы"].apply(fmt_int),
                           "Объём, м³": truck_mark["Объем_м3"].apply(fmt_int),
                           "м³/рейс": truck_mark["м3/рейс"].apply(lambda v: fmt_f(v,1)),
                           "Исп., %": truck_mark["Использование_кузова_%"].apply(lambda v: fmt_f(v,1))})
        draw_table(ax, tm, title="Самосвалы по марке", col_widths=[0.20,0.12,0.12,0.12,0.16,0.12,0.16])
        ax = fig.add_axes([0.04, 0.04, 0.92, 0.46])
        ti = truck_inv.head(28)
        td = pd.DataFrame({"Подр.": ti["Подразделение"],
                           "Марка": ti["Марка транспортировочной единицы"],
                           "Инв №": ti["Инв. № транспортировочной единицы"].apply(fmt_int),
                           "Кузов": ti["Кузов_м3"].apply(lambda v: fmt_f(v,1)),
                           "Рейсы": ti["Рейсы"].apply(fmt_int),
                           "Объём": ti["Объем_м3"].apply(fmt_int),
                           "м³/рейс": ti["м3/рейс"].apply(lambda v: fmt_f(v,1)),
                           "Исп. %": ti["Исп_кузова_%"].apply(lambda v: fmt_f(v,1))})
        draw_table(ax, td, title="Самосвалы по инв.№ (топ-28)",
                   col_widths=[0.12,0.16,0.10,0.11,0.10,0.13,0.12,0.10], fontsize=8)
        pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        # ===== Стр. 7 ABC общий =====
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.suptitle("ABC-анализ водителей самосвалов", fontsize=14, weight="bold", x=0.07, ha="left")
        ax = fig.add_subplot(2,1,1)
        x = np.arange(len(drv))
        ax.bar(x, drv["Объем_м3"], color=[ABC_COLORS[c] for c in drv["ABC"]])
        ax.set_ylabel("Объём, м³"); ax.set_xticks([]); ax.set_title("Парето: объём по водителям")
        ax2 = ax.twinx(); ax2.plot(x, drv["Накопл_%"], color="#1F3864", marker=".", linewidth=1.6)
        ax2.set_ylim(0, 105); ax2.set_ylabel("Накопл, %")
        ax2.axhline(80, linestyle="--", color="#2E7D32", linewidth=1)
        ax2.axhline(95, linestyle="--", color="#F9A825", linewidth=1)
        legend = [Patch(facecolor=ABC_COLORS["A"], label="A — до 80%"),
                  Patch(facecolor=ABC_COLORS["B"], label="B — 80–95%"),
                  Patch(facecolor=ABC_COLORS["C"], label="C — последние 5%")]
        ax.legend(handles=legend, loc="upper right", fontsize=8, frameon=False)

        abc_sum = (drv.groupby("ABC")
                   .agg(Водителей=("Водитель","count"), Рейсы=("Рейсы","sum"), Объем=("Объем_м3","sum"))
                   .reindex(["A","B","C"], fill_value=0).reset_index())
        abc_sum["Доля, %"] = abc_sum["Объем"] / total_v * 100
        ax = fig.add_subplot(2,1,2)
        draw_table(ax, pd.DataFrame({"Класс": abc_sum["ABC"],
                                     "Водителей": abc_sum["Водителей"].apply(fmt_int),
                                     "Рейсы": abc_sum["Рейсы"].apply(fmt_int),
                                     "Объём, м³": abc_sum["Объем"].apply(fmt_int),
                                     "Доля, %": abc_sum["Доля, %"].apply(lambda v: fmt_f(v,1))}),
                   title="Сводка ABC", col_widths=[0.12,0.16,0.18,0.22,0.22], fontsize=10,
                   highlight_col="Класс", highlight_map=ABC_COLORS)
        fig.tight_layout(rect=[0,0,1,0.95]); pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        # ===== Стр. 8 ABC по подразделениям =====
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.suptitle("ABC-анализ по подразделениям", fontsize=14, weight="bold", x=0.07, ha="left")
        for i, u in enumerate(UNITS):
            sub = drv_by_unit[u]
            ax = fig.add_subplot(len(UNITS),1,i+1)
            if not len(sub): ax.axis("off"); ax.set_title(u); continue
            x = np.arange(len(sub))
            ax.bar(x, sub["Объем_м3"], color=[ABC_COLORS[c] for c in sub["ABC"]])
            ax.set_xticks(x)
            ax.set_xticklabels(sub["Водитель"].apply(lambda s: s.replace(" ","\n",1)), fontsize=6)
            ax.set_title(f"{u}: {len(sub)} водителей, объём {fmt_int(sub['Объем_м3'].sum())} м³",
                         fontsize=11, weight="bold", loc="left")
            ax.set_ylabel("м³"); ax.grid(axis="y", linestyle=":", alpha=0.5)
        fig.tight_layout(rect=[0,0,1,0.96]); pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        # ===== Стр. 9+ Водители по подразделениям =====
        for u in UNITS:
            sub = drv_by_unit[u]
            if not len(sub): continue
            fig = plt.figure(figsize=(8.27, 11.69))
            fig.suptitle(f"Водители · {u}", fontsize=14, weight="bold", x=0.07, ha="left")
            ax = fig.add_subplot(2,1,1)
            top = sub.head(15).iloc[::-1]
            ax.barh(top["Водитель"], top["Объем_м3"], color=[ABC_COLORS[c] for c in top["ABC"]])
            for y, (v, r) in enumerate(zip(top["Объем_м3"], top["Рейсы"])):
                ax.text(v+5, y, f"{fmt_int(v)} м³ · {fmt_int(r)} рейс", va="center", fontsize=7)
            ax.set_xlabel("Объём, м³"); ax.set_title(f"Топ-15 · {u}"); ax.grid(axis="x", linestyle=":", alpha=0.5)
            ax = fig.add_subplot(2,1,2)
            tbl = pd.DataFrame({"#":(sub.index+1).astype(int),
                                "Водитель": sub["Водитель"],
                                "Рейсы": sub["Рейсы"].apply(fmt_int),
                                "Объём, м³": sub["Объем_м3"].apply(fmt_int),
                                "м³/рейс": sub["м3/рейс"].apply(lambda v: fmt_f(v,1)),
                                "Доля, %": sub["Доля_%"].apply(lambda v: fmt_f(v,1)),
                                "Накопл, %": sub["Накопл_%"].apply(lambda v: fmt_f(v,1)),
                                "ABC": sub["ABC"]})
            draw_table(ax, tbl, title=f"Все водители · {u}",
                       col_widths=[0.05,0.26,0.10,0.13,0.10,0.10,0.12,0.08], fontsize=7,
                       highlight_col="ABC", highlight_map=ABC_COLORS)
            fig.tight_layout(rect=[0,0,1,0.96]); pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        # ===== Выводы и рекомендации =====
        obs = []
        leader = by_unit.iloc[0]; last = by_unit.iloc[-1]
        obs.append(f"Лидер по объёму — «{leader['Подразделение']}» ({fmt_int(leader['Объем_м3'])} м³, {fmt_f(leader['Доля_%'],1)}%). Минимум — «{last['Подразделение']}» ({fmt_int(last['Объем_м3'])} м³).")
        a_drivers = (drv["ABC"]=="A").sum(); a_share = abc_sum.loc[abc_sum["ABC"]=="A","Доля, %"].iat[0]
        obs.append(f"Класс A: {a_drivers} из {n_drivers} водителей ({fmt_f(a_drivers/n_drivers*100,0)}% штата) дают {fmt_f(a_share,1)}% объёма.")
        over = truck_mark[truck_mark["Использование_кузова_%"] >= 100]
        if len(over):
            obs.append("В реестре объём = кузов × рейсы, поэтому загрузка ~100% — это методология, а не факт. Нужен фактический замер на разгрузке.")
        parks = trans.groupby("Подразделение")["Марка транспортировочной единицы"].nunique()
        one = parks[parks == 1]
        if len(one):
            obs.append("Однотипный парк: " + ", ".join(f"«{u}» — одна марка" for u in one.index) + ". Риск выбытия техники.")
        if len(idle):
            ob = idle.groupby("Подразделение")["Часов простоя"].sum().sort_values(ascending=False)
            obs.append("Зафиксирован простой: " + ", ".join(f"«{u}» — {int(h)}" for u,h in ob.items()) + ".")
        if hour_dyn["Рейсы"].sum():
            peak = hour_dyn.loc[hour_dyn["Рейсы"].idxmax()]
            obs.append(f"Пик рейсов на {report_date}: {int(peak['Час']):02d}:00 ({fmt_int(peak['Рейсы'])} рейсов).")

        recs = [
            "Ввести фактический замер объёма на разгрузке: текущий учёт считает кузов × рейсы, недогрузка не фиксируется.",
            "Перераспределить водителей класса A на ключевые забои; за классом C — наставник из A.",
            "Сократить «низкие» часы внутри смены: контроль пересменок, ТО, обедов.",
            "Перевести часть рейсов с БелАЗ 7547 (12 м³) на 18–20 м³, где позволяет забой.",
            "Сверить кузов БелАЗ 7555В по инв. № (18 vs 20 м³) — документально.",
            "Для подразделений с одной маркой парка предусмотреть резервный самосвал.",
            f"Ежедневный мониторинг 3 KPI: рейсы/час, м³/рейс по марке, доля простоев. Пороги — 95% загрузки, простои <1ч/смену.",
        ]
        chief = [
            "Сверить план/факт по объёму торфа за смену; фиксировать причины отклонений в Примечании.",
            "На утренней планёрке распределить технику класса A на ключевые забои; назначить наставника для C.",
            "Контролировать старт рейсов с первого часа смены (09:00 и 21:00).",
            "Следить за полной разовой загрузкой кузова (≥ 95% номинала).",
            "Вести табель простоев пофамильно с причиной и временем восстановления.",
            "Сверять инв. номера с журналом смены.",
            "Минимум 1 раз/смену — обход забоя и разгрузки, отметить в Примечаниях.",
            "Закрывать реестр и пересылать в общий канал не позже 1 часа после смены.",
            "Пояснять строки с пустым «Обьем работ» — пометить переделом «простой».",
            "Эскалировать дефицит парка («Обман» — только БелАЗ 7547) главному инженеру.",
        ]
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.suptitle("Выводы, рекомендации, действия начальника участка",
                     fontsize=14, weight="bold", x=0.07, ha="left")
        ax = fig.add_axes([0.07, 0.62, 0.86, 0.34]); ax.axis("off")
        ax.text(0, 1.0, "Ключевые наблюдения", fontsize=11, weight="bold")
        y = 0.94
        for line in obs:
            wrapped = textwrap.fill(f"• {line}", width=95)
            ax.text(0, y, wrapped, fontsize=8, va="top"); y -= 0.04 * (wrapped.count("\n")+2)
        ax = fig.add_axes([0.07, 0.30, 0.86, 0.30]); ax.axis("off")
        ax.text(0, 1.0, "Рекомендации", fontsize=11, weight="bold"); y = 0.93
        for i, line in enumerate(recs, 1):
            wrapped = textwrap.fill(f"{i}. {line}", width=95)
            ax.text(0, y, wrapped, fontsize=8, va="top"); y -= 0.05 * (wrapped.count("\n")+2)
        pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        fig = plt.figure(figsize=(8.27, 11.69))
        fig.suptitle("Чек-лист начальника участка на смену",
                     fontsize=14, weight="bold", x=0.07, ha="left")
        ax = fig.add_axes([0.07, 0.04, 0.86, 0.90]); ax.axis("off")
        ax.text(0, 1.0, f"По данным {report_date}. Пройти по списку перед сменой и в её ходе.",
                fontsize=9, color="#555")
        y = 0.94
        for i, line in enumerate(chief, 1):
            wrapped = textwrap.fill(f"[ ] {i}. {line}", width=92)
            ax.text(0, y, wrapped, fontsize=9, va="top"); y -= 0.055 * (wrapped.count("\n")+2)
        pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

        d = pdf.infodict()
        d["Title"] = "Аналитика по хронометражу транспортировки торфов"
        d["Subject"] = f"Отчёт за {report_date}"
        d["CreationDate"] = datetime.now()

    info = {"ok": True, "pdf": str(pdf_path),
            "size": os.path.getsize(pdf_path),
            "report_date": report_date,
            "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    (LOGS / "last_pdf.json").write_text(json.dumps(info, ensure_ascii=False, indent=2),
                                        encoding="utf-8")
    print(json.dumps(info, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
