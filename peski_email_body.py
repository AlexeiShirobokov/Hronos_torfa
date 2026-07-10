"""Тело письма «Хронометраж транспортировки торфов и песков» из финального реестра.

Живой (параллельный) пайплайн раньше слал письмо с пустым телом (--empty-body).
Этот модуль строит компактное HTML-тело за отчётный день (Дата выдачи наряд-задания
== report_date, как в сводной), чтобы не открывать вложение:

  1. Пески по подразделениям (объём м³).
  2. По промывочному прибору «Марка / Инв.№»: объём м³ + часы работы.
  3. Топ причин простоя по часам + горизонтальная диаграмма (CSS-бары,
     без картинок — надёжно во всех почтовых клиентах).

Оформление — inline + <style> в духе email_html.py. Только stdlib + pandas.

CLI (превью):
    python peski_email_body.py --xlsx output/книга.xlsx --date 2026-07-09 --out preview.html
"""
from __future__ import annotations

import argparse
from html import escape
from pathlib import Path

import pandas as pd

# Переделы «пески», как в сводной по умолчанию (hronos.cli: Транспортировка + Подача).
SAND_PEREDELS = ["Подача песков", "Транспортировка песков"]
IDLE_PEREDEL = "Простой"
UNIT_ORDER = ["Дражный", "Обман", "Сайлык", "Талынья", "Эрел"]

COL_UNIT = "Подразделение"
COL_ISSUE = "Дата выдачи наряд-задания"
COL_FACT = "Дата. Факт"
COL_TIME = "Время"
COL_PER = "Передел"
COL_VOL = "Обьем работ, м3"
COL_NOTE = "Примечание"
COL_PRIB_MARK = "Марка промывочного прибора"
COL_PRIB_INV = "Инв. № промывочного прибора"

STYLE = (
    ".pe{font-family:-apple-system,Segoe UI,Arial,sans-serif;color:#24292f}"
    ".pe table{border-collapse:collapse;margin:6px 0}"
    ".pe th{padding:5px 9px;background:#305496;color:#fff;border:1px solid #d0d7de;"
    "text-align:left;white-space:nowrap;font-size:13px}"
    ".pe td{padding:3px 8px;border:1px solid #d0d7de;font-size:13px}"
    ".pe td.n{text-align:right;font-variant-numeric:tabular-nums}"
    ".pe td.t{background:#eef2f8;font-weight:bold}"
    ".pe td.z{background:#fafafa;color:#9aa6b2}"
    ".pe h2{margin:2px 0 2px;font-size:17px}"
    ".pe h3{margin:16px 0 4px;font-size:15px}"
    ".pe .sub{color:#6a737d;font-size:12px;margin:0 0 8px}"
)


def _fmt_int(x) -> str:
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except (ValueError, TypeError):
        return escape(str(x))


def _fmt_inv(x) -> str:
    """Инв.№ приходит как float (715.0) — показываем целым без хвоста."""
    s = str(x).strip()
    if s in ("", "nan", "None"):
        return ""
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except (ValueError, TypeError):
        pass
    return s


def _short_date(s: str) -> str:
    try:
        y, m, d = str(s)[:10].split("-")
        return f"{d}.{m}.{y}"
    except (ValueError, AttributeError):
        return str(s)


def _day_slice(df: pd.DataFrame, report_date: str) -> pd.DataFrame:
    issue = pd.to_datetime(df[COL_ISSUE], errors="coerce")
    return df[issue == pd.Timestamp(report_date)].copy()


# ── 1. Пески по подразделениям — как лист «Сводная_пески» (по часам) ──────────
def _sand_units(day: pd.DataFrame) -> list[str]:
    present = set(day[COL_UNIT].dropna().astype(str).str.strip())
    units = [u for u in UNIT_ORDER if u in present]
    units += [u for u in sorted(present) if u not in units]
    return units


def _table_units(day: pd.DataFrame) -> str:
    sand = day[day[COL_PER].astype(str).str.strip().isin(SAND_PEREDELS)].copy()
    if sand.empty:
        return '<p class="sub">За день нет данных по пескам.</p>'
    sand[COL_UNIT] = sand[COL_UNIT].astype(str).str.strip()
    sand["_fact"] = pd.to_datetime(sand[COL_FACT], errors="coerce")
    sand["_t"] = sand[COL_TIME].astype(str).str.strip()
    units = _sand_units(sand)

    def cells(vals, cls="n"):
        s = "".join(f'<td class="{cls}">{_fmt_int(v)}</td>' for v in vals)
        return s + f'<td class="{cls} t">{_fmt_int(sum(vals))}</td>'

    head = ("<tr><th>Дата / час</th>"
            + "".join(f"<th>{escape(u)}, м³</th>" for u in units)
            + "<th>Итого, м³</th></tr>")
    rows = []
    grand = [0.0] * len(units)
    for fact in sorted(sand["_fact"].dropna().unique()):
        day_df = sand[sand["_fact"] == fact]
        # подытог по дню = сумма по всем часам этого дня
        day_tot = [float(day_df[day_df[COL_UNIT] == u][COL_VOL].sum()) for u in units]
        for i, v in enumerate(day_tot):
            grand[i] += v
        dlabel = pd.Timestamp(fact).strftime("%d.%m.%Y")
        rows.append(f'<tr><td class="t">{dlabel}</td>{cells(day_tot, "n t")}</tr>')
        for t in sorted(day_df["_t"].dropna().unique()):
            hv = [float(day_df[(day_df["_t"] == t) & (day_df[COL_UNIT] == u)][COL_VOL].sum())
                  for u in units]
            rows.append(f"<tr><td>{escape(t)}</td>{cells(hv)}</tr>")
    rows.append(f'<tr><td class="t">Общий итог</td>{cells(grand, "n t")}</tr>')
    return f"<table>{head}{''.join(rows)}</table>"


# ── 2. Пески по промприборам — как «Сводная_пески», колонки = приборы ─────────
def _pribor_label(df: pd.DataFrame) -> pd.Series:
    mark = df[COL_PRIB_MARK].astype(str).str.strip()
    inv = df[COL_PRIB_INV].map(_fmt_inv)
    return (mark.where(mark.str.lower() != "nan", "") + " / " + inv).str.strip(" /")


def _table_pribor(day: pd.DataFrame) -> str:
    sand = day[day[COL_PER].astype(str).str.strip().isin(SAND_PEREDELS)].copy()
    if sand.empty:
        return ""
    sand[COL_UNIT] = sand[COL_UNIT].astype(str).str.strip()
    sand["_fact"] = pd.to_datetime(sand[COL_FACT], errors="coerce")
    sand["_t"] = sand[COL_TIME].astype(str).str.strip()
    sand["_prib"] = _pribor_label(sand)
    sand = sand[sand["_prib"].str.replace("/", "").str.strip() != ""]
    if sand.empty:
        return ""

    # колонки: (подразделение, прибор), приборы внутри подразделения по объёму
    tot = sand.groupby([COL_UNIT, "_prib"])[COL_VOL].sum()
    units = _sand_units(sand)
    columns: list[tuple[str, str]] = []
    for u in units:
        pribs = sorted({p for (uu, p) in tot.index if uu == u},
                       key=lambda p: -float(tot.get((u, p), 0)))
        columns += [(u, p) for p in pribs]

    # многоярусная шапка: подразделение (span) → прибор
    h1 = '<tr><th rowspan="2">Дата / час</th>'
    for u in units:
        n = sum(1 for (uu, _) in columns if uu == u)
        if n:
            h1 += f'<th colspan="{n}">{escape(u)}</th>'
    h1 += '<th rowspan="2">Итого, м³</th></tr>'
    h2 = "<tr>" + "".join(f"<th>{escape(p)}</th>" for (_, p) in columns) + "</tr>"

    def cells(valmap, cls="n"):
        vals = [valmap.get(col, 0.0) for col in columns]
        s = "".join(f'<td class="{cls}">{_fmt_int(v)}</td>' for v in vals)
        return s + f'<td class="{cls} t">{_fmt_int(sum(vals))}</td>'

    rows = []
    grand: dict[tuple[str, str], float] = {}
    for fact in sorted(sand["_fact"].dropna().unique()):
        dd = sand[sand["_fact"] == fact]
        pv = dd.groupby([COL_UNIT, "_prib", "_t"])[COL_VOL].sum()
        daymap = {col: float(dd[(dd[COL_UNIT] == col[0]) & (dd["_prib"] == col[1])][COL_VOL].sum())
                  for col in columns}
        for col, v in daymap.items():
            grand[col] = grand.get(col, 0.0) + v
        dlabel = pd.Timestamp(fact).strftime("%d.%m.%Y")
        rows.append(f'<tr><td class="t">{dlabel}</td>{cells(daymap, "n t")}</tr>')
        for t in sorted(dd["_t"].dropna().unique()):
            hmap = {col: float(pv.get((col[0], col[1], t), 0.0)) for col in columns}
            rows.append(f"<tr><td>{escape(t)}</td>{cells(hmap)}</tr>")
    rows.append(f'<tr><td class="t">Общий итог</td>{cells(grand, "n t")}</tr>')
    return f"<table>{h1}{h2}{''.join(rows)}</table>"


# ── 3. Причины простоя с начала промывочного сезона + диаграмма ───────────────
def _idle_reasons(df: pd.DataFrame, top: int = 10) -> tuple[pd.Series, pd.Timestamp | None]:
    """Причины простоя по всему реестру (сезон), топ по часам. Возвращает (серия, дата начала)."""
    idle = df[df[COL_PER].astype(str).str.strip() == IDLE_PEREDEL].copy()
    if idle.empty:
        return pd.Series(dtype=int), None
    reason = (idle[COL_NOTE].astype(str).str.strip()
              .replace({"nan": "—", "None": "—", "": "—"}))
    reason = reason[reason.str.lower() != "обед"]          # обед (плановый) не показываем
    ser = reason.groupby(reason).size().sort_values(ascending=False).head(top)
    start = pd.to_datetime(idle[COL_FACT], errors="coerce").min()
    return ser, start


def _table_idle(reasons: pd.Series) -> str:
    if reasons.empty:
        return '<p class="sub">За день простоев не зафиксировано.</p>'
    head = "<tr><th>Причина простоя</th><th>Часы</th></tr>"
    rows = [f"<tr><td>{escape(str(r))}</td><td class=\"n\">{_fmt_int(h)}</td></tr>"
            for r, h in reasons.items()]
    rows.append(f'<tr><td class="t">Итого</td>'
                f'<td class="n t">{_fmt_int(int(reasons.sum()))}</td></tr>')
    return f"<table>{head}{''.join(rows)}</table>"


def _chart_idle(reasons: pd.Series) -> str:
    """Горизонтальный бар-чарт (не таблица): подпись, полоса на дорожке, значение."""
    if reasons.empty:
        return ""
    mx = int(reasons.max()) or 1
    rows = []
    for r, h in reasons.items():
        pct = max(3, int(round(100 * h / mx)))
        rows.append(
            '<div style="margin:8px 0">'
            f'<div style="font-size:12px;color:#24292f;margin-bottom:3px">{escape(str(r))}</div>'
            '<span style="display:inline-block;width:360px;max-width:72%;height:15px;'
            'background:#eef2f8;border-radius:3px;vertical-align:middle">'
            f'<span style="display:block;height:15px;width:{pct}%;background:#305496;'
            'border-radius:3px"></span></span>'
            f'<span style="color:#305496;font-weight:bold;font-size:12px;margin-left:8px;'
            f'vertical-align:middle">{_fmt_int(h)} ч</span>'
            "</div>"
        )
    return '<div style="margin:4px 0 2px">' + "".join(rows) + "</div>"


def build_html(df: pd.DataFrame, report_date: str) -> str:
    day = _day_slice(df, report_date)
    reasons, season_start = _idle_reasons(df)
    season_lbl = (f" с начала сезона (с {season_start.strftime('%d.%m.%Y')})"
                  if season_start is not None else " с начала сезона")
    parts = [
        f"<style>{STYLE}</style>",
        '<div class="pe">',
        f"<h2>Хронометраж транспортировки песков</h2>",
        f'<p class="sub">Отчёт за {escape(_short_date(report_date))}</p>',
        "<h3>1. Пески по подразделениям</h3>",
        _table_units(day),
        "<h3>2. Пески по промывочным приборам</h3>",
        _table_pribor(day) or '<p class="sub">Нет данных по приборам.</p>',
        f"<h3>3. Причины простоя{season_lbl}</h3>",
        _table_idle(reasons),
    ]
    chart = _chart_idle(reasons)
    if chart:
        parts += ["<h3>Диаграмма причин простоя</h3>", chart]
    parts.append("</div>")
    return "".join(parts)


def build_text(df: pd.DataFrame, report_date: str) -> str:
    day = _day_slice(df, report_date)
    lines = [f"Хронометраж транспортировки песков — отчёт за {_short_date(report_date)}",
             "", "Таблицы в HTML-версии письма; при необходимости откройте вложение."]
    return "\n".join(lines)


def build_body(df: pd.DataFrame, report_date: str) -> tuple[str, str]:
    return build_html(df, report_date), build_text(df, report_date)


def _main() -> int:
    ap = argparse.ArgumentParser(description="Превью тела письма из финального реестра")
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--date", required=True, help="Дата выдачи наряд-задания YYYY-MM-DD")
    ap.add_argument("--sheet", default="Сводный_Реестр")
    ap.add_argument("--out", type=Path, default=Path("preview.html"))
    args = ap.parse_args()
    df = pd.read_excel(args.xlsx, sheet_name=args.sheet)
    body = build_html(df, args.date)
    # автономный файл-превью: тело письмо получает charset из MIME, а для просмотра
    # в браузере оборачиваем в полноценный документ с указанием кодировки.
    html = f'<!doctype html><html><head><meta charset="utf-8"></head><body>{body}</body></html>'
    args.out.write_text(html, encoding="utf-8")
    print(f"[OK] {args.out} ({len(html.encode('utf-8')) / 1024:.1f} КБ)")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
