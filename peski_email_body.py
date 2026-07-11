"""Тело письма «Хронометраж транспортировки торфов и песков» из финального реестра.

Живой (параллельный) пайплайн раньше слал письмо с пустым телом (--empty-body).
Этот модуль строит компактное HTML-тело за отчётный день (Дата выдачи наряд-задания
== report_date, как в сводной), чтобы не открывать вложение:

  1. Промывка песков по подразделениям (по часам, как «Сводная_пески»).
  2. Детализация по промывочным приборам (по часам; красным — час ниже нормы >20%).
  3. Причины отклонений за последний час: прибор/подразделение, факт, норма, причина.

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
    ".pe td.bad{background:#f8caca;color:#8a1f1f;font-weight:bold}"
    ".pe td.nowrap{white-space:nowrap}"
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


# Норма выработки промприбора, м³/час (по марке).
# СБ-2.1/СБ2.1, ГИТ, СБ-1.7 → 120; ПБШ, ПКБШ → 60. Прочие (напр. ГГМ) — без нормы.
NORM_120 = 120
NORM_60 = 60
DEVIATION_THRESHOLD = 0.20   # отклонение вниз более чем на 20% → подсветка


def _pribor_norm(mark: str) -> int | None:
    m = str(mark).upper().replace(" ", "").replace(" ", "").replace("-", "").replace(".", "")
    if m.startswith("ПКБШ") or m.startswith("ПБШ"):
        return NORM_60
    if m.startswith("СБ") or m.startswith("ГИТ"):
        return NORM_120
    return None


def _mark_of(prib_label: str) -> str:
    return str(prib_label).split(" / ", 1)[0].strip()


def _is_below_norm(value: float, norm: int | None) -> bool:
    return norm is not None and 0 < value < norm * (1 - DEVIATION_THRESHOLD)


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

    norms = [_pribor_norm(_mark_of(p)) for (_, p) in columns]

    # многоярусная шапка: подразделение (span) → прибор
    h1 = '<tr><th rowspan="2">Дата / час</th>'
    for u in units:
        n = sum(1 for (uu, _) in columns if uu == u)
        if n:
            h1 += f'<th colspan="{n}">{escape(u)}</th>'
    h1 += '<th rowspan="2">Итого, м³</th></tr>'
    h2 = "<tr>" + "".join(f"<th>{escape(p)}</th>" for (_, p) in columns) + "</tr>"

    def cells(valmap, cls="n", highlight=False):
        out = []
        for i, col in enumerate(columns):
            v = valmap.get(col, 0.0)
            klass = cls
            if highlight and _is_below_norm(v, norms[i]):
                klass = (cls + " bad").strip()
            out.append(f'<td class="{klass}">{_fmt_int(v)}</td>')
        total = sum(valmap.get(col, 0.0) for col in columns)
        out.append(f'<td class="{cls} t">{_fmt_int(total)}</td>')
        return "".join(out)

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
            rows.append(f'<tr><td>{escape(t)}</td>{cells(hmap, highlight=True)}</tr>')
    rows.append(f'<tr><td class="t">Общий итог</td>{cells(grand, "n t")}</tr>')
    note = ('<p class="sub">Красным — час, когда прибор отработал ниже нормы более чем '
            f'на {int(DEVIATION_THRESHOLD * 100)}% (норма 120 м³/ч — СБ, ГИТ; 60 м³/ч — ПБШ, ПКБШ).</p>')
    return f"<table>{h1}{h2}{''.join(rows)}</table>{note}"


# ── 3. Причины простоя по приборам: под каждым прибором «Причина» + «м³» ──────
def _table_idle_grid(day: pd.DataFrame) -> str:
    d = day.copy()
    d[COL_UNIT] = d[COL_UNIT].astype(str).str.strip()
    d["_fact"] = pd.to_datetime(d[COL_FACT], errors="coerce")
    d["_t"] = d[COL_TIME].astype(str).str.strip()
    d["_prib"] = _pribor_label(d)

    idle = d[d[COL_PER].astype(str).str.strip() == IDLE_PEREDEL].copy()
    idle["_reason"] = idle[COL_NOTE].astype(str).str.strip().replace({"nan": "", "None": ""})
    idle = idle[(idle["_reason"] != "") & (idle["_reason"].str.lower() != "обед")]
    idle = idle[idle["_prib"].str.replace("/", "").str.strip() != ""]     # только с прибором
    idle = idle.dropna(subset=["_fact"])
    idle = idle[idle["_t"].str.match(r"^\d{1,2}:\d{2}$")]
    if idle.empty:
        return '<p class="sub">Простоев с указанием прибора за отчёт нет.</p>'

    # колонки: подразделение → прибор (как в п.2), среди приборов с простоем
    units = _sand_units(idle)
    pribs_by_unit = {u: sorted({p for uu, p in zip(idle[COL_UNIT], idle["_prib"]) if uu == u})
                     for u in units}
    columns = [(u, p) for u in units for p in pribs_by_unit[u]]

    # объём песков по (Дата.Факт, час, подразделение, прибор) — для колонки «м³»
    sand = d[d[COL_PER].astype(str).str.strip().isin(SAND_PEREDELS)]
    volmap = sand.groupby(["_fact", "_t", COL_UNIT, "_prib"])[COL_VOL].sum()
    rmap = (idle.groupby(["_fact", "_t", COL_UNIT, "_prib"])["_reason"]
            .apply(lambda s: "; ".join(dict.fromkeys(s))))

    # трёхъярусная шапка: подразделение → прибор → (Причина | м³)
    h1 = '<tr><th rowspan="3">Дата / час</th>'
    for u in units:
        h1 += f'<th colspan="{2 * len(pribs_by_unit[u])}">{escape(u)}</th>'
    h1 += "</tr>"
    h2 = "<tr>" + "".join(f'<th colspan="2">{escape(p)}</th>' for (_, p) in columns) + "</tr>"
    h3 = "<tr>" + "".join('<th>Причина простоя</th><th>м³</th>' for _ in columns) + "</tr>"

    rows = []
    for fact in sorted(idle["_fact"].dropna().unique()):
        sub = idle[idle["_fact"] == fact]
        dshort = pd.Timestamp(fact).strftime("%d.%m")
        for t in sorted(sub["_t"].unique()):
            tds = []
            for (u, p) in columns:
                r = rmap.get((fact, t, u, p), "")
                v = float(volmap.get((fact, t, u, p), 0.0))
                rcell = f'<td class="bad">{escape(r)}</td>' if r else "<td></td>"
                vcell = f'<td class="n">{_fmt_int(v)}</td>' if v else '<td class="n z"></td>'
                tds.append(rcell + vcell)
            rows.append(f'<tr><td class="nowrap">{t} · {dshort}</td>{"".join(tds)}</tr>')
    return f"<table>{h1}{h2}{h3}{''.join(rows)}</table>"


def build_html(df: pd.DataFrame, report_date: str) -> str:
    day = _day_slice(df, report_date)
    parts = [
        f"<style>{STYLE}</style>",
        '<div class="pe">',
        f"<h2>Хронометраж транспортировки песков</h2>",
        f'<p class="sub">Отчёт за {escape(_short_date(report_date))}</p>',
        "<h3>1. Промывка песков по подразделениям</h3>",
        _table_units(day),
        "<h3>2. Детализация по промывочным приборам</h3>",
        _table_pribor(day) or '<p class="sub">Нет данных по приборам.</p>',
        "<h3>3. Причины простоя по приборам</h3>",
        _table_idle_grid(day),
        "</div>",
    ]
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
