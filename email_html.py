"""Рендер HTML- и текстового тела письма из metrics + записки.
Операционный отчёт за ТЕКУЩИЕ сутки (граница смен 07:00) нарастающим итогом,
с выделением прошедшего часа: KPI по подразделениям → отклонения → записка →
почасовка (актуализируется по часам) → динамика час×дата (по предприятию и по
подразделениям) → песок по приборам → откатка → причины простоя.
Оформление — в <style> (классы), чтобы письмо с детализацией влезало в лимит Gmail
~102 КБ; фолбэк — text/plain. Stdlib.
"""
from __future__ import annotations
from html import escape

# Оформление вынесено в <style> (Gmail поддерживает <style> в <head>), чтобы ячейки
# таблиц были компактными — иначе письмо с детализацией по подразделениям превышает
# лимит Gmail ~102 КБ и обрезается. Фолбэк — text/plain.
STYLE = (
    ".rpt table{border-collapse:collapse;margin:6px 0}"
    ".rpt th{padding:5px 9px;background:#305496;color:#fff;border:1px solid #d0d7de;"
    "text-align:left;white-space:nowrap;font-size:13px}"
    ".rpt td{padding:3px 8px;border:1px solid #d0d7de;font-size:13px}"
    ".rpt td.a{background:#f6f8fa}.rpt td.t{background:#eef2f8;font-weight:bold}"
    ".rpt td.r{background:#fff5f5}.rpt td.g{background:#f3fbf4}"
    ".rpt td.hl{background:#fff3cd}.rpt td.w{background:#fff8e6}"
    ".rpt td.z{background:#fafafa;color:#9aa6b2}"
    ".rpt h3{margin:16px 0 4px}.rpt h4{margin:12px 0 2px}.rpt h5{margin:8px 0 2px}"
)
# карта цвет-фона → CSS-класс (совместимо с прежними вызовами _row(..., bg))
_BGC = {"#ffffff": "", "#f6f8fa": "a", "#eef2f8": "t", "#fff5f5": "r",
        "#f3fbf4": "g", "#fff3cd": "hl", "#fff8e6": "w", "#fafafa": "z"}


def _fmt_int(x) -> str:
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except (ValueError, TypeError):
        return str(x)


def _short_date(s: str) -> str:
    try:
        y, m, d = str(s).split("-")
        return f"{d}.{m}"
    except (ValueError, AttributeError):
        return str(s)


def _shift_lbl(hour_str: str) -> str:
    try:
        h = int(str(hour_str).split(":")[0])
    except (ValueError, IndexError):
        return ""
    return "1 смена" if 7 <= h < 20 else "2 смена"


def _th(*cols) -> str:
    return "<tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>"


def _row(cells, bg="#ffffff") -> str:
    cls = _BGC.get(bg, "")
    td = f'<td class="{cls}">' if cls else "<td>"
    return "<tr>" + "".join(f"{td}{c}</td>" for c in cells) + "</tr>"


# ── 1. KPI по подразделениям за сутки ─────────────────────────────────────────
def _kpi_by_unit(m: dict) -> str:
    rows = m.get("by_unit_day", [])
    if not rows:
        return ""
    head = _th("Подразделение", "Торф, м³", "Пески, м³", "Кол-во машин торф, шт",
               "Кол-во машин пески, шт", "Простои, ч")
    body = ""
    tot = {"vt": 0, "vp": 0, "mt": 0, "mp": 0, "id": 0}
    for i, u in enumerate(rows):
        bg = "#f6f8fa" if i % 2 else "#ffffff"
        body += _row([escape(u["unit"]), _fmt_int(u["vol_torf"]), _fmt_int(u["vol_pesok"]),
                      _fmt_int(u["mach_torf"]), _fmt_int(u["mach_pesok"]),
                      _fmt_int(u["idle_h"])], bg)
        tot["vt"] += u["vol_torf"]; tot["vp"] += u["vol_pesok"]
        tot["mt"] += u["mach_torf"]; tot["mp"] += u["mach_pesok"]; tot["id"] += u["idle_h"]
    body += _row(["Итого", _fmt_int(tot["vt"]), _fmt_int(tot["vp"]),
                  _fmt_int(tot["mt"]), _fmt_int(tot["mp"]), _fmt_int(tot["id"])], "#eef2f8")
    return ('<h3>1. Ключевые показатели за сутки — по подразделениям</h3>'
            f'<table>{head}{body}</table>')


# ── 2. План/факт по пескам (промывочные приборы) ──────────────────────────────
def _window_label(m: dict) -> str:
    oh = m.get("op_hours") or []
    return f"{oh[0]}–{oh[-1]}" if oh else "сутки"


def _plan_fact_block(m: dict) -> str:
    rows = m.get("plan_fact", [])
    if not rows:
        return ""
    win = _window_label(m)
    head = _th("Подразделение", f"Текущий объём ({win}), м³", "Средний за 7 дн, м³",
               "Плановый объём, м³", "Ожидаемый за сутки, м³",
               "% плана (факт)", "% плана (прогноз)", "Статус")
    body = ""
    for d in sorted(rows, key=lambda x: x["unit"]):
        plan, pf, pj = d["plan"], d.get("pct_fact"), d.get("pct_proj")
        if not plan:
            status, bg = "— нет плана", "#ffffff"
        elif pj is not None and pj < 90:
            status, bg = f"⚠ прогноз {pj:.0f}% плана", "#fff5f5"
        else:
            status = f"✓ прогноз {pj:.0f}% плана" if pj is not None else "✓"
            bg = "#f3fbf4"
        body += _row([escape(d["unit"]), _fmt_int(d["cur"]), _fmt_int(d["avg7"]),
                      _fmt_int(plan) if plan else "—", _fmt_int(d["expected"]),
                      f"{pf:.0f}%" if pf is not None else "—",
                      f"{pj:.0f}%" if pj is not None else "—", escape(status)], bg)
    return ('<h3>2. План/факт по пескам (промывочные приборы)</h3>'
            '<div style="color:#57606a;font-size:13px;">План = суточная производительность '
            'приборов (СБ-2.1/ГИТ-62 — 2400 м³, ПБШ-100/СБ-1.7 — 1200) × число приборов '
            f'подразделения. Текущий — накоплено за <b>{escape(win)}</b> текущих суток; '
            'ожидаемый — средний темп текущих суток × 24 ч; средний — за предыдущие 7 дней.'
            '</div>'
            f'<table>{head}{body}</table>')


# ── 3. Пояснительная записка ──────────────────────────────────────────────────
def _note_block(note: str) -> str:
    note_html = "".join(f"<p>{escape(p)}</p>" for p in note.split("\n\n") if p.strip())
    return f"<h3>Пояснительная записка</h3>{note_html}"


# ── 4. Почасовая раскладка по подразделениям (актуализируется по часам) ────────
def _hourly_blocks(m: dict) -> str:
    rows = m.get("hourly_unit", [])
    if not rows:
        return ""
    last_h = (m.get("last_hour_kpi") or {}).get("hour")
    units: dict[str, list] = {}
    for r in rows:
        units.setdefault(r["unit"], []).append(r)
    out = ['<h3>3. Почасовая раскладка за сутки по подразделениям</h3>']
    for unit in sorted(units):
        head = _th("Смена", "Час", "Торф", "Пески", "Причина простоя")
        body = ""
        for r in units[unit]:
            t, p = r.get("torf", 0), r.get("pesok", 0)
            reason = escape(r.get("reason") or "")
            shift = escape(r.get("shift") or _shift_lbl(r["hour"]))
            if r["hour"] == last_h:
                bg = "#fff3cd"
            elif t == 0 and p == 0 and reason:
                bg = "#fff8e6"
            elif t == 0 and p == 0:
                bg = "#fafafa"
            else:
                bg = "#ffffff"
            body += _row([shift, r["hour"], _fmt_int(t) if t else "—",
                          _fmt_int(p) if p else "—", reason or ""], bg)
        out.append(f'<h4>{escape(unit)}</h4>'
                   f'<table>{head}{body}</table>')
    return "".join(out)


# ── 4–5. Динамика машин «час × дата» по материалам ────────────────────────────
def _matrix_table(d: dict, label: str, tag: str = "h4") -> str:
    """Одна матрица «час × дата» (пустые часы пропускаются); '' если данных нет."""
    if not d or not d.get("dates"):
        return ""
    dates = d["dates"]
    body = ""
    shown = 0
    for r in d["rows"]:
        vals = [int(r.get(c, 0) or 0) for c in dates]
        if not any(vals):
            continue
        shown += 1
        cells = [r["hour"]] + [_fmt_int(v) if v else "—" for v in vals]
        body += _row(cells, "#ffffff")
    if not shown:
        return ""
    head = _th("Час", *[_short_date(c) for c in dates])
    return f"<{tag}>{label}</{tag}><table>{head}{body}</table>"


def _dyn_tables(dyn: dict, tag: str = "h4") -> str:
    """Таблицы «час × дата» для торфа и песков (без склада); пустые часы пропускаются."""
    return "".join(_matrix_table(dyn.get(mat), label, tag)
                   for mat, label in (("Торф", "Торф"), ("Песок", "Пески")))


def _dynamics_block(m: dict) -> str:
    tables = _dyn_tables(m.get("mach_dynamics", {}))
    if not tables:
        return ""
    return ('<h3>4. Динамика по часам и дням — кол-во машин (по предприятию)</h3>'
            '<div style="color:#57606a;font-size:13px;">Строки — время (операционные сутки '
            '07:00→07:00), столбцы — даты (последние 7). Отдельно торф и пески.</div>'
            + tables)


def _dynamics_unit_block(m: dict) -> str:
    dyn_u = m.get("mach_dynamics_unit", {})
    if not dyn_u:
        return ""
    out = ['<h3>5. Динамика по часам и дням — в разрезе подразделений (кол-во машин)</h3>']
    has = False
    for unit in sorted(dyn_u):
        tables = _dyn_tables(dyn_u[unit], tag="h5")
        if not tables:
            continue
        has = True
        out.append(f'<h4>{escape(unit)}</h4>{tables}')
    return "".join(out) if has else ""


# ── 6. Песок в разрезе промывочных приборов ───────────────────────────────────
def _pesok_devices_block(m: dict) -> str:
    pdv = m.get("pesok_devices", {})
    if not pdv:
        return ""
    out = ['<h3>6. Пески — по промывочным приборам (машины/час)</h3>']
    has_any = False
    for unit in sorted(pdv):
        devices = pdv[unit]["devices"]
        rows = pdv[unit]["rows"]
        if not devices:
            continue
        has_any = True
        head = _th("Час", *[escape(d) for d in devices], "Итого")
        body = ""
        for r in rows:
            vals = [int(round(float(r.get(d, 0) or 0))) for d in devices]
            tot = sum(vals)
            bg = "#ffffff" if tot else "#fafafa"
            cells = [r["Час"]] + [_fmt_int(v) if v else "—" for v in vals] + \
                    [f"<b>{_fmt_int(tot)}</b>" if tot else "—"]
            body += _row(cells, bg)
        out.append(f'<h4>{escape(unit)}</h4>'
                   f'<table>{head}{body}</table>')
    return "".join(out) if has_any else ""


# ── 7. Вывоз песков на склад (без названия прибора) ───────────────────────────
def _sklad_block(m: dict) -> str:
    dyn = m.get("mach_dynamics", {}) or {}
    dyn_u = m.get("mach_dynamics_unit", {}) or {}
    overall = _matrix_table(dyn.get("Песок_склад"), "По предприятию", "h4")
    unit_tables = []
    for unit in sorted(dyn_u):
        t = _matrix_table((dyn_u[unit] or {}).get("Песок_склад"), escape(unit), "h5")
        if t:
            unit_tables.append(t)
    if not overall and not unit_tables:
        return ""
    body = overall
    if unit_tables:
        body += "<h4>По подразделениям</h4>" + "".join(unit_tables)
    return ('<h3>7. Вывоз песков на склад (машины, час×дата)</h3>'
            '<div style="color:#57606a;font-size:13px;">Пески без названия промывочного '
            'прибора — вывоз на склад; в «Пески» (KPI, почасовка, динамика, приборы) и в '
            'план/факт не входят.</div>' + body)


# ── 8. Аналитика откатки ──────────────────────────────────────────────────────
def _otkatka_block(m: dict) -> str:
    units = m.get("otkatka_unit", [])
    dyn = m.get("otkatka_dyn", {}) or {}
    if not units and not dyn.get("rows"):
        return ""
    out = ['<h3>8. Аналитика откатки (среднее расстояние транспортировки, м)</h3>'
           '<div style="color:#57606a;font-size:13px;">Откатка — дистанция транспортировки; '
           'среднее взвешено по числу машин. Длиннее откатка → дороже и дольше рейс.</div>']
    if units:
        head = _th("Подразделение", "Торф, м", "Пески, м")
        body = ""
        for u in sorted(units, key=lambda x: x["unit"]):
            body += _row([escape(u["unit"]),
                          _fmt_int(u["torf"]) if u["torf"] is not None else "—",
                          _fmt_int(u["pesok"]) if u["pesok"] is not None else "—"])
        out.append('<h4>За текущие сутки</h4>'
                   f'<table>{head}{body}</table>')
    us = dyn.get("units", [])
    rows = dyn.get("rows", [])
    if rows and us:
        head = _th("Дата", *[escape(u) for u in us])
        body = ""
        for r in rows:
            cells = [_short_date(r["Дата"])] + \
                    [_fmt_int(r.get(u)) if r.get(u) is not None else "—" for u in us]
            body += _row(cells)
        out.append('<h4>Динамика по дням</h4>'
                   f'<table>{head}{body}</table>')
    return "".join(out)


# ── 8. Аналитика причин простоя (сегодня + динамика) ──────────────────────────
def _idle_dyn_table(m: dict) -> str:
    dyn = m.get("idle_dyn", [])
    if not dyn:
        return ""
    head = _th("Дата", "Плановые, ч", "Внеплановые, ч", "Итого, ч")
    body = ""
    for r in dyn:
        bg = "#fff5f5" if r["unplanned"] else "#ffffff"
        body += _row([_short_date(r["date"]), _fmt_int(r["planned"]),
                      f'<span style="color:#d1242f;">{_fmt_int(r["unplanned"])}</span>'
                      if r["unplanned"] else "0", _fmt_int(r["total"])], bg)
    return ('<h4>Динамика простоев по дням (по предприятию)</h4>'
            f'<table>{head}{body}</table>')


def _idle_reasons_block(m: dict) -> str:
    rows = m.get("idle_reasons", [])
    dyn_html = _idle_dyn_table(m)
    if not rows:
        return ('<h3>9. Аналитика причин простоя</h3>'
                '<p>Простои за текущие сутки не зафиксированы.</p>' + dyn_html)
    planned = sum(r["hours"] for r in rows if r["kind"] == "плановый")
    unplanned = sum(r["hours"] for r in rows if r["kind"] == "внеплановый")
    head = _th("Подразделение", "Причина", "Тип", "Часов")
    body = ""
    for r in rows:
        bg = "#fff5f5" if r["kind"] == "внеплановый" else "#f6f8fa"
        body += _row([escape(r["unit"]), escape(r["reason"]), escape(r["kind"]),
                      _fmt_int(r["hours"])], bg)
    summary = (f'<p style="margin:6px 0;">За текущие сутки простоев: '
               f'<b>{_fmt_int(planned + unplanned)} ч</b> · '
               f'плановые (обед/пересменка/ЕТО) — {_fmt_int(planned)} ч · '
               f'<span style="color:#d1242f;">внеплановые (поломки/нет напряжения) — '
               f'{_fmt_int(unplanned)} ч</span>.</p>')
    return (f'<h3>9. Аналитика причин простоя</h3>{summary}'
            f'<h4>За текущие сутки</h4>'
            f'<table>{head}{body}</table>'
            + dyn_html)


def _header(m: dict) -> str:
    rd = escape(str(m.get("report_date", "")))
    lh = m.get("last_hour_kpi") or {}
    line = f'Отчёт за сутки: <b>{rd}</b>'
    if m.get("partial"):
        line += (' <span style="color:#57606a;">(текущие операционные сутки, '
                 'нарастающим итогом)</span>')
    if lh.get("hour"):
        line += f' · прошедший час <b>{escape(lh["hour"])}–{escape(lh.get("hour_to", ""))}</b>'
    per = m.get("period", ["", ""])
    line += (f'<div style="color:#8c959f;font-size:12px;">обновлено по данным до '
             f'{escape(str(per[1]))}</div>')
    return f'<div style="color:#1f2328;">{line}</div>'


def build_html(m: dict, note: str, alerts: list[dict]) -> str:
    return (
        '<!DOCTYPE html><html lang="ru"><head>'
        '<meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1">'
        f'<style>{STYLE}</style>'
        '</head><body>'
        '<div class="rpt" style="font-family:-apple-system,Segoe UI,Arial,sans-serif;color:#1f2328;max-width:860px;">'
        '<h2 style="margin:0 0 4px;">Хронометраж транспортировки торфов и песков</h2>'
        + _header(m)
        + _kpi_by_unit(m)
        + _plan_fact_block(m)
        + _note_block(note)
        + _hourly_blocks(m)
        + _dynamics_block(m)
        + _dynamics_unit_block(m)
        + _pesok_devices_block(m)
        + _sklad_block(m)
        + _otkatka_block(m)
        + _idle_reasons_block(m)
        + '<p style="color:#8c959f;font-size:12px;margin-top:16px;">'
          'Детализация — во вложении (Excel: реестр + листы аналитики, почасовка, причины простоя). '
          'Автоматическая рассылка хронометража.</p>'
        '</div></body></html>'
    )


def build_text(m: dict, note: str, alerts: list[dict]) -> str:
    lh = m.get("last_hour_kpi") or {}
    head = f"Хронометраж торф/пески — отчёт за сутки {m.get('report_date', '')}"
    if lh.get("hour"):
        head += f" · прошедший час {lh['hour']}–{lh.get('hour_to', '')}"
    lines = [head, ""]
    lines.append("KPI по подразделениям (торф/пески, м³ | машин | простои ч):")
    for u in m.get("by_unit_day", []):
        lines.append(f"  {u['unit']}: торф {_fmt_int(u['vol_torf'])} / пески {_fmt_int(u['vol_pesok'])} м³ · "
                     f"машин {_fmt_int(u['mach_torf'])}/{_fmt_int(u['mach_pesok'])} · простои {u['idle_h']} ч")
    lines += ["", "План/факт по пескам (текущий / план / ожидаемый, м³):"]
    for d in m.get("plan_fact", []):
        pj = d.get("pct_proj")
        tail = f" · прогноз {pj:.0f}% плана" if pj is not None else ""
        lines.append(f"  {d['unit']}: {_fmt_int(d['cur'])} / {_fmt_int(d['plan'])} / "
                     f"{_fmt_int(d['expected'])}{tail}")
    lines += ["", "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА:", note, "",
              "Почасовка, динамика, причины простоя — см. HTML-письмо / Excel."]
    return "\n".join(lines)
