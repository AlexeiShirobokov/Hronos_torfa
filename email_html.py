"""Рендер HTML- и текстового тела письма из metrics + записки.
Операционный отчёт за прошедшие сутки в разрезе подразделений:
KPI по подразделениям → отклонения → записка → почасовка 24ч → причины простоя.
Только инлайн-стили (почтовые клиенты режут <style>). Зависимости — stdlib.
"""
from __future__ import annotations
from html import escape

TH = ("padding:6px 10px;background:#305496;color:#fff;border:1px solid #d0d7de;"
      "text-align:left;white-space:nowrap;")
TD = "padding:5px 10px;border:1px solid #d0d7de;"


def _fmt_int(x) -> str:
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except (ValueError, TypeError):
        return str(x)


def _th(*cols) -> str:
    return "<tr>" + "".join(f'<th style="{TH}">{c}</th>' for c in cols) + "</tr>"


def _row(cells, bg="#ffffff") -> str:
    return "<tr>" + "".join(f'<td style="{TD}background:{bg};">{c}</td>' for c in cells) + "</tr>"


# ── 1. KPI по подразделениям за сутки ─────────────────────────────────────────
def _kpi_by_unit(m: dict) -> str:
    rows = m.get("by_unit_day", [])
    if not rows:
        return ""
    head = _th("Подразделение", "Торф, м³", "Песок, м³", "Машин торф",
               "Машин песок", "Простои, ч")
    body = ""
    tot = {"vt": 0, "vp": 0, "mt": 0, "mp": 0, "id": 0}
    for i, u in enumerate(rows):
        bg = "#f6f8fa" if i % 2 else "#ffffff"
        body += _row([escape(u["unit"]), _fmt_int(u["vol_torf"]), _fmt_int(u["vol_pesok"]),
                      _fmt_int(u["mach_torf"]), _fmt_int(u["mach_pesok"]),
                      _fmt_int(u["idle_h"])], bg)
        tot["vt"] += u["vol_torf"]; tot["vp"] += u["vol_pesok"]
        tot["mt"] += u["mach_torf"]; tot["mp"] += u["mach_pesok"]; tot["id"] += u["idle_h"]
    body += ("<tr style='font-weight:bold;'>" + "".join(
        f'<td style="{TD}background:#eef2f8;">{c}</td>' for c in
        ["Итого", _fmt_int(tot["vt"]), _fmt_int(tot["vp"]),
         _fmt_int(tot["mt"]), _fmt_int(tot["mp"]), _fmt_int(tot["id"])]) + "</tr>")
    return (f'<h3>1. Ключевые показатели за сутки — по подразделениям</h3>'
            f'<table style="border-collapse:collapse;margin:8px 0;">{head}{body}</table>')


# ── 2. Отклонения по подразделениям ───────────────────────────────────────────
def _dev_flags(d: dict) -> list[str]:
    flags = []
    base_v, vol = d.get("vol_base", 0), d.get("vol", 0)
    if base_v > 0:
        dv = (vol - base_v) / base_v * 100
        if dv <= -20:
            flags.append(f"объём ниже нормы на {abs(dv):.0f}%")
    base_i, idle = d.get("idle_base", 0), d.get("idle", 0)
    if idle > 0 and (base_i == 0 or idle > base_i * 1.5):
        flags.append(f"простои выше нормы на {(idle/base_i-1)*100:.0f}%" if base_i > 0
                     else f"простои {idle} ч (раньше не было)")
    return flags


def _dev_block(m: dict) -> str:
    rows = m.get("unit_dev", [])
    if not rows:
        return ""
    head = _th("Подразделение", "Объём / норма, м³", "Простои / норма, ч", "Статус")
    body = ""
    for d in sorted(rows, key=lambda x: x["unit"]):
        flags = _dev_flags(d)
        bg = "#fff5f5" if flags else "#f3fbf4"
        status = ("⚠ " + "; ".join(flags)) if flags else "✓ в норме"
        body += _row([escape(d["unit"]),
                      f'{_fmt_int(d["vol"])} / {_fmt_int(d["vol_base"])}',
                      f'{_fmt_int(d["idle"])} / {_fmt_int(d["idle_base"])}',
                      escape(status)], bg)
    return (f'<h3>2. Внимание — отклонения по подразделениям (vs медиана 7 дней)</h3>'
            f'<table style="border-collapse:collapse;margin:8px 0;">{head}{body}</table>')


# ── 3. Пояснительная записка ──────────────────────────────────────────────────
def _note_block(note: str) -> str:
    note_html = "".join(f"<p>{escape(p)}</p>" for p in note.split("\n\n") if p.strip())
    return f"<h3>Пояснительная записка</h3>{note_html}"


# ── 4. Почасовая раскладка по подразделениям (24 ч) ───────────────────────────
def _hourly_blocks(m: dict) -> str:
    rows = m.get("hourly_unit", [])
    if not rows:
        return ""
    units: dict[str, list] = {}
    for r in rows:
        units.setdefault(r["unit"], []).append(r)
    out = ['<h3>3. Почасовая раскладка за сутки по подразделениям</h3>'
           '<div style="color:#57606a;font-size:13px;">Машины раздельно торф/песок; '
           'где машин нет — указана причина (простой из «Примечания»).</div>']
    for unit in sorted(units):
        head = _th("Час", "Торф", "Песок", "Причина простоя")
        body = ""
        for r in units[unit]:
            t, p = r.get("torf", 0), r.get("pesok", 0)
            reason = escape(r["reason"] or "")
            if t == 0 and p == 0 and reason:
                bg = "#fff8e6"
            elif t == 0 and p == 0:
                bg = "#fafafa"
            else:
                bg = "#ffffff"
            body += _row([r["hour"], _fmt_int(t) if t else "—",
                          _fmt_int(p) if p else "—", reason or ""], bg)
        out.append(f'<h4 style="margin:12px 0 2px;">{escape(unit)}</h4>'
                   f'<table style="border-collapse:collapse;margin:2px 0;font-size:13px;">{head}{body}</table>')
    return "".join(out)


# ── 4. Песок в разрезе промывочных приборов ───────────────────────────────────
def _pesok_devices_block(m: dict) -> str:
    pdv = m.get("pesok_devices", {})
    if not pdv:
        return ""
    out = ['<h3>4. Песок — по промывочным приборам (машины/час)</h3>']
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
        out.append(f'<h4 style="margin:12px 0 2px;">{escape(unit)}</h4>'
                   f'<table style="border-collapse:collapse;margin:2px 0;font-size:13px;">{head}{body}</table>')
    return "".join(out) if has_any else ""


# ── 5. Аналитика причин простоя ───────────────────────────────────────────────
def _idle_reasons_block(m: dict) -> str:
    rows = m.get("idle_reasons", [])
    if not rows:
        return '<h3>5. Аналитика причин простоя</h3><p>Простои за сутки не зафиксированы.</p>'
    planned = sum(r["hours"] for r in rows if r["kind"] == "плановый")
    unplanned = sum(r["hours"] for r in rows if r["kind"] == "внеплановый")
    head = _th("Подразделение", "Причина", "Тип", "Часов")
    body = ""
    for r in rows:
        bg = "#fff5f5" if r["kind"] == "внеплановый" else "#f6f8fa"
        body += _row([escape(r["unit"]), escape(r["reason"]), escape(r["kind"]),
                      _fmt_int(r["hours"])], bg)
    summary = (f'<p style="margin:6px 0;">Итого простоев: '
               f'<b>{_fmt_int(planned + unplanned)} ч</b> · '
               f'плановые (обед/пересменка/ЕТО) — {_fmt_int(planned)} ч · '
               f'<span style="color:#d1242f;">внеплановые (поломки/нет напряжения) — '
               f'{_fmt_int(unplanned)} ч</span>.</p>')
    return (f'<h3>5. Аналитика причин простоя</h3>{summary}'
            f'<table style="border-collapse:collapse;margin:8px 0;">{head}{body}</table>')


def build_html(m: dict, note: str, alerts: list[dict]) -> str:
    rd = escape(str(m.get("report_date", "")))
    per = m.get("period", ["", ""])
    return (
        '<!DOCTYPE html><html lang="ru"><head>'
        '<meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1">'
        '</head><body>'
        '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;color:#1f2328;max-width:860px;">'
        '<h2 style="margin:0 0 4px;">Хронометраж транспортировки торфов</h2>'
        f'<div style="color:#57606a;">Отчёт за прошедшие сутки: <b>{rd}</b> · '
        f'данные {escape(str(per[0]))}…{escape(str(per[1]))}</div>'
        + _kpi_by_unit(m)
        + _dev_block(m)
        + _note_block(note)
        + _hourly_blocks(m)
        + _pesok_devices_block(m)
        + _idle_reasons_block(m)
        + '<p style="color:#8c959f;font-size:12px;margin-top:16px;">'
          'Детализация — во вложении (Excel: реестр + листы аналитики, почасовка, причины простоя). '
          'Автоматическая рассылка хронометража.</p>'
        '</div></body></html>'
    )


def build_text(m: dict, note: str, alerts: list[dict]) -> str:
    lines = [f"Хронометраж торфов — отчёт за сутки {m.get('report_date', '')}", ""]
    lines.append("KPI по подразделениям (торф/песок, м³ | машин | простои ч):")
    for u in m.get("by_unit_day", []):
        lines.append(f"  {u['unit']}: торф {_fmt_int(u['vol_torf'])} / песок {_fmt_int(u['vol_pesok'])} м³ · "
                     f"машин {_fmt_int(u['mach_torf'])}/{_fmt_int(u['mach_pesok'])} · простои {u['idle_h']} ч")
    lines += ["", "Отклонения по подразделениям:"]
    for d in m.get("unit_dev", []):
        fl = _dev_flags(d)
        lines.append(f"  {d['unit']}: {'⚠ ' + '; '.join(fl) if fl else 'в норме'}")
    lines += ["", "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА:", note, "",
              "Причины простоя — см. HTML-письмо / Excel.",
              "Почасовая раскладка — см. HTML-письмо / Excel."]
    return "\n".join(lines)