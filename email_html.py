"""Рендер HTML- и текстового тела письма из metrics + записки + аномалий.
Только инлайн-стили (почтовые клиенты режут <style>). Зависимости — stdlib.
"""
from __future__ import annotations
from html import escape


def _fmt_int(x) -> str:
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except (ValueError, TypeError):
        return str(x)


def _kpi_tiles(m: dict) -> str:
    t = m["totals"]
    tiles = [("Объём, м³", _fmt_int(t["volume"])), ("Рейсов", _fmt_int(t["trips"])),
             ("м³/рейс", t.get("m3_per_trip")), ("Водителей", t["drivers"]),
             ("Самосвалов", t["trucks"]), ("Подразделений", t["units"])]
    cells = "".join(
        f'<td style="padding:8px 14px;border:1px solid #d0d7de;">'
        f'<div style="color:#57606a;font-size:12px;">{escape(str(k))}</div>'
        f'<div style="font-size:18px;font-weight:bold;">{escape(str(v))}</div></td>'
        for k, v in tiles)
    return f'<table style="border-collapse:collapse;margin:8px 0;"><tr>{cells}</tr></table>'


def _unit_table(m: dict) -> str:
    head = ("<tr>" + "".join(
        f'<th style="padding:6px 10px;background:#305496;color:#fff;'
        f'border:1px solid #d0d7de;text-align:left;">{h}</th>'
        for h in ("Подразделение", "Рейсы", "Объём, м³", "м³/рейс", "Доля, %")) + "</tr>")
    rows = ""
    for i, u in enumerate(m["by_unit"]):
        bg = "#f6f8fa" if i % 2 else "#ffffff"
        rows += ("<tr>" + "".join(
            f'<td style="padding:6px 10px;border:1px solid #d0d7de;background:{bg};">{c}</td>'
            for c in (escape(u["unit"]), _fmt_int(u["trips"]), _fmt_int(u["volume"]),
                      u["m3_per_trip"], u["share_pct"])) + "</tr>")
    return f'<table style="border-collapse:collapse;margin:8px 0;">{head}{rows}</table>'


def _trend_table(m: dict, days: int = 7) -> str:
    bd = m["by_date"][-days:]
    head = ('<tr><th style="padding:6px 10px;background:#305496;color:#fff;'
            'border:1px solid #d0d7de;">Дата</th>'
            '<th style="padding:6px 10px;background:#305496;color:#fff;'
            'border:1px solid #d0d7de;">Объём, м³</th>'
            '<th style="padding:6px 10px;background:#305496;color:#fff;'
            'border:1px solid #d0d7de;">Рейсы</th></tr>')
    rows = "".join(
        f'<tr><td style="padding:6px 10px;border:1px solid #d0d7de;">{escape(str(d["date"]))}</td>'
        f'<td style="padding:6px 10px;border:1px solid #d0d7de;">{_fmt_int(d["volume"])}</td>'
        f'<td style="padding:6px 10px;border:1px solid #d0d7de;">{_fmt_int(d["trips"])}</td></tr>'
        for d in bd)
    return f'<table style="border-collapse:collapse;margin:8px 0;">{head}{rows}</table>'


def _alerts_block(alerts: list[dict]) -> str:
    if not alerts:
        return '<p style="color:#1a7f37;">Существенных отклонений KPI не зафиксировано.</p>'
    items = "".join(f"<li>{escape(a['text'])}</li>" for a in alerts)
    return ('<div style="border-left:4px solid #d1242f;background:#fff8f8;padding:8px 14px;margin:8px 0;">'
            '<b style="color:#d1242f;">Внимание — отклонения KPI:</b>'
            f'<ul style="margin:6px 0;">{items}</ul></div>')


def build_html(m: dict, note: str, alerts: list[dict]) -> str:
    note_html = "".join(f"<p>{escape(p)}</p>" for p in note.split("\n\n") if p.strip())
    return (
        '<!DOCTYPE html><html lang="ru"><head>'
        '<meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1">'
        '</head><body>'
        '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;color:#1f2328;max-width:720px;">'
        '<h2 style="margin:0 0 4px;">Хронометраж транспортировки торфов</h2>'
        f'<div style="color:#57606a;">Отчётная дата: {escape(m["report_date"])} · '
        f'период {escape(str(m["period"][0]))}…{escape(str(m["period"][1]))}</div>'
        f'<h3>Ключевые показатели</h3>{_kpi_tiles(m)}'
        f'{_alerts_block(alerts)}'
        f'<h3>Пояснительная записка</h3>{note_html}'
        f'<h3>По подразделениям</h3>{_unit_table(m)}'
        f'<h3>Динамика (последние дни)</h3>{_trend_table(m)}'
        '<p style="color:#8c959f;font-size:12px;margin-top:16px;">'
        'Детализация — во вложении (Excel: реестр + листы аналитики). '
        'Автоматическая рассылка хронометража.</p>'
        '</div></body></html>'
    )


def build_text(m: dict, note: str, alerts: list[dict]) -> str:
    t = m["totals"]
    lines = [f"Хронометраж торфов — отчёт за {m['report_date']}",
             f"Период: {m['period'][0]}…{m['period'][1]}", "",
             f"Объём {_fmt_int(t['volume'])} м³ · рейсов {_fmt_int(t['trips'])} · "
             f"м³/рейс {t.get('m3_per_trip')} · водителей {t['drivers']} · "
             f"самосвалов {t['trucks']}", ""]
    if alerts:
        lines.append("ВНИМАНИЕ — отклонения KPI:")
        lines += [f"  - {a['text']}" for a in alerts]
    else:
        lines.append("Существенных отклонений KPI не зафиксировано.")
    lines += ["", "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА:", note, "", "По подразделениям:"]
    for u in m["by_unit"]:
        lines.append(f"  {u['unit']}: {_fmt_int(u['volume'])} м³ ({u['share_pct']}%), "
                     f"рейсов {_fmt_int(u['trips'])}")
    lines += ["", "Детализация — во вложении (Excel)."]
    return "\n".join(lines)
