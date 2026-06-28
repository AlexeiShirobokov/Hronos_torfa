"""Пояснительная записка (Фаза 1 — по правилам, без модели) и аномалии KPI.
Запуск: python3 explain.py — читает state/last_metrics.json, пишет записку и аномалии.
"""
from __future__ import annotations
import json, os, sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
STATE = BASE / "state"
OUT = BASE / "output"

MIN_BODY_UTIL = float(os.environ.get("HR_MIN_BODY_UTIL", "95"))
MAX_IDLE_H = float(os.environ.get("HR_MAX_IDLE_H", "1.0"))
MAX_VOLUME_DROP = float(os.environ.get("HR_MAX_VOLUME_DROP", "30"))


def _fmt_int(x) -> str:
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except (ValueError, TypeError):
        return str(x)


def build_brief(m: dict) -> str:
    t = m["totals"]
    lines = [
        f"Отчётная дата: {m['report_date']}. Период данных: {m['period'][0]}…{m['period'][1]}.",
        f"Итоги за период: объём {_fmt_int(t['volume'])} м³, рейсов {_fmt_int(t['trips'])}, "
        f"м³/рейс {t.get('m3_per_trip')}, водителей {t['drivers']}, самосвалов {t['trucks']}, "
        f"подразделений {t['units']}.",
        "По подразделениям (объём, доля%):",
    ]
    lines += [f"  - {u['unit']}: {_fmt_int(u['volume'])} м³ ({u['share_pct']}%), "
              f"м³/рейс {u['m3_per_trip']}" for u in m["by_unit"]]
    bd = m["by_date"]
    if len(bd) >= 2:
        lines.append(f"Динамика: {bd[-2]['date']} {_fmt_int(bd[-2]['volume'])} м³ → "
                     f"{bd[-1]['date']} {_fmt_int(bd[-1]['volume'])} м³.")
    lines.append(f"Загрузка кузова (общая): {m['truck_util'].get('overall_pct')}%.")
    if m["idle"]:
        lines.append("Простои (часов): " +
                     ", ".join(f"{i['unit']} {i['hours']}" for i in m["idle"]))
    lines.append(f"ABC водителей: A={m['abc']['A']}, B={m['abc']['B']}, C={m['abc']['C']}.")
    return "\n".join(lines)


def detect_anomalies(m: dict, thresholds: dict | None = None) -> list[dict]:
    th = {"util": MIN_BODY_UTIL, "idle": MAX_IDLE_H, "drop": MAX_VOLUME_DROP}
    if thresholds:
        th.update(thresholds)
    out: list[dict] = []
    util = m.get("truck_util", {}).get("overall_pct")
    if util is not None and util < th["util"]:
        out.append({"type": "body_util",
                    "text": f"Загрузка кузова {util}% ниже нормы {th['util']:.0f}%."})
    for i in m.get("idle", []):
        if i["hours"] > th["idle"]:
            out.append({"type": "idle",
                        "text": f"Простои в «{i['unit']}»: {i['hours']} ч "
                                f"(порог {th['idle']:.0f} ч/смену)."})
    # сравниваем ОТЧЁТНУЮ дату с днём перед ней (а не с частичным «сегодня»,
    # который мог попасть в by_date с утренним fetch)
    bd = m.get("by_date", [])
    rd = m.get("report_date")
    idx = next((i for i, d in enumerate(bd) if d.get("date") == rd), None)
    if idx is not None and idx >= 1 and bd[idx - 1]["volume"] > 0:
        prev, cur = bd[idx - 1], bd[idx]
        drop = (prev["volume"] - cur["volume"]) / prev["volume"] * 100
        if drop > th["drop"]:
            out.append({"type": "volume_drop",
                        "text": f"Объём упал на {drop:.0f}% к {prev['date']} "
                                f"({_fmt_int(prev['volume'])} → {_fmt_int(cur['volume'])} м³)."})
    return out


def rule_based_note(m: dict) -> str:
    t = m["totals"]
    units = sorted(m["by_unit"], key=lambda u: u["volume"], reverse=True)
    leader = units[0] if units else None
    al = detect_anomalies(m)
    parts = [
        f"# Пояснительная записка · {m['report_date']}",
        f"За период {m['period'][0]}…{m['period'][1]} перевезено "
        f"{_fmt_int(t['volume'])} м³ торфа ({_fmt_int(t['trips'])} рейсов, "
        f"в среднем {t.get('m3_per_trip')} м³/рейс). Задействовано {t['drivers']} водителей "
        f"и {t['trucks']} самосвалов в {t['units']} подразделениях.",
    ]
    if leader:
        parts.append(f"Наибольший вклад — «{leader['unit']}»: {_fmt_int(leader['volume'])} м³ "
                     f"({leader['share_pct']}% объёма).")
    bd = m["by_date"]
    if len(bd) >= 2:
        d = "снизился" if bd[-1]["volume"] < bd[-2]["volume"] else "вырос"
        parts.append(f"К предыдущему дню объём {d}: "
                     f"{_fmt_int(bd[-2]['volume'])} → {_fmt_int(bd[-1]['volume'])} м³.")
    if al:
        parts.append("Внимание: " + " ".join(a["text"] for a in al))
    else:
        parts.append("Существенных отклонений KPI не зафиксировано.")
    return "\n\n".join(parts)


def generate_note(m: dict) -> str:
    # Фаза 1: записка собирается по правилам (без модели/моста).
    # Точка расширения: здесь можно подключить llm-мост (Task 3) в следующей фазе.
    return rule_based_note(m)


def main() -> int:
    mp = STATE / "last_metrics.json"
    if not mp.exists():
        print(f"[ERR] нет {mp} (запусти build_xlsx/analytics)"); return 2
    m = json.loads(mp.read_text(encoding="utf-8"))
    note = generate_note(m)
    alerts = detect_anomalies(m)

    OUT.mkdir(parents=True, exist_ok=True); STATE.mkdir(parents=True, exist_ok=True)
    (OUT / f"Пояснительная_записка_{m['report_date']}.md").write_text(note, encoding="utf-8")
    (STATE / "last_note.txt").write_text(note, encoding="utf-8")
    (STATE / "last_alerts.json").write_text(
        json.dumps(alerts, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] записка ({len(note)} симв.), аномалий: {len(alerts)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
