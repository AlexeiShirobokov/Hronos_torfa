"""Обратная связь диспетчерам: запрос причины по приборо-часам без причины.

Для отчётного дня находит приборо-часы, где выработка промприбора ниже нормы
более чем на 20% ИЛИ ноль, и при этом в «Примечании» (строки песка + «Простой»)
нет причины. По каждому подразделению шлёт его диспетчеру одно сводное письмо
с просьбой указать причину и внести её в Хронометраж.

Правила (согласовано):
- флагаем «красные + нули без причины»;
- напоминаем каждый прогон, пока причина не внесена (без антиспама);
- окно 09:00–19:00 Asia/Ust-Nera (вне окна — пропуск, если не --force).

Запуск (боевой, из cron в :35):
    TZ=Asia/Ust-Nera python dispatcher_feedback.py
Тест на себя по одному подразделению:
    python dispatcher_feedback.py --only-unit Дражный --to shirobokov@pskgold.ru --force
"""
from __future__ import annotations

import argparse
import re
import ssl
import smtplib
from datetime import datetime, time, timedelta
from email.message import EmailMessage
from pathlib import Path

import pandas as pd

from peski_email_body import (
    _pribor_label, _pribor_norm, _mark_of, _day_slice, _fmt_int,
    SAND_PEREDELS, IDLE_PEREDEL, DEVIATION_THRESHOLD,
    COL_UNIT, COL_FACT, COL_TIME, COL_PER, COL_VOL, COL_NOTE,
)
from send_email import load_env

BASE = Path(__file__).resolve().parent
OUT = BASE / "output"
PARALLEL_OUT = BASE / "parallel_hronos" / "output"
LOGS = BASE / "logs"
STATE = BASE / "state"
TODAY_REPORT_CUTOFF = time(10, 0)          # как в run_hronos_peski_autosend
WINDOW_START_HOUR, WINDOW_END_HOUR = 9, 19   # окно 09:xx–19:xx Asia/Ust-Nera
SUBJECT = "Хронометраж: прошу указать причину простоя"

DISPATCHERS = {
    "Дражный": "drazhny@pskgold.ru",
    "Сайлык": "davidov@pskgold.ru",
    "Талынья": "smorodina@pskgold.ru",
    "Эрел": "haptagay-haya@pskgold.ru",
    "Обман": "obman@pskgold.ru",
}


def log(message: str) -> None:
    LOGS.mkdir(parents=True, exist_ok=True)
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}"
    print(line)
    with (LOGS / "dispatcher_feedback.log").open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def report_date_for_run(now: datetime | None = None) -> str:
    now = now or datetime.now()
    d = now.date() if now.time() >= TODAY_REPORT_CUTOFF else now.date() - timedelta(days=1)
    return d.isoformat()


def latest_book() -> Path:
    finals = sorted(PARALLEL_OUT.glob("*_final.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if finals:
        return finals[0]
    target = OUT / "Хронометраж транспортировки торфов и песков.xlsx"
    if target.exists():
        return target
    raise RuntimeError("не найден финальный xlsx")


def _inv_of(prib: str) -> str:
    parts = str(prib).split(" / ", 1)
    return parts[1].strip() if len(parts) > 1 else ""


def flagged_by_unit(df: pd.DataFrame, report_date: str) -> dict[str, list[dict]]:
    """{подразделение: [ {прибор, время, факт, vol, норма}, ... ]}.

    Флагаем приборо-час, если прибор появлялся в наряде и в этот час его нет/0
    (ниже нормы >20%), причина не указана, и он не работает в другом подразделении
    (перестановка). Часы — от первого появления прибора до последнего отчётного
    часа подразделения.
    """
    day = _day_slice(df, report_date)
    d = day.copy()
    d[COL_UNIT] = d[COL_UNIT].astype(str).str.strip()
    d["_fact"] = pd.to_datetime(d[COL_FACT], errors="coerce")
    d["_t"] = d[COL_TIME].astype(str).str.strip()
    d["_prib"] = _pribor_label(d)
    d = d.dropna(subset=["_fact"])
    d = d[d["_t"].str.match(r"^\d{1,2}:\d{2}$")]
    d["_slot"] = list(zip(d["_fact"], d["_t"]))

    sand = d[d[COL_PER].astype(str).str.strip().isin(SAND_PEREDELS)].copy()
    vol = sand.groupby([COL_UNIT, "_prib", "_slot"])[COL_VOL].sum()

    # Причина «есть» — любое непустое Примечание (любой передел, вкл. «Обед»).
    notes = d.copy()
    notes["_r"] = notes[COL_NOTE].astype(str).str.strip().replace({"nan": "", "None": ""})
    notes = notes[notes["_r"] != ""]
    have_reason = set(zip(notes[COL_UNIT], notes["_prib"], notes["_slot"]))

    # Перестановка: где прибор (по инв.№) реально выдаёт >0 в этот слот.
    pos = sand[sand[COL_VOL] > 0].copy()
    pos["_inv"] = pos["_prib"].map(_inv_of)
    working: dict[tuple, set] = {}
    for inv, slot, unit in zip(pos["_inv"], pos["_slot"], pos[COL_UNIT]):
        working.setdefault((inv, slot), set()).add(unit)

    slots_by_unit = {u: sorted(set(sub["_slot"])) for u, sub in d.groupby(COL_UNIT)}
    first_slot = sand.groupby([COL_UNIT, "_prib"])["_slot"].min()

    out: dict[str, list[dict]] = {}
    for (unit, prib), fslot in first_slot.items():
        if not str(prib).replace("/", "").strip():
            continue
        norm = _pribor_norm(_mark_of(prib))
        if norm is None:
            continue
        inv = _inv_of(prib)
        for slot in slots_by_unit.get(unit, []):
            if slot < fslot:                               # до первого появления прибора
                continue
            v = float(vol.get((unit, prib, slot), 0.0))    # нет строки → 0
            if v >= norm * (1 - DEVIATION_THRESHOLD):      # норма выполнена
                continue
            if (unit, prib, slot) in have_reason:          # причина указана
                continue
            if inv and (working.get((inv, slot), set()) - {unit}):   # работает в др. подразделении
                continue
            fact, t = slot
            out.setdefault(unit, []).append(
                {"prib": prib, "time": t, "fact": pd.Timestamp(fact), "vol": v, "norm": norm})
    for u in out:
        out[u].sort(key=lambda r: (r["fact"], r["time"]))
    return out


def build_text(unit: str, items: list[dict]) -> str:
    lines = [
        "Здравствуйте!",
        "",
        f"По подразделению «{unit}» за отчётный день выявлены часы работы промприборов "
        "ниже нормы без указанной причины. Прошу пояснить причину простоя и внести её "
        "в Хронометраж:",
        "",
    ]
    for r in items:
        d = r["fact"].strftime("%d.%m.%Y")
        lines.append(f"  • прибор {r['prib']} — {r['time']} {d} "
                     f"(факт {_fmt_int(r['vol'])} м³ при норме {_fmt_int(r['norm'])} м³/ч)")
    lines += ["", "Спасибо."]
    return "\n".join(lines)


def send_plain(env: dict, to_addr: str, subject: str, text: str) -> None:
    msg = EmailMessage()
    msg["From"] = env["YANDEX_LOGIN"]
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(text)
    host = env.get("SMTP_HOST", "smtp.yandex.ru")
    port = int(env.get("SMTP_PORT", "465"))
    login = env.get("SMTP_LOGIN") or env["YANDEX_LOGIN"]
    password = env.get("SMTP_PASSWORD") or env["YANDEX_APP_PASSWORD"]
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
        s.login(login, password)
        s.send_message(msg)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Запросы диспетчерам о причинах простоя")
    ap.add_argument("--xlsx", type=Path, help="книга (по умолчанию — последний final)")
    ap.add_argument("--date", help="Дата выдачи наряд-задания YYYY-MM-DD (по умолчанию — авто)")
    ap.add_argument("--only-unit", help="только это подразделение")
    ap.add_argument("--to", help="переопределить получателя (тест на себя)")
    ap.add_argument("--dry-run", action="store_true", help="не отправлять, только показать")
    ap.add_argument("--force", action="store_true", help="игнорировать окно 09:00–19:00")
    args = ap.parse_args(argv)

    now = datetime.now()
    if not args.force and not (WINDOW_START_HOUR <= now.hour <= WINDOW_END_HOUR):
        log(f"вне окна 09:00–19:59 ({now:%H:%M}) — пропуск")
        return 0

    report_date = args.date or report_date_for_run(now)
    xlsx = args.xlsx or latest_book()
    df = pd.read_excel(xlsx, sheet_name="Сводный_Реестр")
    flagged = flagged_by_unit(df, report_date)
    if args.only_unit:
        flagged = {u: v for u, v in flagged.items() if u == args.only_unit}

    log(f"дата {report_date}, книга {Path(xlsx).name}: подразделений с флагами {len(flagged)}")
    if not flagged:
        log("нет приборо-часов без причины — писем нет")
        return 0

    env = None if args.dry_run else load_env()
    sent = 0
    for unit, items in flagged.items():
        to_addr = args.to or DISPATCHERS.get(unit)
        if not to_addr:
            log(f"[WARN] нет email диспетчера для «{unit}» — пропуск ({len(items)} шт.)")
            continue
        text = build_text(unit, items)
        if args.dry_run:
            log(f"[DRY] «{unit}» → {to_addr} ({len(items)} шт.)\n{text}")
            continue
        try:
            send_plain(env, to_addr, SUBJECT, text)
            sent += 1
            log(f"[OK] «{unit}» → {to_addr}: {len(items)} приборо-часов")
        except Exception as exc:
            log(f"[ERR] «{unit}» → {to_addr}: {type(exc).__name__}: {exc}")
    log(f"DONE отправлено писем: {sent}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
