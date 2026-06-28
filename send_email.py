"""SMTP-отправка консолидированных файлов получателям из recipients.txt.

Использование:
    python3 send_email.py             # отправить последние output-файлы
    python3 send_email.py --check     # только проверка SMTP-логина

Требования: только стандартная библиотека.
"""
from __future__ import annotations
import argparse
import mimetypes
import os
import smtplib
import ssl
import sys
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

BASE = Path(__file__).resolve().parent
ENV = BASE / ".env"
OUT = BASE / "output"
LOGS = BASE / "logs"
RCPT_FILE = BASE / "recipients.txt"


def load_env() -> dict:
    data = {}
    if not ENV.exists():
        raise SystemExit(f"[ERR] .env не найден: {ENV}")
    for raw in ENV.read_text(encoding="utf-8").splitlines():
        s = raw.strip()
        if not s or s.startswith("#") or "=" not in s:
            continue
        k, v = s.split("=", 1)
        data[k.strip()] = v.strip().strip('"').strip("'")
    return data


def load_recipients() -> list[str]:
    if not RCPT_FILE.exists():
        return []
    out = []
    for raw in RCPT_FILE.read_text(encoding="utf-8").splitlines():
        s = raw.strip()
        if not s or s.startswith("#"):
            continue
        if "@" not in s:
            continue
        out.append(s)
    # dedup, сохраняя порядок
    seen = set(); ord_ = []
    for e in out:
        if e.lower() in seen: continue
        seen.add(e.lower()); ord_.append(e)
    return ord_


def find_latest_xlsx() -> Path | None:
    """Найти самую свежую книгу реестра в /output."""
    xlsx = sorted(OUT.glob("Хронометраж_транспортировки_торфов_*.xlsx"),
                  key=lambda p: p.stat().st_mtime, reverse=True)
    return xlsx[0] if xlsx else None


def attach_file(msg: EmailMessage, path: Path) -> None:
    ctype, enc = mimetypes.guess_type(path.name)
    if ctype is None or enc is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    msg.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype,
                       filename=path.name)


def build_message(env: dict, recipients: list[str], xlsx: Path | None,
                  html_body: str, text_body: str,
                  report_date: str | None = None) -> EmailMessage:
    sender = env["YANDEX_LOGIN"]
    msg = EmailMessage()
    rd = report_date or datetime.now().strftime("%Y-%m-%d")
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"Хронометраж торфов — отчёт за {rd}"
    msg.set_content(text_body)                       # text/plain (фолбэк)
    msg.add_alternative(html_body, subtype="html")   # text/html
    if xlsx and xlsx.exists():
        attach_file(msg, xlsx)
    return msg


def send(env: dict, msg: EmailMessage) -> None:
    host = env.get("SMTP_HOST", "smtp.yandex.ru")
    port = int(env.get("SMTP_PORT", "465"))
    login = env.get("SMTP_LOGIN") or env["YANDEX_LOGIN"]
    password = env.get("SMTP_PASSWORD") or env["YANDEX_APP_PASSWORD"]
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
        s.login(login, password)
        s.send_message(msg)


def log(msg: str) -> None:
    LOGS.mkdir(parents=True, exist_ok=True)
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line)
    with (LOGS / "mail.log").open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def main() -> int:
    import json
    import email_html
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="только проверить SMTP-логин")
    ap.add_argument("--xlsx", type=Path, help="путь к xlsx (по умолчанию — последний)")
    args = ap.parse_args()

    env = load_env()
    if "@" not in env.get("YANDEX_LOGIN", ""):
        log(f"[ERR] YANDEX_LOGIN некорректный: {env.get('YANDEX_LOGIN')!r}")
        return 2

    if args.check:
        host = env.get("SMTP_HOST", "smtp.yandex.ru")
        port = int(env.get("SMTP_PORT", "465"))
        try:
            with smtplib.SMTP_SSL(host, port, timeout=20) as s:
                s.login(env["YANDEX_LOGIN"], env["YANDEX_APP_PASSWORD"])
            log("[OK] SMTP-логин принят")
            return 0
        except Exception as e:
            log(f"[ERR] SMTP-проверка не удалась: {e!r}")
            return 3

    rcpts = load_recipients()
    if not rcpts:
        log(f"[ERR] список получателей пуст: {RCPT_FILE}")
        return 4

    xlsx = args.xlsx or find_latest_xlsx()
    if not xlsx:
        log(f"[ERR] не найден xlsx в {OUT}")
        return 5

    state = BASE / "state"
    metrics = json.loads((state / "last_metrics.json").read_text(encoding="utf-8"))
    note_p = state / "last_note.txt"
    note = note_p.read_text(encoding="utf-8") if note_p.exists() else ""
    alerts_p = state / "last_alerts.json"
    alerts = json.loads(alerts_p.read_text(encoding="utf-8")) if alerts_p.exists() else []
    html = email_html.build_html(metrics, note, alerts)
    text = email_html.build_text(metrics, note, alerts)

    msg = build_message(env, rcpts, xlsx, html, text, metrics.get("report_date"))
    try:
        send(env, msg)
    except Exception as e:
        log(f"[ERR] отправка не удалась: {e!r}")
        return 6
    log(f"[OK] отправлено {len(rcpts)} получателям: {', '.join(rcpts)}  |  {xlsx.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
