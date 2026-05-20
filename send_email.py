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


def find_latest_outputs() -> tuple[Path | None, Path | None]:
    """Найти самые свежие xlsx и pdf в /output."""
    xlsx = sorted(OUT.glob("Хронометраж_транспортировки_торфов_*.xlsx"),
                  key=lambda p: p.stat().st_mtime, reverse=True)
    pdf = sorted(OUT.glob("Аналитика_хронометраж_торфов_*.pdf"),
                 key=lambda p: p.stat().st_mtime, reverse=True)
    return (xlsx[0] if xlsx else None, pdf[0] if pdf else None)


def attach_file(msg: EmailMessage, path: Path) -> None:
    ctype, enc = mimetypes.guess_type(path.name)
    if ctype is None or enc is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    msg.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype,
                       filename=path.name)


def build_message(env: dict, recipients: list[str],
                  xlsx: Path | None, pdf: Path | None,
                  body_extra: str = "") -> EmailMessage:
    sender = env["YANDEX_LOGIN"]
    msg = EmailMessage()
    today = datetime.now().strftime("%Y-%m-%d")
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"Хронометраж торфов — отчёт за {today}"
    body = (
        f"Автоматическая рассылка по результатам сборки реестра и аналитики.\n"
        f"Сформировано: {datetime.now():%Y-%m-%d %H:%M}\n\n"
        f"Во вложении:\n"
        f"  • Консолидированный реестр (xlsx)\n"
        f"  • Аналитика по подразделениям, водителям, грузоподъёмности (pdf)\n"
    )
    if body_extra:
        body += "\n" + body_extra + "\n"
    body += "\n— @alexeids_bot"
    msg.set_content(body)
    if xlsx and xlsx.exists(): attach_file(msg, xlsx)
    if pdf and pdf.exists():   attach_file(msg, pdf)
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
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true",
                    help="только проверить SMTP-логин")
    ap.add_argument("--xlsx", type=Path, help="путь к xlsx (по умолчанию — последний)")
    ap.add_argument("--pdf", type=Path, help="путь к pdf (по умолчанию — последний)")
    ap.add_argument("--note", default="", help="дополнительный текст в теле письма")
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

    xlsx = args.xlsx or None
    pdf = args.pdf or None
    if not xlsx or not pdf:
        a, b = find_latest_outputs()
        xlsx = xlsx or a
        pdf = pdf or b
    if not xlsx and not pdf:
        log(f"[ERR] не найдены файлы в {OUT}")
        return 5

    msg = build_message(env, rcpts, xlsx, pdf, args.note)
    try:
        send(env, msg)
    except Exception as e:
        log(f"[ERR] отправка не удалась: {e!r}")
        return 6
    sizes = []
    if xlsx and xlsx.exists(): sizes.append(f"{xlsx.name}={xlsx.stat().st_size}b")
    if pdf and pdf.exists():   sizes.append(f"{pdf.name}={pdf.stat().st_size}b")
    log(f"[OK] отправлено {len(rcpts)} получателям: {', '.join(rcpts)}  |  {' ; '.join(sizes)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
