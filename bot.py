"""Telegram-бот @alexeids_bot для управления пайплайном хронометража.

Команды:
  /start           — регистрация чата как админ (только первый раз / по списку BOT_ADMINS)
  /run_now         — запустить пайплайн (fetch → consolidate → pdf) и прислать уведомление
  /status          — последний запуск (state/last_run.json) и хвост логов
  /list_emails     — показать recipients.txt
  /add_email a@b   — добавить адрес в recipients.txt
  /remove_email a@b — удалить адрес
  /pdf, /xlsx      — прислать последние файлы
  /send_now        — выполнить рассылку (после подтверждения)
  /confirm         — то же, что нажатие кнопки «Отправить»
  /cancel          — снять флаг ожидания подтверждения

Зависимости: только стандартная библиотека.
Polling Telegram Bot API (long polling).
"""
# DEPRECATED (Фаза 1): Telegram убран из пайплайна — доставка и алерты идут по email.
# launchd-задача бота снята. Оставлено для истории; удалить после Фазы 2.
from __future__ import annotations
import json, os, subprocess, sys, time, urllib.parse, urllib.request
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
LOGS = BASE / "logs"
STATE = BASE / "state"
RCPT = BASE / "recipients.txt"
PY = sys.executable


def load_env() -> dict:
    p = BASE / ".env"
    out = {}
    if not p.exists(): return out
    for raw in p.read_text(encoding="utf-8").splitlines():
        s = raw.strip()
        if not s or s.startswith("#") or "=" not in s: continue
        k, v = s.split("=", 1)
        out[k.strip()] = v.strip().strip('"').strip("'")
    return out


ENV = load_env()
TOKEN = ENV.get("BOT_TOKEN") or ENV.get("api_token")
ADMINS = {x.strip() for x in (ENV.get("BOT_ADMINS", "")).split(",") if x.strip()}

if not TOKEN:
    sys.stderr.write("[FATAL] нет BOT_TOKEN / api_token в .env\n"); sys.exit(2)

API = f"https://api.telegram.org/bot{TOKEN}"


def log(msg: str) -> None:
    LOGS.mkdir(parents=True, exist_ok=True)
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    with (LOGS / "bot.log").open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def api(method: str, **params) -> dict:
    url = f"{API}/{method}"
    fields = {k: ("" if v is None else v) for k, v in params.items()}
    data = urllib.parse.urlencode(fields).encode()
    req = urllib.request.Request(url, data=data)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode())
    except Exception as e:
        log(f"[ERR] api({method}): {e!r}")
        return {"ok": False, "description": str(e)}


def api_send_document(chat_id: str | int, file_path: Path, caption: str = "") -> dict:
    """multipart upload через стандартную библиотеку."""
    boundary = "----formdata-cw-" + datetime.now().strftime("%s%f")
    body = []
    def add_field(name, value):
        body.append(f"--{boundary}\r\n".encode())
        body.append(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode())
        body.append(str(value).encode() + b"\r\n")
    def add_file(name, fname, payload, ctype):
        body.append(f"--{boundary}\r\n".encode())
        body.append(f'Content-Disposition: form-data; name="{name}"; filename="{fname}"\r\n'.encode())
        body.append(f"Content-Type: {ctype}\r\n\r\n".encode())
        body.append(payload)
        body.append(b"\r\n")
    add_field("chat_id", chat_id)
    if caption: add_field("caption", caption)
    ctype = "application/pdf" if file_path.suffix.lower() == ".pdf" else "application/octet-stream"
    add_file("document", file_path.name, file_path.read_bytes(), ctype)
    body.append(f"--{boundary}--\r\n".encode())
    payload = b"".join(body)
    req = urllib.request.Request(f"{API}/sendDocument", data=payload,
                                 headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            return json.loads(r.read().decode())
    except Exception as e:
        log(f"[ERR] sendDocument: {e!r}")
        return {"ok": False}


def admin_chat() -> str | None:
    p = STATE / "admin_chat_id.txt"
    return p.read_text(encoding="utf-8").strip() if p.exists() else None


def is_admin(user_id, chat_id) -> bool:
    if ADMINS:
        return str(user_id) in ADMINS
    a = admin_chat()
    return a is None or str(chat_id) == a


def find_latest():
    xlsx = sorted((BASE / "output").glob("Хронометраж_транспортировки_торфов_*.xlsx"),
                  key=lambda p: p.stat().st_mtime, reverse=True)
    pdf = sorted((BASE / "output").glob("Аналитика_хронометраж_торфов_*.pdf"),
                 key=lambda p: p.stat().st_mtime, reverse=True)
    return (xlsx[0] if xlsx else None, pdf[0] if pdf else None)


def cmd_start(chat_id, user_id, args):
    STATE.mkdir(parents=True, exist_ok=True)
    p = STATE / "admin_chat_id.txt"
    if not p.exists():
        p.write_text(str(chat_id), encoding="utf-8")
        api("sendMessage", chat_id=chat_id,
            text=f"Привет! Чат {chat_id} зарегистрирован как админ.\n"
                 f"Команды: /run_now, /status, /list_emails, /add_email, /remove_email, "
                 f"/pdf, /xlsx, /send_now, /confirm, /cancel")
    else:
        api("sendMessage", chat_id=chat_id,
            text=f"Бот уже привязан к чату {p.read_text(encoding='utf-8').strip()}.\n"
                 f"Команды: /run_now /status /list_emails /add_email /remove_email /pdf /xlsx /send_now /confirm /cancel")


def cmd_run_now(chat_id, user_id, args):
    api("sendMessage", chat_id=chat_id, text="🔄 Запускаю пайплайн… (fetch → consolidate → pdf)")
    p = subprocess.Popen([PY, str(BASE / "run_daily.py")], cwd=str(BASE),
                         stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    out, _ = p.communicate(timeout=600)
    tail = out.decode("utf-8", errors="replace")[-1500:]
    api("sendMessage", chat_id=chat_id,
        text=f"Готово (exit={p.returncode}).\n<pre>{tail}</pre>", parse_mode="HTML")


def cmd_status(chat_id, user_id, args):
    p = STATE / "last_run.json"
    if not p.exists():
        api("sendMessage", chat_id=chat_id, text="Нет данных о запусках.")
        return
    data = json.loads(p.read_text(encoding="utf-8"))
    txt = ("<b>Последний запуск</b>\n"
           f"Когда: {data.get('ts')}\n"
           f"Отчётная дата: {data.get('report_date')}\n"
           f"Длительность: {data.get('duration_sec')} сек\n"
           f"OK: {data.get('ok')}\n"
           f"Ожидает подтверждения отправки: {data.get('pending_send')}\n"
           f"XLSX: {Path(data['xlsx']).name if data.get('xlsx') else '—'}\n"
           f"PDF: {Path(data['pdf']).name if data.get('pdf') else '—'}")
    # хвост лога
    rl = LOGS / "run_daily.log"
    if rl.exists():
        lines = rl.read_text(encoding="utf-8", errors="replace").splitlines()[-15:]
        txt += "\n\n<b>Лог (хвост)</b>\n<pre>" + "\n".join(lines).replace("<", "&lt;") + "</pre>"
    api("sendMessage", chat_id=chat_id, text=txt, parse_mode="HTML")


def cmd_list_emails(chat_id, user_id, args):
    if not RCPT.exists():
        api("sendMessage", chat_id=chat_id, text="recipients.txt отсутствует.")
        return
    lines = [l.strip() for l in RCPT.read_text(encoding="utf-8").splitlines()
             if l.strip() and not l.startswith("#")]
    api("sendMessage", chat_id=chat_id,
        text="📧 Получатели:\n" + ("\n".join(f"• {e}" for e in lines) or "(пусто)"))


def cmd_add_email(chat_id, user_id, args):
    if not args or "@" not in args[0]:
        api("sendMessage", chat_id=chat_id, text="Используй: /add_email name@domain")
        return
    e = args[0].strip()
    cur = [l.strip() for l in RCPT.read_text(encoding="utf-8").splitlines()] if RCPT.exists() else []
    if e in cur:
        api("sendMessage", chat_id=chat_id, text=f"{e} уже в списке."); return
    cur.append(e)
    RCPT.write_text("\n".join(cur) + "\n", encoding="utf-8")
    api("sendMessage", chat_id=chat_id, text=f"✅ Добавлен {e}.")


def cmd_remove_email(chat_id, user_id, args):
    if not args:
        api("sendMessage", chat_id=chat_id, text="Используй: /remove_email name@domain"); return
    e = args[0].strip().lower()
    cur = [l.strip() for l in RCPT.read_text(encoding="utf-8").splitlines()] if RCPT.exists() else []
    new = [x for x in cur if x.strip().lower() != e]
    RCPT.write_text("\n".join(new) + "\n", encoding="utf-8")
    api("sendMessage", chat_id=chat_id, text=f"🗑 Удалено: {e}.")


def cmd_send_file(chat_id, user_id, args, which: str):
    xlsx, pdf = find_latest()
    f = pdf if which == "pdf" else xlsx
    if not f or not f.exists():
        api("sendMessage", chat_id=chat_id, text=f"Файл {which} не найден."); return
    api_send_document(chat_id, f, caption=f.name)


def cmd_send_now(chat_id, user_id, args):
    p = subprocess.run([PY, str(BASE / "send_email.py")], cwd=str(BASE),
                       capture_output=True, text=True, timeout=180)
    out = (p.stdout + p.stderr).strip()[-1500:]
    if p.returncode == 0:
        api("sendMessage", chat_id=chat_id, text=f"📨 Отправлено.\n<pre>{out}</pre>", parse_mode="HTML")
        # сбросить pending
        sp = STATE / "last_run.json"
        if sp.exists():
            d = json.loads(sp.read_text(encoding="utf-8"))
            d["pending_send"] = False; d["last_send_ts"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            sp.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        api("sendMessage", chat_id=chat_id,
            text=f"❌ Ошибка отправки (exit={p.returncode}):\n<pre>{out}</pre>", parse_mode="HTML")


def cmd_cancel(chat_id, user_id, args):
    sp = STATE / "last_run.json"
    if sp.exists():
        d = json.loads(sp.read_text(encoding="utf-8")); d["pending_send"] = False
        sp.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    api("sendMessage", chat_id=chat_id, text="✖️ Отправка отменена.")


COMMANDS = {
    "start": cmd_start,
    "run_now": cmd_run_now,
    "status": cmd_status,
    "list_emails": cmd_list_emails,
    "add_email": cmd_add_email,
    "remove_email": cmd_remove_email,
    "pdf": lambda c,u,a: cmd_send_file(c,u,a,"pdf"),
    "xlsx": lambda c,u,a: cmd_send_file(c,u,a,"xlsx"),
    "send_now": cmd_send_now,
    "confirm": cmd_send_now,
    "cancel": cmd_cancel,
}


def handle_message(msg: dict) -> None:
    chat_id = msg["chat"]["id"]
    user_id = msg["from"]["id"]
    text = (msg.get("text") or "").strip()
    if not text.startswith("/"): return
    parts = text.split()
    cmd = parts[0][1:].lower()
    # отбросим @username
    if "@" in cmd: cmd = cmd.split("@", 1)[0]
    args = parts[1:]
    if not is_admin(user_id, chat_id):
        api("sendMessage", chat_id=chat_id, text="⛔ Доступ ограничен.")
        return
    fn = COMMANDS.get(cmd)
    if not fn:
        api("sendMessage", chat_id=chat_id,
            text="Неизвестная команда. /start /run_now /status /list_emails /add_email /remove_email /pdf /xlsx /send_now /cancel")
        return
    try:
        fn(chat_id, user_id, args)
    except Exception as e:
        log(f"[ERR] {cmd}: {e!r}")
        api("sendMessage", chat_id=chat_id, text=f"⚠️ Ошибка: {e!r}")


def handle_callback(cq: dict) -> None:
    chat_id = cq["message"]["chat"]["id"]
    user_id = cq["from"]["id"]
    data = cq.get("data", "")
    api("answerCallbackQuery", callback_query_id=cq["id"])
    if not is_admin(user_id, chat_id):
        api("sendMessage", chat_id=chat_id, text="⛔ Доступ ограничен."); return
    if data == "send_now":
        cmd_send_now(chat_id, user_id, [])
    elif data == "cancel":
        cmd_cancel(chat_id, user_id, [])


def main() -> int:
    log("Bot started")
    offset = 0
    while True:
        try:
            r = api("getUpdates", offset=offset, timeout=25)
            if not r.get("ok"):
                time.sleep(3); continue
            for upd in r.get("result", []):
                offset = upd["update_id"] + 1
                if "message" in upd:
                    handle_message(upd["message"])
                elif "callback_query" in upd:
                    handle_callback(upd["callback_query"])
        except KeyboardInterrupt:
            log("Bot stopped"); return 0
        except Exception as e:
            log(f"[ERR] main loop: {e!r}"); time.sleep(5)


if __name__ == "__main__":
    sys.exit(main())
