"""Оркестратор: fetch → consolidate → build_pdf → уведомить бота (или отправить).

Запускается launchd-ом каждый час в xx:10 с 9 до 19.
Поведение по умолчанию: после сборки PDF/XLSX отправляет уведомление в бот
с кнопкой «Подтвердить рассылку». Авто-отправка писем выполняется ТОЛЬКО
после нажатия кнопки или через /confirm в боте.
"""
from __future__ import annotations
import json, os, subprocess, sys, time, urllib.parse, urllib.request
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
LOGS = BASE / "logs"
STATE = BASE / "state"
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


def log(msg: str) -> None:
    LOGS.mkdir(parents=True, exist_ok=True)
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line)
    with (LOGS / "run_daily.log").open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def run(script: str) -> tuple[int, str]:
    log(f"→ {script}")
    p = subprocess.run([PY, str(BASE / script)], cwd=str(BASE),
                       capture_output=True, text=True)
    if p.stdout: log(p.stdout.strip()[:2000])
    if p.stderr: log("STDERR: " + p.stderr.strip()[:2000])
    log(f"   exit={p.returncode}")
    return p.returncode, p.stdout


def tg_request(token: str, method: str, **params) -> dict:
    url = f"https://api.telegram.org/bot{token}/{method}"
    data = urllib.parse.urlencode({k: v for k, v in params.items() if v is not None}).encode()
    req = urllib.request.Request(url, data=data)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


def notify_bot(env: dict, report_date: str, xlsx: Path | None, pdf: Path | None) -> None:
    token = env.get("BOT_TOKEN") or env.get("api_token")
    chat_id = env.get("BOT_ADMIN_CHAT_ID") or _read_admin_chat()
    if not token or not chat_id:
        log("[WARN] нет BOT_TOKEN или BOT_ADMIN_CHAT_ID — уведомление не отправляется")
        return
    text = (f"📦 Отчёт за <b>{report_date}</b> готов.\n"
            f"• XLSX: {xlsx.name if xlsx else '—'}\n"
            f"• PDF: {pdf.name if pdf else '—'}\n\n"
            f"Подтвердить рассылку по recipients.txt?")
    kb = {
        "inline_keyboard": [[
            {"text": "✅ Отправить", "callback_data": "send_now"},
            {"text": "✖️ Отмена",    "callback_data": "cancel"},
        ]]
    }
    try:
        r = tg_request(token, "sendMessage",
                       chat_id=chat_id, text=text, parse_mode="HTML",
                       reply_markup=json.dumps(kb))
        log(f"[INFO] уведомление отправлено (msg id={r.get('result',{}).get('message_id')})")
    except Exception as e:
        log(f"[ERR] не удалось уведомить бот: {e!r}")


def _read_admin_chat() -> str | None:
    p = STATE / "admin_chat_id.txt"
    return p.read_text(encoding="utf-8").strip() if p.exists() else None


def find_latest():
    xlsx = sorted((BASE / "output").glob("Хронометраж_транспортировки_торфов_*.xlsx"),
                  key=lambda p: p.stat().st_mtime, reverse=True)
    pdf = sorted((BASE / "output").glob("Аналитика_хронометраж_торфов_*.pdf"),
                 key=lambda p: p.stat().st_mtime, reverse=True)
    return (xlsx[0] if xlsx else None, pdf[0] if pdf else None)


def main() -> int:
    STATE.mkdir(parents=True, exist_ok=True)
    LOGS.mkdir(parents=True, exist_ok=True)
    env = load_env()
    t0 = time.time()
    log("="*60)
    log("RUN_DAILY START")

    code = 0
    rc, _ = run("fetch_kronos_torf.py")
    if rc != 0:
        log("[WARN] fetch завершился с ошибкой, продолжаю с локальными вложениями")
        code = rc
    rc, _ = run("consolidate.py")
    if rc != 0:
        log(f"[ERR] consolidate завершился с кодом {rc}")
        _write_state({"ok": False, "step": "consolidate", "exit": rc,
                      "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        return 10
    rc, _ = run("build_pdf.py")
    if rc != 0:
        log(f"[ERR] build_pdf завершился с кодом {rc}")
        _write_state({"ok": False, "step": "build_pdf", "exit": rc,
                      "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        return 11

    xlsx, pdf = find_latest()
    # дата отчёта
    rd = datetime.now().strftime("%Y-%m-%d")
    meta = LOGS / "last_fetch.meta"
    if meta.exists():
        for line in meta.read_text(encoding="utf-8").splitlines():
            if line.startswith("date="): rd = line.split("=", 1)[1].strip()

    notify_bot(env, rd, xlsx, pdf)

    _write_state({"ok": True, "report_date": rd,
                  "xlsx": str(xlsx) if xlsx else None,
                  "pdf": str(pdf) if pdf else None,
                  "pending_send": True,
                  "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                  "duration_sec": round(time.time() - t0, 1)})
    log(f"RUN_DAILY DONE за {round(time.time()-t0,1)}s")
    return code


def _write_state(d: dict) -> None:
    (STATE / "last_run.json").write_text(
        json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    sys.exit(main())
