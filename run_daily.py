"""Оркестратор: fetch → consolidate → build_xlsx → explain → алерты → send_email (авто).
Запускается launchd-ом каждый час в xx:10 (6–22). Доставка — только email.
"""
from __future__ import annotations
import json, smtplib, ssl, subprocess, sys, time
from datetime import datetime, date, timedelta
from email.message import EmailMessage
from pathlib import Path

BASE = Path(__file__).resolve().parent
LOGS = BASE / "logs"
STATE = BASE / "state"
OUT = BASE / "output"
PY = sys.executable
# Нативная сводная Excel: финальную книгу собирает BookBuilder (C#/EPPlus) из шаблона
# со ГОТОВОЙ сводной на умной таблице «Таблица1», подменяя данные реестра свежими.
DOTNET = "/usr/local/share/dotnet/dotnet"
BOOKBUILDER_DLL = BASE / "tools" / "pivot" / "bin" / "Release" / "net9.0" / "PivotBuilder.dll"
PIVOT_TEMPLATE = BASE / "tools" / "pivot" / "pivot_template.xlsx"
# Алерты (аномалии volume_drop/idle/… и технические сбои) временно ОТКЛЮЧЕНЫ:
# получателям они шли как путаница. Отправляем только стандартный отчёт.
# Вернуть — поставить True (тогда уйдут всем из recipients.txt).
ALERTS_ENABLED = False
# Авто-рассылка получателям временно ОТКЛЮЧЕНА (по просьбе Алексея, 2026-07-02):
# дорабатываем нативную сводную, не спамим 5 боевых получателей. Пайплайн продолжает
# собирать книгу (fetch→consolidate→build_xlsx→pivot), но письмо не отправляет.
# Вернуть рассылку — поставить True.
SEND_ENABLED = False


def load_env() -> dict:
    p = BASE / ".env"
    out = {}
    if not p.exists():
        return out
    for raw in p.read_text(encoding="utf-8").splitlines():
        s = raw.strip()
        if not s or s.startswith("#") or "=" not in s:
            continue
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
    if p.stdout:
        log(p.stdout.strip()[:2000])
    if p.stderr:
        log("STDERR: " + p.stderr.strip()[:2000])
    log(f"   exit={p.returncode}")
    return p.returncode, p.stdout


def _alert_recipients(env: dict) -> list[str]:
    """Получатели алертов = список рассылки (recipients.txt); фолбэк — отправитель."""
    try:
        import send_email
        rcpts = send_email.load_recipients()
    except Exception:
        rcpts = []
    if not rcpts:
        login = env.get("YANDEX_LOGIN")
        rcpts = [login] if login else []
    return rcpts


def _send_alert_email(env: dict, subject: str, body: str) -> None:
    rcpts = _alert_recipients(env)
    if not rcpts:
        raise RuntimeError("нет получателей для алерта (recipients.txt пуст)")
    msg = EmailMessage()
    msg["From"] = env.get("YANDEX_LOGIN", "")
    msg["To"] = ", ".join(rcpts)
    msg["Subject"] = f"[hronos_torfa] {subject}"
    msg.set_content(body)
    host = env.get("SMTP_HOST", "smtp.yandex.ru")
    port = int(env.get("SMTP_PORT", "465"))
    login = env.get("SMTP_LOGIN") or env.get("YANDEX_LOGIN")
    password = env.get("SMTP_PASSWORD") or env.get("YANDEX_APP_PASSWORD")
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
        s.login(login, password)
        s.send_message(msg)


def alert(env: dict, subject: str, body: str, key: str) -> None:
    """Email-алерт с анти-спамом: один и тот же key за дату отправляется один раз."""
    if not ALERTS_ENABLED:
        log(f"[alert] отключены — пропуск: {key}")
        return
    STATE.mkdir(parents=True, exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")
    full = f"{today}|{key}"
    sent_p = STATE / "alerts_sent.json"
    sent = []
    if sent_p.exists():
        try:
            sent = json.loads(sent_p.read_text(encoding="utf-8"))
        except Exception:
            sent = []
    sent = [k for k in sent if k.startswith(today)]  # чистим прошлые даты
    if full in sent:
        log(f"[alert] подавлен дубль: {key}")
        return
    try:
        _send_alert_email(env, subject, body)
        log(f"[alert] отправлен: {key}")
    except Exception as e:
        log(f"[alert][ERR] {key}: {e!r}")
        return
    sent.append(full)
    sent_p.write_text(json.dumps(sent, ensure_ascii=False, indent=2), encoding="utf-8")


def _report_date() -> str:
    rd = datetime.now().strftime("%Y-%m-%d")
    meta = LOGS / "last_fetch.meta"
    if meta.exists():
        for line in meta.read_text(encoding="utf-8").splitlines():
            if line.startswith("date="):
                rd = line.split("=", 1)[1].strip()
    return rd


def _write_state(d: dict) -> None:
    STATE.mkdir(parents=True, exist_ok=True)
    (STATE / "last_run.json").write_text(
        json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def build_pivot_book() -> None:
    """Собирает финальную книгу из шаблона (готовая нативная сводная на умной таблице
    «Таблица1»), подменяя данные реестра свежими, затем пост-обработка (ресайз таблицы +
    refreshOnLoad). Не-фатально: при сбое отчёт всё равно уйдёт (без обновлённой сводной).
    """
    try:
        books = sorted(OUT.glob("Хронометраж_транспортировки_торфов_*.xlsx"),
                       key=lambda p: p.stat().st_mtime, reverse=True)
        if not books:
            log("[pivot][WARN] нет книги в output/ — пропуск сводной")
            return
        book = books[0]
        if not PIVOT_TEMPLATE.exists():
            log(f"[pivot][WARN] нет шаблона {PIVOT_TEMPLATE.name} — книга без сводной")
            return
        dotnet = DOTNET if Path(DOTNET).exists() else "dotnet"
        if not BOOKBUILDER_DLL.exists():
            log(f"[pivot][WARN] нет {BOOKBUILDER_DLL.name} — собрать: "
                f"dotnet build -c Release в tools/pivot; книга без сводной")
            return
        log(f"→ BookBuilder {book.name}")
        p = subprocess.run([dotnet, str(BOOKBUILDER_DLL), str(book), str(PIVOT_TEMPLATE)],
                           cwd=str(BASE), capture_output=True, text=True, timeout=300)
        if p.stdout:
            log("[pivot] " + p.stdout.strip()[:500])
        if p.stderr:
            log("[pivot] STDERR: " + p.stderr.strip()[:500])
        if p.returncode != 0:
            log(f"[pivot][WARN] BookBuilder exit={p.returncode} — книга без обновлённой сводной")
            return
        import pivot_postprocess
        ref = pivot_postprocess.postprocess(book)
        log(f"[pivot] сводная материализована из «Таблица1» {ref}, refreshOnLoad=0")
    except Exception as e:
        log(f"[pivot][WARN] сбой сборки сводной: {e!r} — продолжаю без неё")


def main() -> int:
    STATE.mkdir(parents=True, exist_ok=True)
    LOGS.mkdir(parents=True, exist_ok=True)
    env = load_env()
    t0 = time.time()
    log("=" * 60)
    log("RUN_DAILY START")

    rc, _ = run("fetch_kronos_torf.py")
    if rc != 0:
        log("[WARN] fetch с ошибкой, продолжаю с локальными вложениями")

    rd = _report_date()
    # отчётная дата штатно отстаёт на день (письма за вчерашнюю смену);
    # алерт только если данные старше вчерашних (2+ дня) — значит письма не приходят
    try:
        rd_date = datetime.strptime(rd, "%Y-%m-%d").date()
        stale = rd_date < date.today() - timedelta(days=1)
    except ValueError:
        stale = False
    if stale:
        alert(env, "нет свежих данных",
              f"Дата отчёта {rd} устарела (старше вчерашней) — возможно, не приходят новые письма.",
              key=f"stale_data:{rd}")

    rc, _ = run("consolidate.py")
    if rc != 0:
        alert(env, "сбой consolidate", f"consolidate.py exit={rc}", key=f"fail_consolidate:{rd}")
        _write_state({"ok": False, "step": "consolidate", "exit": rc,
                      "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        return 10

    rc, _ = run("build_xlsx.py")
    if rc != 0:
        alert(env, "сбой build_xlsx", f"build_xlsx.py exit={rc}", key=f"fail_build_xlsx:{rd}")
        _write_state({"ok": False, "step": "build_xlsx", "exit": rc,
                      "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        return 11

    # финальная книга из шаблона с нативной сводной на умной таблице (до отправки письма)
    build_pivot_book()

    # метрики уже записаны build_xlsx → state/last_metrics.json
    rc, _ = run("explain.py")
    if rc != 0:
        log(f"[WARN] explain.py exit={rc} — продолжаю без записки")

    # email-алерты об аномалиях
    alerts_p = STATE / "last_alerts.json"
    if alerts_p.exists():
        try:
            for a in json.loads(alerts_p.read_text(encoding="utf-8")):
                alert(env, f"аномалия: {a['type']}", a["text"], key=f"{a['type']}:{rd}")
        except Exception as e:
            log(f"[WARN] чтение last_alerts.json: {e!r}")

    # авто-рассылка (временно отключена флагом SEND_ENABLED)
    if SEND_ENABLED:
        rc, _ = run("send_email.py")
        if rc != 0:
            alert(env, "сбой отправки письма", f"send_email.py exit={rc}", key=f"fail_send:{rd}")
    else:
        log("[send] рассылка получателям ОТКЛЮЧЕНА (SEND_ENABLED=False) — письмо не отправлено")

    _write_state({"ok": True, "report_date": rd,
                  "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                  "duration_sec": round(time.time() - t0, 1)})
    log(f"RUN_DAILY DONE за {round(time.time()-t0,1)}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
