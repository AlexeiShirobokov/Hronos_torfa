"""
Скачивание Excel-вложений из папки Yandex 360 за последнюю доступную дату.

Запуск:  python3 fetch_kronos_torf.py
Требования: только стандартная библиотека Python 3.8+.

Что делает:
1. Читает учётные данные из .env рядом со скриптом.
2. Подключается к IMAP Яндекс 360 (imap.yandex.ru:993, SSL).
3. Открывает папку IMAP_FOLDER (по умолчанию Hronos_torfa).
4. Находит самую позднюю дату писем (по дате внутри письма).
5. Скачивает Excel-вложения (.xlsx, .xls, .xlsm) только из этих писем
   в input/mail_attachments/.
6. Пишет лог в logs/fetch_<timestamp>.log и краткий итог в stdout.

Ничего не удаляет и не отправляет.
"""
from __future__ import annotations

import email
import email.header
import email.utils
import imaplib
import os
import re
import sys
from datetime import datetime, date
from pathlib import Path

imaplib._MAXLINE = 10_000_000

BASE = Path(__file__).resolve().parent
ENV_PATH = BASE / ".env"
ATTACH_DIR = BASE / "input" / "mail_attachments"
LOG_DIR = BASE / "logs"
EXCEL_EXT = {".xlsx", ".xls", ".xlsm"}


def load_env(path: Path) -> dict:
    data = {}
    if not path.exists():
        raise SystemExit(f"[ERR] .env не найден: {path}")
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        v = v.strip().strip('"').strip("'")
        data[k.strip()] = v
    return data


def decode_mime(value: str) -> str:
    if not value:
        return ""
    parts = email.header.decode_header(value)
    out = []
    for txt, enc in parts:
        if isinstance(txt, bytes):
            try:
                out.append(txt.decode(enc or "utf-8", errors="replace"))
            except LookupError:
                out.append(txt.decode("utf-8", errors="replace"))
        else:
            out.append(txt)
    return "".join(out)


def safe_filename(name: str) -> str:
    name = decode_mime(name or "")
    name = name.replace("\\", "_").replace("/", "_")
    name = re.sub(r'[<>:"|?*\x00-\x1f]+', "_", name).strip().strip(".")
    return name or "attachment.bin"


def find_folder(M: imaplib.IMAP4_SSL, wanted: str) -> str:
    """Найти реальное IMAP-имя папки по человекочитаемому имени."""
    typ, data = M.list()
    if typ != "OK":
        raise SystemExit("[ERR] LIST не удался")
    candidates = []
    for raw in data:
        if not raw:
            continue
        line = raw.decode("utf-8", errors="replace")
        # формат:  (\HasNoChildren) "|" "Имя папки"
        m = re.match(r'\([^)]*\)\s+"[^"]*"\s+"?([^"]+)"?\s*$', line)
        if not m:
            continue
        mbox = m.group(1)
        candidates.append(mbox)
        if mbox == wanted:
            return mbox
    # поиск без учёта регистра и подстрокой
    wanted_l = wanted.lower()
    for mbox in candidates:
        if mbox.lower() == wanted_l:
            return mbox
    for mbox in candidates:
        if wanted_l in mbox.lower():
            return mbox
    # вернём первый кандидат, чтобы вызвать select и получить осмысленную ошибку
    sys.stderr.write("[WARN] точное совпадение не найдено. Доступные папки:\n")
    for mbox in candidates:
        sys.stderr.write(f"   {mbox}\n")
    raise SystemExit(f"[ERR] папка '{wanted}' не найдена в ящике")


def main() -> int:
    env = load_env(ENV_PATH)
    login = env.get("YANDEX_LOGIN", "").strip()
    pwd = env.get("YANDEX_APP_PASSWORD", "").strip()
    host = env.get("IMAP_HOST", "imap.yandex.ru").strip()
    port = int(env.get("IMAP_PORT", "993").strip() or "993")
    folder = env.get("IMAP_FOLDER", "Hronos_torfa").strip()

    if "@" not in login:
        sys.stderr.write(
            f"[ERR] В YANDEX_LOGIN ('{login}') похоже пропущен символ '@'. "
            "Укажи полный email вида name@domain.ru и запусти снова.\n"
        )
        return 2
    if not pwd:
        sys.stderr.write("[ERR] YANDEX_APP_PASSWORD пуст.\n")
        return 2

    ATTACH_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_path = LOG_DIR / f"fetch_{datetime.now():%Y%m%d_%H%M%S}.log"
    log_f = log_path.open("w", encoding="utf-8")

    def log(msg: str) -> None:
        print(msg)
        log_f.write(msg + "\n")
        log_f.flush()

    log(f"[INFO] Запуск {datetime.now():%Y-%m-%d %H:%M:%S}")
    log(f"[INFO] Хост: {host}:{port}, логин: {login}, папка: {folder}")

    try:
        M = imaplib.IMAP4_SSL(host, port)
    except Exception as e:
        log(f"[ERR] Не удалось подключиться: {e!r}")
        return 3
    try:
        M.login(login, pwd)
    except imaplib.IMAP4.error as e:
        log(f"[ERR] Логин не принят: {e!r}\n"
            "      Проверь, что используется ПАРОЛЬ ПРИЛОЖЕНИЯ для IMAP\n"
            "      (id.yandex.ru → Безопасность → Пароли приложений).")
        return 4

    try:
        real_folder = find_folder(M, folder)
        log(f"[INFO] IMAP-имя папки: {real_folder!r}")

        typ, _ = M.select(f'"{real_folder}"', readonly=True)
        if typ != "OK":
            log(f"[ERR] SELECT не удался для {real_folder!r}")
            return 5

        typ, data = M.search(None, "ALL")
        ids = data[0].split() if data and data[0] else []
        log(f"[INFO] Всего писем в папке: {len(ids)}")
        if not ids:
            log("[WARN] Папка пуста — нечего скачивать.")
            return 0

        # 1-й проход: считываем INTERNALDATE, чтобы найти последнюю дату
        latest_date: date | None = None
        date_by_id: dict[bytes, date] = {}
        BATCH_SIZE = 200
        id_list = [b.decode() for b in ids]
        for chunk_start in range(0, len(id_list), BATCH_SIZE):
            chunk = id_list[chunk_start:chunk_start + BATCH_SIZE]
            batch = ",".join(chunk)
            typ, resp = M.fetch(batch, "(INTERNALDATE)")
            if typ != "OK":
                log(f"[ERR] FETCH INTERNALDATE не удался (chunk {chunk_start})")
                return 6
            for item in resp:
                if not item:
                    continue
                s = item.decode("utf-8", errors="replace") if isinstance(item, bytes) else str(item)
                m_id = re.match(r"(\d+)\s+\(INTERNALDATE\s+\"([^\"]+)\"\)", s)
                if not m_id:
                    continue
                num = m_id.group(1).encode()
                dt = email.utils.parsedate_to_datetime(m_id.group(2))
                d = dt.date()
                date_by_id[num] = d
                if latest_date is None or d > latest_date:
                    latest_date = d

        if latest_date is None:
            log("[ERR] Не удалось определить даты писем.")
            return 7
        log(f"[INFO] Последняя дата в папке: {latest_date.isoformat()}")

        target_ids = [n for n, d in date_by_id.items() if d == latest_date]
        log(f"[INFO] Писем за последнюю дату: {len(target_ids)}")

        saved = 0
        for num in target_ids:
            typ, msg_data = M.fetch(num, "(RFC822)")
            if typ != "OK" or not msg_data or not msg_data[0]:
                log(f"[WARN] Не удалось получить письмо {num!r}")
                continue
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)
            subj = decode_mime(msg.get("Subject", ""))
            log(f"[MAIL] id={num.decode()} subj={subj!r}")
            for part in msg.walk():
                if part.is_multipart():
                    continue
                disp = (part.get("Content-Disposition") or "").lower()
                fname = part.get_filename()
                if not fname:
                    continue
                fname = safe_filename(fname)
                ext = Path(fname).suffix.lower()
                if ext not in EXCEL_EXT:
                    log(f"   [SKIP] {fname} (не Excel)")
                    continue
                payload = part.get_payload(decode=True)
                if not payload:
                    log(f"   [WARN] пустое вложение {fname}")
                    continue
                dest = ATTACH_DIR / fname
                # избегаем перезаписи
                if dest.exists():
                    stem, suf = dest.stem, dest.suffix
                    i = 1
                    while True:
                        cand = ATTACH_DIR / f"{stem}__{i}{suf}"
                        if not cand.exists():
                            dest = cand
                            break
                        i += 1
                dest.write_bytes(payload)
                saved += 1
                log(f"   [SAVE] {dest.name} ({len(payload)} bytes)")

        log(f"[DONE] Скачано вложений: {saved}")
        log(f"[DONE] Дата писем: {latest_date.isoformat()}")
        log(f"[DONE] Папка вложений: {ATTACH_DIR}")
        # запишем метаинформацию для следующего шага
        meta = BASE / "logs" / "last_fetch.meta"
        meta.write_text(
            f"date={latest_date.isoformat()}\n"
            f"mails={len(target_ids)}\n"
            f"files={saved}\n"
            f"folder={real_folder}\n",
            encoding="utf-8",
        )
        return 0
    finally:
        try:
            M.logout()
        except Exception:
            pass
        log_f.close()


if __name__ == "__main__":
    sys.exit(main())
