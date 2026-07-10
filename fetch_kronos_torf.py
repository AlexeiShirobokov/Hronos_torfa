"""
Скачивание Excel-вложений из папки Yandex 360 за последнюю доступную дату.

Запуск:  python3 fetch_kronos_torf.py
Требования: только стандартная библиотека Python 3.8+.

Что делает:
1. Читает учётные данные из .env рядом со скриптом.
2. Подключается к IMAP Яндекс 360 (imap.yandex.ru:993, SSL).
3. Открывает папку IMAP_FOLDER (по умолчанию Hronos_torfa).
4. Находит самую позднюю дату писем (по дате внутри письма).
5. Скачивает Excel-вложения (.xlsx, .xls, .xlsm, .xlsb) только из этих писем
   в input/mail_attachments/.
6. Пишет лог в logs/fetch_<timestamp>.log и краткий итог в stdout.

Ничего не удаляет и не отправляет.
"""
from __future__ import annotations

import email
import email.header
import email.utils
import imaplib
import json
import os
import re
import socket
import sys
from datetime import datetime, date, timedelta
from pathlib import Path

imaplib._MAXLINE = 10_000_000
# сокет-таймаут на все сетевые операции: fetch не зависнет навсегда, если IMAP
# перестанет отвечать. Без этого зависший FETCH блокирует launchd и пропускает
# следующие запуски (инцидент 2026-07-01: висел 2.5 ч, пропустил 07:10 и 08:10).
socket.setdefaulttimeout(120)

BASE = Path(__file__).resolve().parent
ENV_PATH = BASE / ".env"
ATTACH_DIR = BASE / "input" / "mail_attachments"
LOG_DIR = BASE / "logs"
STATE = BASE / "state"
EXCEL_EXT = {".xlsx", ".xls", ".xlsm", ".xlsb"}


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

        # Инкремент по UID: качаем ТОЛЬКО новые письма (которых ещё не видели) из окна
        # последних N дней. Это и есть «перепроверять новые каждый раз» — без перекачки
        # всего ящика и без роста папки дублями.
        fetch_days = max(1, int(env.get("FETCH_DAYS", "7")))
        cutoff = date.today() - timedelta(days=fetch_days)
        _MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        since_str = f"{cutoff.day:02d}-{_MON[cutoff.month - 1]}-{cutoff.year}"
        typ, data = M.uid("SEARCH", None, "SINCE", since_str)
        win_uids = [u.decode() for u in (data[0].split() if data and data[0] else [])]
        log(f"[INFO] UID-писем за окно (SINCE {since_str}, {fetch_days} дн): {len(win_uids)}")
        if not win_uids:
            log("[WARN] Нет писем за окно — нечего скачивать.")
            return 0

        # latest_date по INTERNALDATE окна (метка для consolidate), без тел писем
        latest_date = cutoff
        typ, resp = M.uid("FETCH", ",".join(win_uids), "(INTERNALDATE)")
        for item in (resp or []):
            if not isinstance(item, (bytes, bytearray)):
                continue
            mm = re.search(rb'INTERNALDATE "([^"]+)"', item)
            if mm:
                try:
                    latest_date = max(latest_date,
                                      email.utils.parsedate_to_datetime(mm.group(1).decode()).date())
                except Exception:
                    pass

        # карта uid → ключ файла: хвост имени Excel-вложения из BODYSTRUCTURE
        # (русские имена MIME-кодированы; хвост стабильно идентифицирует подразделение),
        # без скачивания тел писем — дёшево
        uid_key: dict[str, str] = {}
        typ, resp = M.uid("FETCH", ",".join(win_uids), "(BODYSTRUCTURE)")
        for item in (resp or []):
            blob = item[0] if isinstance(item, tuple) else item
            if not isinstance(blob, (bytes, bytearray)):
                continue
            s = blob.decode("utf-8", errors="replace")
            mu = re.search(r"UID (\d+)", s)
            if not mu:
                continue
            for w in re.findall(r"=\?[^?]+\?[BbQq]\?[^?]+\?=", s):
                try:
                    dn = str(email.header.make_header(email.header.decode_header(w)))
                except Exception:
                    continue
                if Path(dn).suffix.lower() in EXCEL_EXT:
                    uid_key[mu.group(1)] = dn
                    break

        # новейший uid на каждый файл-источник (UID растёт со временем)
        latest_uid: dict[str, str] = {}
        for uid, key in uid_key.items():
            if key not in latest_uid or int(uid) > int(latest_uid[key]):
                latest_uid[key] = uid
        log(f"[INFO] Файлов-источников в окне: {len(latest_uid)}")

        # что уже скачано (ключ → последний UID): качаем только если появился новее
        STATE.mkdir(parents=True, exist_ok=True)
        bu_path = STATE / "base_uid.json"
        base_uid: dict[str, str] = {}
        if bu_path.exists():
            try:
                base_uid = json.loads(bu_path.read_text(encoding="utf-8"))
            except Exception:
                base_uid = {}
        to_dl = [(k, u) for k, u in latest_uid.items() if base_uid.get(k) != u]
        log(f"[INFO] Новее прежнего (к загрузке): {len(to_dl)}")

        saved = 0
        for key, uid in to_dl:
            typ, msg_data = M.uid("FETCH", uid, "(RFC822)")
            if typ != "OK" or not msg_data or not msg_data[0]:
                log(f"[WARN] Не удалось получить письмо uid={uid}")
                continue
            msg = email.message_from_bytes(msg_data[0][1])
            log(f"[MAIL] uid={uid} subj={decode_mime(msg.get('Subject', ''))!r}")
            for part in msg.walk():
                if part.is_multipart():
                    continue
                fname = part.get_filename()
                if not fname:
                    continue
                fname = safe_filename(fname)
                if Path(fname).suffix.lower() not in EXCEL_EXT:
                    continue
                payload = part.get_payload(decode=True)
                if not payload:
                    continue
                # пишем свежую версию как базовое имя, удалив прежние версии этого файла
                # (включая старые суффиксы __N) — «последняя версия» = только что скачанная
                base_name = re.sub(r"__\d+(?=\.[^.]+$)", "", fname)
                source_key = _source_key(base_name)
                for old in list(ATTACH_DIR.glob("*")):
                    if old.is_file() and _source_key(re.sub(r"__\d+(?=\.[^.]+$)", "", old.name)) == source_key:
                        try:
                            old.unlink()
                        except OSError:
                            pass
                dest = ATTACH_DIR / base_name
                dest.write_bytes(payload)
                saved += 1
                log(f"   [SAVE] {dest.name} ({len(payload)} bytes)")
            base_uid[key] = uid
        bu_path.write_text(json.dumps(base_uid, ensure_ascii=False, indent=2), encoding="utf-8")
        log(f"[INFO] Файлов во вложениях: {sum(1 for _ in ATTACH_DIR.iterdir())}")

        log(f"[DONE] Скачано вложений: {saved}")
        log(f"[DONE] Дата писем: {latest_date.isoformat()}")
        log(f"[DONE] Папка вложений: {ATTACH_DIR}")
        # запишем метаинформацию для следующего шага
        meta = BASE / "logs" / "last_fetch.meta"
        meta.write_text(
            f"date={latest_date.isoformat()}\n"
            f"mails={len(to_dl)}\n"
            f"files={saved}\n"
            f"window_uids={len(win_uids)}\n"
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


def _source_key(name: str) -> str:
    path = Path(name)
    return re.sub(r"\s+", " ", path.with_suffix("").name).strip().casefold()


if __name__ == "__main__":
    sys.exit(main())
