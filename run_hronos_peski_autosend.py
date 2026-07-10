#!/usr/bin/env python3
from __future__ import annotations

import argparse
import fcntl
import json
import re
import shutil
import subprocess
import sys
import zipfile
from datetime import datetime, time, timedelta
from pathlib import Path

BASE = Path(__file__).resolve().parent
PARALLEL = BASE / "parallel_hronos"
OUT = BASE / "output"
LOGS = BASE / "logs"
STATE = BASE / "state"
LOCK = STATE / "hronos_peski_autosend.lock"
LOG = LOGS / "hronos_peski_autosend.log"
TARGET_NAME = "Хронометраж транспортировки торфов и песков.xlsx"
TARGET = OUT / TARGET_NAME
SUBJECT = "Хронометраж транспортировки торфов и песков"
DOTNET = Path("/usr/local/share/dotnet/dotnet")
TODAY_REPORT_CUTOFF = time(10, 20)


def resolve_dotnet() -> str:
    for candidate in (DOTNET, Path.home() / ".dotnet" / "dotnet"):
        if candidate.exists():
            return str(candidate)
    return shutil.which("dotnet") or "dotnet"


def log(message: str) -> None:
    LOGS.mkdir(parents=True, exist_ok=True)
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}"
    print(line)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def run_cmd(cmd: list[str], *, cwd: Path, timeout: int) -> subprocess.CompletedProcess[str]:
    return subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, timeout=timeout)


def short_output(text: str, limit: int = 1800) -> str:
    text = text.strip()
    if len(text) <= limit:
        return text
    return text[:limit] + "\n...<cut>"


def log_key_lines(text: str) -> None:
    keys = ("[INFO]", "[OK]", "[DONE]", "[ERR]", "финальная книга", "Скачано вложений")
    for line in text.splitlines():
        if any(key in line for key in keys):
            log(line)


def parse_final_path(text: str) -> Path | None:
    match = re.search(r"финальная книга -> (.+?\.xlsx)", text)
    if not match:
        return None
    return Path(match.group(1).strip())


def latest_final() -> Path:
    finals = sorted(PARALLEL.glob("output/*_final.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not finals:
        raise RuntimeError("не найден финальный *_final.xlsx после сборки")
    return finals[0]


def metadata_path(final_path: Path) -> Path:
    if final_path.name.endswith("_final.xlsx"):
        return final_path.with_name(final_path.name[: -len("_final.xlsx")] + ".json")
    return final_path.with_suffix(".json")


def validate_metadata(final_path: Path) -> dict:
    meta_path = metadata_path(final_path)
    if not meta_path.exists():
        raise RuntimeError(f"не найден файл контроля: {meta_path}")
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    if not meta.get("ok"):
        raise RuntimeError("контрольный json: ok != true")

    volume_bad = [r for r in meta.get("volume_check", []) if r.get("Статус") != "OK"]
    fresh_bad = [r for r in meta.get("freshness_check", []) if r.get("Статус") != "OK"]
    if volume_bad:
        raise RuntimeError("сверка объемов не прошла: " + ", ".join(str(r.get("Подразделение")) for r in volume_bad))
    if fresh_bad:
        log(
            "[WARN] неполная свежесть, отправка не блокируется: "
            + ", ".join(
                f"{r.get('Подразделение')} max={r.get('Макс_Дата_Факт') or '-'}"
                for r in fresh_bad
            )
        )

    volumes = ", ".join(
        f"{r.get('Подразделение')}={r.get('Объем_реестр')}" for r in meta.get("volume_check", [])
    )
    fresh = ", ".join(
        f"{r.get('Подразделение')}:{r.get('Макс_Дата_Факт')}" for r in meta.get("freshness_check", [])
    )
    log(f"{'свежесть' if fresh_bad else 'свежесть OK'}: {fresh}")
    log(f"объемы OK: {volumes}")
    return meta


def validate_xlsx(path: Path) -> None:
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad:
            raise RuntimeError(f"битый xlsx, первый проблемный файл: {bad}")
        table_xml = zf.read("xl/tables/table1.xml").decode("utf-8")
        pivot_xml = zf.read("xl/pivotCache/pivotCacheDefinition1.xml").decode("utf-8")
    table_name = re.search(r'displayName="([^"]+)"', table_xml)
    table_ref = re.search(r'ref="([^"]+)"', table_xml)
    if not table_name or table_name.group(1) != "Таблица1":
        raise RuntimeError("на листе реестра не найдена умная таблица Таблица1")
    if "Таблица1" not in pivot_xml:
        raise RuntimeError("сводная не ссылается на Таблица1")
    log(f"xlsx OK: {path.name}, Таблица1 {table_ref.group(1) if table_ref else '-'}")


def report_date_for_run(now: datetime | None = None) -> str:
    now = now or datetime.now()
    report_date = now.date() if now.time() >= TODAY_REPORT_CUTOFF else now.date() - timedelta(days=1)
    return report_date.isoformat()


def build_book(no_fetch: bool) -> Path:
    dotnet = resolve_dotnet()
    report_date = report_date_for_run()
    cmd = [
        sys.executable,
        "-m",
        "hronos.cli",
        "run",
        "--dotnet",
        dotnet,
        "--allow-stale",
        "--report-date",
        report_date,
    ]
    if no_fetch:
        cmd.append("--no-fetch")
    log("старт сборки: свежая загрузка=" + ("нет" if no_fetch else "да"))
    log(f"дата фильтра «Дата выдачи наряд-задания»: {report_date}")
    proc = run_cmd(cmd, cwd=PARALLEL, timeout=1200)
    if proc.stdout:
        log_key_lines(proc.stdout)
    if proc.stderr:
        log("[stderr] " + short_output(proc.stderr))
    if proc.returncode != 0:
        raise RuntimeError(f"hronos.cli run exit={proc.returncode}")
    return parse_final_path(proc.stdout + "\n" + proc.stderr) or latest_final()


def send_book(no_send: bool) -> None:
    if no_send:
        log("отправка пропущена: --no-send")
        return
    cmd = [
        sys.executable,
        "send_email.py",
        "--empty-body",
        "--subject",
        SUBJECT,
        "--xlsx",
        str(TARGET),
    ]
    proc = run_cmd(cmd, cwd=BASE, timeout=180)
    if proc.stdout:
        log_key_lines(proc.stdout)
    if proc.stderr:
        log("[send stderr] " + short_output(proc.stderr))
    if proc.returncode != 0:
        raise RuntimeError(f"send_email.py exit={proc.returncode}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Автосборка и рассылка хронометража торфов и песков")
    parser.add_argument("--no-fetch", action="store_true", help="тест: не скачивать новые письма")
    parser.add_argument("--no-send", action="store_true", help="тест: не отправлять письмо")
    args = parser.parse_args(argv)

    STATE.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)
    with LOCK.open("w", encoding="utf-8") as lock_file:
        try:
            fcntl.flock(lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            log("предыдущий запуск еще идет, текущий пропущен")
            return 0
        try:
            log("START")
            final_path = build_book(args.no_fetch)
            validate_metadata(final_path)
            validate_xlsx(final_path)
            shutil.copy2(final_path, TARGET)
            validate_xlsx(TARGET)
            send_book(args.no_send)
            log("DONE")
            return 0
        except Exception as exc:
            log(f"[ERR] {type(exc).__name__}: {exc}")
            return 1


if __name__ == "__main__":
    raise SystemExit(main())
