from __future__ import annotations

import argparse
import logging
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from .consolidate import consolidate


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="hronos", description="Параллельная консолидация хронометража")
    parser.add_argument("--log-level", default="INFO", choices=["ERROR", "WARNING", "INFO", "DEBUG"])
    sub = parser.add_subparsers(dest="cmd", required=True)

    for name in ("consolidate", "run"):
        p = sub.add_parser(name)
        p.add_argument("--input", default="../input/mail_attachments", help="Папка с .xlsx/.xlsb")
        p.add_argument("--output", default="output", help="Папка результата")
        p.add_argument("--report-date", default=None, help="Дата выдачи наряд-задания YYYY-MM-DD; по умолчанию вчера")
        p.add_argument("--fresh-date", default=None, help="Контрольная дата свежести YYYY-MM-DD; по умолчанию сегодня")
        p.add_argument("--sheet", default="Реестр", help="Имя листа-источника")
        p.add_argument("--required-unit", action="append", default=[], help="Обязательное подразделение для контроля свежести")
        p.add_argument("--allow-stale", action="store_true", help="Собрать файл даже если часть подразделений без свежего исходника")

    run_p = sub.choices["run"]
    run_p.add_argument("--no-fetch", action="store_true", help="Не скачивать новые исходники перед сборкой")
    run_p.add_argument("--fetch-script", default=None, help="Путь к fetch_kronos_torf.py")
    run_p.add_argument("--dotnet", default="dotnet", help="Команда dotnet")
    run_p.add_argument("--pivot-project", default=None, help="Путь к C# проекту PivotBuilder")
    run_p.add_argument("--keep-plain", action="store_true", help="Оставить также книгу до C#-постобработки")
    run_p.add_argument(
        "--peredel",
        action="append",
        default=[],
        help="Передел для сводной; можно указать несколько раз. По умолчанию: Транспортировка песков + Подача песков",
    )

    args = parser.parse_args(argv)
    logging.basicConfig(level=getattr(logging, args.log_level), format="[%(levelname)s] %(message)s")

    report_date = _parse_report_date(args.report_date)
    freshness_date = _parse_fresh_date(args.fresh_date)
    try:
        if args.cmd == "run" and not args.no_fetch:
            _run_fetch(args)
        result = consolidate(
            Path(args.input),
            Path(args.output),
            report_date=report_date,
            sheet_name=args.sheet,
            pivot_peredels=args.peredel or None,
            freshness_date=freshness_date,
            required_units=args.required_unit or None,
        )
        print(f"[OK] Python: {result.rows} строк -> {result.workbook}")
        stale = _stale_units(result.freshness_check)
        if stale and not args.allow_stale:
            print("[ERR] нет свежих исходников по подразделениям:", file=sys.stderr)
            for row in stale:
                print(
                    f"  - {row['Подразделение']}: макс. Дата. Факт={row['Макс_Дата_Факт'] or '-'}, "
                    f"нужно {row['Контрольная_дата']}; файл: {row['Файлы'] or '-'}",
                    file=sys.stderr,
                )
            print(f"[ERR] сборка остановлена до C#; проверка записана в {result.workbook}", file=sys.stderr)
            return 1
        if args.cmd == "consolidate":
            return 0
        final_path = _run_pivot_builder(args, result.workbook, report_date)
        if not args.keep_plain and final_path != result.workbook and result.workbook.exists():
            result.workbook.unlink()
        print(f"[OK] C#: финальная книга -> {final_path}")
        return 0
    except Exception as exc:
        print(f"[ERR] {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


def _parse_report_date(value: str | None) -> date:
    if not value:
        return date.today() - timedelta(days=1)
    return datetime.strptime(value, "%Y-%m-%d").date()


def _parse_fresh_date(value: str | None) -> date:
    if not value:
        return date.today()
    return datetime.strptime(value, "%Y-%m-%d").date()


def _run_fetch(args) -> None:
    root = Path(__file__).resolve().parents[2]
    script = Path(args.fetch_script) if args.fetch_script else root / "fetch_kronos_torf.py"
    proc = subprocess.run([sys.executable, str(script)], cwd=str(root), capture_output=True, text=True, timeout=600)
    if proc.stdout.strip():
        print(proc.stdout.strip())
    if proc.stderr.strip():
        print(proc.stderr.strip(), file=sys.stderr)
    if proc.returncode != 0:
        raise RuntimeError(f"fetch exit={proc.returncode}")


def _stale_units(freshness_check):
    if freshness_check.empty or "Статус" not in freshness_check:
        return []
    stale = freshness_check[freshness_check["Статус"] != "OK"]
    return stale.to_dict(orient="records")


def _run_pivot_builder(args, workbook: Path, report_date: date) -> Path:
    project = Path(args.pivot_project) if args.pivot_project else Path(__file__).resolve().parents[1] / "tools" / "pivot" / "HronosPivotBuilder.csproj"
    root = Path(__file__).resolve().parents[2]
    template = root / "tools" / "pivot" / "pivot_template.xlsx"
    final_path = workbook.with_name(workbook.stem + "_final.xlsx")
    cmd = [
        args.dotnet,
        "run",
        "--project",
        str(project),
        "--configuration",
        "Release",
        "--",
        str(workbook),
        "--out",
        str(final_path),
        "--date",
        report_date.isoformat(),
        "--template",
        str(template),
    ]
    for peredel in args.peredel:
        cmd.extend(["--peredel", peredel])
    proc = subprocess.run(cmd, cwd=str(project.parents[2]), capture_output=True, text=True, timeout=300)
    if proc.stdout.strip():
        print(proc.stdout.strip())
    if proc.stderr.strip():
        print(proc.stderr.strip(), file=sys.stderr)
    if proc.returncode != 0:
        raise RuntimeError(f"C# PivotBuilder exit={proc.returncode}")
    return final_path


if __name__ == "__main__":
    raise SystemExit(main())
