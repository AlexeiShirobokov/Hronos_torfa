# Hronos_torfa Фаза 1 — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Перевести ежедневный отчёт хронометража торфов на формат «HTML-письмо + Excel с листами аналитики + пояснительная записка от модели», добавить email-алерты, убрать PDF и Telegram-бот; всё работает на Mac (Фаза 1), рассылка с Mac.

**Architecture:** Аггрегация данных выносится из `build_pdf.py` в отдельный модуль `analytics.py` (DataFrame-агрегаты + плоский `metrics`-словарь, сбрасываемый в `state/last_metrics.json`). `build_xlsx.py` дописывает к консолидированному реестру листы аналитики. `explain.py` строит пояснительную записку через SSH-мост к Claude (`llm.py`, как в Geol) с правиловым фолбэком и считает аномалии. `email_html.py` рендерит HTML-тело письма из `metrics` + записки. `run_daily.py` оркеструет шаги, шлёт email-алерты (с анти-спамом) и автоматически отправляет письмо. `build_pdf.py` и `bot.py` выводятся из эксплуатации.

**Tech Stack:** Python 3.12 (framework-интерпретатор), pandas 3.0.3, openpyxl 3.1.5, стандартная библиотека (smtplib/email/ssl, subprocess для SSH-моста), `unittest` для тестов.

## Global Constraints

- Интерпретатор: `/Library/Frameworks/Python.framework/Versions/3.12/bin/python3` (далее `$PY`). Запуск тестов — `$PY -m unittest`.
- Зависимости — только уже установленные (pandas, openpyxl) + стандартная библиотека. Новых пакетов не добавлять. pytest НЕ используется.
- Язык всего пользовательского текста (письма, записка, алерты, README) — русский.
- Единственный активный получатель — `alexeimvc@gmail.com`; остальные адреса в `recipients.txt` закомментированы (`#`).
- Мост к Claude — best-effort: при недоступности отчёт всё равно собирается (правиловый фолбэк записки). Сбой моста/SMTP не должен валить сборку Excel.
- PDF не формируется; Telegram-кода в пайплайне быть не должно.
- Расписание — каждый час xx:10 в окне 6:00–22:00 (локальное время Mac).
- Частые коммиты: каждый Task завершается коммитом.

---

## File Structure

| Файл | Ответственность |
|---|---|
| `scripts/com.alexei.hronos_torfa.plist` | (modify) расписание launchd 6–22 |
| `recipients.txt` | (modify) активен только `alexeimvc@gmail.com` |
| `llm.py` | (create) SSH-мост к Claude: `available()`, `ask()` |
| `analytics.py` | (create) `compute()` агрегаты, `to_metrics()`, `dump_metrics()` |
| `build_xlsx.py` | (create) дописать листы аналитики в книгу реестра |
| `explain.py` | (create) бриф, записка (мост+фолбэк), аномалии, `main()` |
| `email_html.py` | (create) `build_html()`, `build_text()` тела письма |
| `send_email.py` | (modify) HTML-письмо + вложение Excel, без PDF |
| `run_daily.py` | (modify) новый поток, `alert()` по email, анти-спам, без Telegram |
| `README.md` | (modify) новый формат, окно 6–22, фазы, SMTP-зависимость |
| `tests/` | (create) unittest для analytics/explain/email_html |

Поток `run_daily`: `fetch → consolidate → build_xlsx → explain → (email-алерты об аномалиях) → send_email (авто)`. Алерты о сбоях шага шлются в точке сбоя.

---

## Task 1: Расписание launchd 6–22

**Files:**
- Modify: `scripts/com.alexei.hronos_torfa.plist:18-31`

**Interfaces:**
- Consumes: ничего.
- Produces: launchd-задача запускает `run_daily.py` в xx:10 каждый час 6…22.

- [ ] **Step 1: Заменить блок StartCalendarInterval**

В `scripts/com.alexei.hronos_torfa.plist` заменить весь массив (строки 18–31) на 17 записей 6…22:

```xml
    <key>StartCalendarInterval</key>
    <array>
        <dict><key>Hour</key><integer>6</integer> <key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>7</integer> <key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>8</integer> <key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>9</integer> <key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>10</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>11</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>12</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>13</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>14</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>15</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>16</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>17</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>18</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>19</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>20</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>21</integer><key>Minute</key><integer>10</integer></dict>
        <dict><key>Hour</key><integer>22</integer><key>Minute</key><integer>10</integer></dict>
    </array>
```

- [ ] **Step 2: Проверить валидность plist**

Run: `plutil -lint scripts/com.alexei.hronos_torfa.plist`
Expected: `scripts/com.alexei.hronos_torfa.plist: OK`

- [ ] **Step 3: Перезагрузить задачу (применить расписание)**

Run:
```bash
launchctl unload ~/Library/LaunchAgents/com.alexei.hronos_torfa.plist 2>/dev/null
cp scripts/com.alexei.hronos_torfa.plist ~/Library/LaunchAgents/com.alexei.hronos_torfa.plist
launchctl load ~/Library/LaunchAgents/com.alexei.hronos_torfa.plist
launchctl list | grep hronos_torfa
```
Expected: строка с `com.alexei.hronos_torfa` присутствует.

- [ ] **Step 4: Commit**

```bash
git add scripts/com.alexei.hronos_torfa.plist
git commit -m "feat: расписание launchd 6–22 ежечасно"
```

---

## Task 2: Получатели — только alexeimvc@gmail.com

**Files:**
- Modify: `recipients.txt`

**Interfaces:**
- Consumes: `load_recipients()` из `send_email.py` (уже игнорирует строки с `#`).
- Produces: активный список получателей = `["alexeimvc@gmail.com"]`.

- [ ] **Step 1: Закомментировать остальные адреса**

Привести `recipients.txt` к виду:

```text
alexeimvc@gmail.com
# временно отключены до завершения отладки (Фаза 1):
# kazdobin@pskgold.ru
# bardadym@pskgold.ru
# paramonov@pskgold.ru
# turchinskas@pskgold.ru
```

- [ ] **Step 2: Проверить, что активный адрес ровно один**

Run: `$PY -c "import send_email as s; print(s.load_recipients())"`
Expected: `['alexeimvc@gmail.com']`

- [ ] **Step 3: Commit**

```bash
git add recipients.txt
git commit -m "chore: рассылка только на alexeimvc@gmail.com (остальные закомментированы)"
```

---

## Task 3: Модуль llm.py — SSH-мост к Claude

**Files:**
- Create: `llm.py`
- Test: `tests/test_llm.py`

**Interfaces:**
- Consumes: ключ `~/.ssh/eu_claude`, env `EU_CLAUDE_*` (опционально).
- Produces:
  - `available() -> bool`
  - `ask(prompt: str, timeout: int | None = None) -> str`

- [ ] **Step 1: Написать падающий тест**

`tests/test_llm.py`:

```python
import os, unittest
import llm

class TestLlmAvailable(unittest.TestCase):
    def test_available_false_when_key_missing(self):
        old = os.environ.get("EU_CLAUDE_KEY")
        os.environ["EU_CLAUDE_KEY"] = "/nonexistent/key"
        try:
            self.assertFalse(llm.available())
        finally:
            if old is None:
                del os.environ["EU_CLAUDE_KEY"]
            else:
                os.environ["EU_CLAUDE_KEY"] = old

    def test_ask_raises_without_key(self):
        os.environ["EU_CLAUDE_KEY"] = "/nonexistent/key"
        with self.assertRaises(RuntimeError):
            llm.ask("привет")

if __name__ == "__main__":
    unittest.main()
```

Примечание: `available()`/`ask()` читают `EU_CLAUDE_KEY` динамически (через `os.environ.get` внутри функции), чтобы тест мог переопределить путь.

- [ ] **Step 2: Запустить тест — убедиться, что падает**

Run: `$PY -m unittest tests.test_llm -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'llm'`.

- [ ] **Step 3: Реализовать llm.py**

`llm.py`:

```python
"""Мост к Claude Code на EU-сервере по SSH (тот же паттерн, что в проекте Geol).
Без API-ключа: аутентификация по SSH-ключу. Промпт уходит в stdin `claude -p`.
Зависимости — только стандартная библиотека.
"""
from __future__ import annotations
import os
import subprocess
from pathlib import Path


def _server() -> str:
    return os.environ.get("EU_CLAUDE_SERVER", "185.200.179.31")


def _user() -> str:
    return os.environ.get("EU_CLAUDE_USER", "root")


def _key() -> str:
    return os.environ.get("EU_CLAUDE_KEY", os.path.expanduser("~/.ssh/eu_claude"))


def _timeout() -> int:
    return int(os.environ.get("EU_CLAUDE_TIMEOUT", "300"))


def _ssh_opts() -> list[str]:
    return [
        "-i", _key(),
        "-o", "BatchMode=yes",
        "-o", "StrictHostKeyChecking=accept-new",
        "-o", "UserKnownHostsFile=/tmp/eu_claude_known_hosts",
        "-o", "ConnectTimeout=10",
    ]


def available() -> bool:
    return Path(_key()).exists()


def ask(prompt: str, timeout: int | None = None) -> str:
    """Запрос к Claude через `claude -p` по SSH. Промпт — в stdin (не аргументом:
    большой контекст упирался бы в ARG_MAX). Бросает RuntimeError при недоступности
    ключа или ненулевом коде возврата.
    """
    if not available():
        raise RuntimeError(f"SSH-ключ к Claude не найден: {_key()}")
    cmd = ["ssh", *_ssh_opts(), f"{_user()}@{_server()}", "claude -p"]
    res = subprocess.run(
        cmd, input=prompt, capture_output=True, text=True,
        timeout=timeout or _timeout(),
    )
    if res.returncode != 0:
        raise RuntimeError((res.stderr or res.stdout or "").strip()[:800])
    return res.stdout.strip()
```

- [ ] **Step 4: Запустить тест — убедиться, что проходит**

Run: `$PY -m unittest tests.test_llm -v`
Expected: PASS (2 теста).

- [ ] **Step 5: Живой smoke моста (ручная проверка)**

Run: `$PY -c "import llm; print('available', llm.available()); print(llm.ask('Ответь одним словом: тест')[:120])"`
Expected: `available True`, и одна строка ответа модели. (Если мост недоступен — зафиксировать, это не блокер: фолбэк в Task 6.)

- [ ] **Step 6: Commit**

```bash
git add llm.py tests/test_llm.py
git commit -m "feat: llm.py — SSH-мост к Claude (как в Geol)"
```

---

## Task 4: analytics.py — агрегаты и metrics

**Files:**
- Create: `analytics.py`
- Test: `tests/test_analytics.py`

**Interfaces:**
- Consumes: `logs/consolidated.csv` (формат как у `consolidate.py`), `logs/last_fetch.meta` (строка `date=YYYY-MM-DD`).
- Produces:
  - `compute(csv_path: Path) -> dict` — агрегаты с DataFrame'ами и скалярами. Ключи: `report_date:str`, `date_min`, `date_max`, `totals:dict`, `by_unit:DataFrame`, `by_date:DataFrame`, `by_hour:DataFrame`, `drivers:DataFrame`, `truck_mark:DataFrame`, `truck_inv:DataFrame`, `idle:DataFrame`, `units:list[str]`.
  - `to_metrics(aggr: dict) -> dict` — плоский JSON-совместимый словарь (см. ниже).
  - `dump_metrics(aggr: dict, path: Path) -> dict` — пишет `to_metrics(aggr)` в JSON, возвращает его.

`to_metrics` возвращает словарь со структурой:
```python
{
  "report_date": "2026-06-28",
  "period": ["2026-05-15", "2026-06-28"],
  "totals": {"volume": 1922928.0, "trips": 116459.0, "m3_per_trip": 16.5,
             "drivers": 65, "trucks": 33, "units": 5},
  "by_unit": [{"unit": "Сайлык", "trips": 37626.0, "volume": 751320.0,
               "m3_per_trip": 20.0, "share_pct": 39.1}, ...],
  "by_date": [{"date": "2026-06-27", "volume": 31870.0, "trips": ...}, ...],
  "truck_util": {"overall_pct": 99.8,
                 "by_mark": [{"mark": "HD465-7", "util_pct": 99.9}, ...]},
  "idle": [{"unit": "Дражный", "hours": 2456}, ...],
  "abc": {"A": 37, "B": 15, "C": 13}
}
```

- [ ] **Step 0: Создать пакет тестов**

Run: `mkdir -p tests && touch tests/__init__.py`
(нужно, чтобы `from tests.test_analytics import ...` работал в Task 5/6/7/8.)

- [ ] **Step 1: Написать падающий тест на форму metrics**

`tests/test_analytics.py`:

```python
import json, unittest
from pathlib import Path
import pandas as pd
import analytics

SCRATCH = Path("tests/_tmp"); SCRATCH.mkdir(parents=True, exist_ok=True)

def _make_csv(path: Path):
    rows = []
    # два дня, два подразделения, передел "Транспортировка торфов" + одна "простой"
    base = {"Передел": "Транспортировка торфов"}
    def r(date, unit, drv, vol, kuzov, inv, mark, t="08:00"):
        return {**base, "Дата. Факт": date, "Подразделение": unit,
                "Ф.И.О. водителя самосвала": drv, "Обьем работ, м3": vol,
                "Количство машин, шт": vol/20.0, "Обьем кузова,м3": kuzov,
                "Инв. № транспортировочной единицы": inv,
                "Марка транспортировочной единицы": mark, "Время": t}
    rows += [r("2026-06-27", "Эрел", "Иванов И.И.", 200, 20, "1001", "HD465-7"),
             r("2026-06-28", "Эрел", "Иванов И.И.", 100, 20, "1001", "HD465-7"),
             r("2026-06-28", "Обман", "Петров П.П.", 60, 12, "2001", "БелАЗ 7547")]
    idle = {"Передел": "простой", "Дата. Факт": "2026-06-28", "Подразделение": "Обман",
            "Ф.И.О. водителя самосвала": "", "Обьем работ, м3": "",
            "Количство машин, шт": "", "Обьем кузова,м3": "",
            "Инв. № транспортировочной единицы": "", "Марка транспортировочной единицы": "",
            "Время": "10:00"}
    rows.append(idle)
    pd.DataFrame(rows).to_csv(path, index=False)

class TestMetrics(unittest.TestCase):
    def setUp(self):
        self.csv = SCRATCH / "consolidated.csv"
        _make_csv(self.csv)

    def test_metrics_shape(self):
        aggr = analytics.compute(self.csv, report_date="2026-06-28")
        m = analytics.to_metrics(aggr)
        self.assertEqual(m["report_date"], "2026-06-28")
        self.assertEqual(m["totals"]["units"], 2)
        self.assertTrue(any(u["unit"] == "Эрел" for u in m["by_unit"]))
        # day-over-day: в by_date есть и 27-е, и 28-е
        dates = {d["date"] for d in m["by_date"]}
        self.assertIn("2026-06-27", dates)
        self.assertIn("2026-06-28", dates)
        # idle посчитан для Обман
        self.assertTrue(any(i["unit"] == "Обман" and i["hours"] >= 1 for i in m["idle"]))

    def test_dump_writes_json(self):
        aggr = analytics.compute(self.csv, report_date="2026-06-28")
        out = SCRATCH / "metrics.json"
        analytics.dump_metrics(aggr, out)
        data = json.loads(out.read_text(encoding="utf-8"))
        self.assertIn("by_unit", data)

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Запустить тест — убедиться, что падает**

Run: `$PY -m unittest tests.test_analytics -v`
Expected: FAIL — `No module named 'analytics'`.

- [ ] **Step 3: Реализовать analytics.py**

Перенести расчётный блок из `build_pdf.py:107-202` (агрегации) в `analytics.compute()`, добавить сигнатуру `report_date` и функции `to_metrics`/`dump_metrics`. `compute` не зависит от matplotlib.

`analytics.py`:

```python
"""Расчёт агрегатов по консолидированному реестру + плоский metrics-словарь.
Без matplotlib/сети. Используется build_xlsx.py (DataFrame'ы) и explain.py/email_html.py (metrics).
"""
from __future__ import annotations
import json
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd


def _abc_class(p: float) -> str:
    if p <= 80.0:
        return "A"
    if p <= 95.0:
        return "B"
    return "C"


def _to_hour(x):
    if pd.isna(x) or x in (None, "", " "):
        return None
    s = str(x).strip()
    if " " in s and ":" in s.split(" ", 1)[1]:
        s = s.split(" ", 1)[1]
    if ":" in s:
        try:
            return int(s.split(":", 1)[0])
        except ValueError:
            return None
    return None


def compute(csv_path: Path, report_date: str | None = None) -> dict:
    df = pd.read_csv(csv_path)
    if report_date is None:
        report_date = datetime.now().strftime("%Y-%m-%d")

    df["Дата. Факт"] = pd.to_datetime(df["Дата. Факт"], errors="coerce")
    for c in ("Количство машин, шт", "Обьем работ, м3", "Обьем кузова,м3"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Подразделение"] = df["Подразделение"].astype(str).str.strip()
    df["Час"] = df["Время"].apply(_to_hour)

    trans = df[df["Передел"].astype(str).str.strip() == "Транспортировка торфов"].copy()
    trans["Водитель"] = (trans["Ф.И.О. водителя самосвала"].astype(str).str.strip()
                         .replace({"nan": np.nan, "": np.nan}))
    trans = trans[trans["Водитель"].notna()]

    totals = {
        "volume": float(trans["Обьем работ, м3"].sum()),
        "trips": float(trans["Количство машин, шт"].sum()),
        "drivers": int(trans["Водитель"].nunique()),
        "trucks": int(trans["Инв. № транспортировочной единицы"].dropna().nunique()),
        "units": int(trans["Подразделение"].nunique()),
    }
    totals["m3_per_trip"] = round(totals["volume"] / totals["trips"], 1) if totals["trips"] else None

    by_unit = (trans.groupby("Подразделение")
               .agg(Рейсы=("Количство машин, шт", "sum"),
                    Объем_м3=("Обьем работ, м3", "sum"),
                    Водителей=("Водитель", "nunique"),
                    Самосвалов=("Инв. № транспортировочной единицы",
                                lambda s: s.dropna().nunique()),
                    Дней=("Дата. Факт", "nunique")).reset_index())
    by_unit["м3/рейс"] = by_unit["Объем_м3"] / by_unit["Рейсы"].replace(0, np.nan)
    by_unit["Доля_%"] = by_unit["Объем_м3"] / by_unit["Объем_м3"].sum() * 100
    by_unit = by_unit.sort_values("Объем_м3", ascending=False).reset_index(drop=True)

    by_date = (trans.groupby(trans["Дата. Факт"].dt.date)
               .agg(Рейсы=("Количство машин, шт", "sum"),
                    Объем_м3=("Обьем работ, м3", "sum"),
                    Записей=("Водитель", "count")).reset_index()
               .rename(columns={"Дата. Факт": "Дата"}))

    drv = (trans.groupby("Водитель")
           .agg(Рейсы=("Количство машин, шт", "sum"),
                Объем_м3=("Обьем работ, м3", "sum"),
                Дней=("Дата. Факт", "nunique"),
                Подразделение=("Подразделение",
                               lambda s: s.mode().iat[0] if len(s.mode()) else "—"))
           .reset_index())
    drv["м3/рейс"] = drv["Объем_м3"] / drv["Рейсы"].replace(0, np.nan)
    drv = drv.sort_values("Объем_м3", ascending=False).reset_index(drop=True)
    tv = drv["Объем_м3"].sum()
    drv["Доля_%"] = drv["Объем_м3"] / tv * 100 if tv else 0
    drv["Накопл_%"] = drv["Доля_%"].cumsum()
    drv["ABC"] = drv["Накопл_%"].apply(_abc_class)

    truck_mark = (trans.groupby("Марка транспортировочной единицы")
                  .agg(Кузов_м3_сред=("Обьем кузова,м3", "mean"),
                       Рейсы=("Количство машин, шт", "sum"),
                       Объем_м3=("Обьем работ, м3", "sum"),
                       Самосвалов=("Инв. № транспортировочной единицы",
                                   lambda s: s.dropna().nunique())).reset_index())
    truck_mark["м3/рейс"] = truck_mark["Объем_м3"] / truck_mark["Рейсы"].replace(0, np.nan)
    truck_mark["Использование_%"] = truck_mark["м3/рейс"] / truck_mark["Кузов_м3_сред"] * 100
    truck_mark = truck_mark.sort_values("Объем_м3", ascending=False).reset_index(drop=True)

    truck_inv = (trans.dropna(subset=["Инв. № транспортировочной единицы"])
                 .groupby(["Марка транспортировочной единицы",
                           "Инв. № транспортировочной единицы", "Подразделение"])
                 .agg(Кузов_м3=("Обьем кузова,м3", "mean"),
                      Рейсы=("Количство машин, шт", "sum"),
                      Объем_м3=("Обьем работ, м3", "sum"),
                      Дней=("Дата. Факт", "nunique")).reset_index())
    truck_inv["м3/рейс"] = truck_inv["Объем_м3"] / truck_inv["Рейсы"].replace(0, np.nan)
    truck_inv["Исп_%"] = truck_inv["м3/рейс"] / truck_inv["Кузов_м3"] * 100
    truck_inv = truck_inv.sort_values(["Подразделение", "Объем_м3"],
                                      ascending=[True, False]).reset_index(drop=True)

    by_hour = (trans.groupby("Час").agg(Рейсы=("Количство машин, шт", "sum"),
                                        Объем_м3=("Обьем работ, м3", "sum"))
               .reindex(range(24), fill_value=0).reset_index())

    idle = (df[df["Передел"].astype(str).str.strip() == "простой"]
            .groupby(["Подразделение"]).size().reset_index(name="Часов простоя"))

    return {
        "report_date": report_date,
        "date_min": df["Дата. Факт"].min(),
        "date_max": df["Дата. Факт"].max(),
        "totals": totals,
        "by_unit": by_unit, "by_date": by_date, "by_hour": by_hour,
        "drivers": drv, "truck_mark": truck_mark, "truck_inv": truck_inv,
        "idle": idle, "units": sorted(trans["Подразделение"].unique().tolist()),
        "n_rows": int(len(df)), "n_trans": int(len(trans)),
    }


def to_metrics(aggr: dict) -> dict:
    bu = aggr["by_unit"]
    bd = aggr["by_date"].sort_values("Дата")
    tm = aggr["truck_mark"]
    drv = aggr["drivers"]
    idle = aggr["idle"]

    def _d(x):
        return x.isoformat() if hasattr(x, "isoformat") else (str(x) if x is not None else None)

    overall_util = (tm["Объем_м3"].sum() /
                    (tm["Рейсы"] * tm["Кузов_м3_сред"]).sum() * 100
                    ) if tm["Рейсы"].sum() else None

    abc_counts = drv["ABC"].value_counts().to_dict()

    return {
        "report_date": aggr["report_date"],
        "period": [_d(aggr["date_min"]), _d(aggr["date_max"])],
        "n_rows": aggr["n_rows"], "n_trans": aggr["n_trans"],
        "totals": aggr["totals"],
        "by_unit": [{"unit": r["Подразделение"], "trips": round(float(r["Рейсы"]), 1),
                     "volume": round(float(r["Объем_м3"]), 1),
                     "m3_per_trip": round(float(r["м3/рейс"]), 1) if pd.notna(r["м3/рейс"]) else None,
                     "share_pct": round(float(r["Доля_%"]), 1)}
                    for _, r in bu.iterrows()],
        "by_date": [{"date": _d(r["Дата"]), "volume": round(float(r["Объем_м3"]), 1),
                     "trips": round(float(r["Рейсы"]), 1)}
                    for _, r in bd.iterrows()],
        "truck_util": {
            "overall_pct": round(float(overall_util), 1) if overall_util is not None else None,
            "by_mark": [{"mark": r["Марка транспортировочной единицы"],
                         "util_pct": round(float(r["Использование_%"]), 1) if pd.notna(r["Использование_%"]) else None}
                        for _, r in tm.iterrows()],
        },
        "idle": [{"unit": r["Подразделение"], "hours": int(r["Часов простоя"])}
                 for _, r in idle.iterrows()],
        "abc": {k: int(abc_counts.get(k, 0)) for k in ("A", "B", "C")},
    }


def dump_metrics(aggr: dict, path: Path) -> dict:
    m = to_metrics(aggr)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")
    return m
```

- [ ] **Step 4: Запустить тест — убедиться, что проходит**

Run: `$PY -m unittest tests.test_analytics -v`
Expected: PASS (2 теста).

- [ ] **Step 5: Smoke на реальных данных**

Run: `$PY -c "import analytics,json; a=analytics.compute('logs/consolidated.csv','2026-06-28'); print(json.dumps(analytics.to_metrics(a)['totals'], ensure_ascii=False))"`
Expected: словарь totals с volume ≈ 1.9e6, units 5.

- [ ] **Step 6: Commit**

```bash
git add analytics.py tests/test_analytics.py
git commit -m "feat: analytics.py — агрегаты и metrics из консолидированного реестра"
```

---

## Task 5: build_xlsx.py — листы аналитики в книгу реестра

**Files:**
- Create: `build_xlsx.py`
- Test: `tests/test_build_xlsx.py`

**Interfaces:**
- Consumes: `analytics.compute()`; путь к существующей книге реестра из `output/Хронометраж_транспортировки_торфов_<дата>.xlsx` (её делает `consolidate.py`).
- Produces: `main() -> int`; добавляет в книгу листы `Свод_подразделения`, `Свод_даты`, `ABC_водители`, `Парк_марки`, `Парк_инв`. Печатает путь книги.

- [ ] **Step 1: Написать падающий тест**

`tests/test_build_xlsx.py`:

```python
import unittest
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
import build_xlsx, analytics
from tests.test_analytics import _make_csv, SCRATCH

class TestBuildXlsx(unittest.TestCase):
    def test_sheets_appended(self):
        csv = SCRATCH / "consolidated.csv"; _make_csv(csv)
        book = SCRATCH / "reestr.xlsx"
        pd.DataFrame({"Подразделение": ["Эрел"]}).to_excel(book, index=False, sheet_name="Сводный_Реестр")
        aggr = analytics.compute(csv, "2026-06-28")
        build_xlsx.write_sheets(aggr, book)
        wb = load_workbook(book)
        for s in ("Свод_подразделения", "Свод_даты", "ABC_водители", "Парк_марки", "Парк_инв"):
            self.assertIn(s, wb.sheetnames)

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Запустить тест — убедиться, что падает**

Run: `$PY -m unittest tests.test_build_xlsx -v`
Expected: FAIL — `No module named 'build_xlsx'`.

- [ ] **Step 3: Реализовать build_xlsx.py**

`build_xlsx.py`:

```python
"""Дописывает к книге консолидированного реестра листы аналитики (замена графикам PDF).
Запуск: python3 build_xlsx.py  — берёт последнюю книгу из output/ и logs/consolidated.csv.
"""
from __future__ import annotations
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd
import analytics

BASE = Path(__file__).resolve().parent
OUT = BASE / "output"
LOGS = BASE / "logs"
CSV = LOGS / "consolidated.csv"

SHEETS = {
    "Свод_подразделения": "by_unit",
    "Свод_даты": "by_date",
    "ABC_водители": "drivers",
    "Парк_марки": "truck_mark",
    "Парк_инв": "truck_inv",
}


def write_sheets(aggr: dict, book_path: Path) -> None:
    with pd.ExcelWriter(book_path, engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as xw:
        for sheet, key in SHEETS.items():
            aggr[key].to_excel(xw, sheet_name=sheet, index=False)


def _report_date() -> str:
    rd = datetime.now().strftime("%Y-%m-%d")
    meta = LOGS / "last_fetch.meta"
    if meta.exists():
        for line in meta.read_text(encoding="utf-8").splitlines():
            if line.startswith("date="):
                rd = line.split("=", 1)[1].strip()
    return rd


def _latest_book() -> Path | None:
    books = sorted(OUT.glob("Хронометраж_транспортировки_торфов_*.xlsx"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return books[0] if books else None


def main() -> int:
    if not CSV.exists():
        print(f"[ERR] нет {CSV}"); return 2
    book = _latest_book()
    if not book:
        print("[ERR] нет книги реестра в output/"); return 3
    aggr = analytics.compute(CSV, _report_date())
    write_sheets(aggr, book)
    # метрики для explain.py / email_html / алертов
    analytics.dump_metrics(aggr, BASE / "state" / "last_metrics.json")
    print(f"[OK] листы аналитики добавлены: {book}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Запустить тест — убедиться, что проходит**

Run: `$PY -m unittest tests.test_build_xlsx -v`
Expected: PASS.

- [ ] **Step 5: Smoke на реальной книге (копия, чтобы не портить рабочую)**

Run:
```bash
cp "output/Хронометраж_транспортировки_торфов_2026-06-28.xlsx" /tmp/reestr_test.xlsx
$PY -c "import analytics,build_xlsx; from pathlib import Path; a=analytics.compute('logs/consolidated.csv','2026-06-28'); build_xlsx.write_sheets(a, Path('/tmp/reestr_test.xlsx')); print('ok')"
$PY -c "from openpyxl import load_workbook; print(load_workbook('/tmp/reestr_test.xlsx').sheetnames)"
```
Expected: список листов содержит `Свод_подразделения`, `Свод_даты`, `ABC_водители`, `Парк_марки`, `Парк_инв`.

- [ ] **Step 6: Commit**

```bash
git add build_xlsx.py tests/test_build_xlsx.py
git commit -m "feat: build_xlsx.py — листы аналитики в книге реестра"
```

---

## Task 6: explain.py — записка и аномалии

**Files:**
- Create: `explain.py`
- Test: `tests/test_explain.py`

**Interfaces:**
- Consumes: `state/last_metrics.json` (формат `analytics.to_metrics`), `llm.available()/ask()`.
- Produces:
  - `build_brief(metrics: dict) -> str`
  - `detect_anomalies(metrics: dict, thresholds: dict | None = None) -> list[dict]` → `[{"type": str, "text": str}]`
  - `rule_based_note(metrics: dict) -> str`
  - `generate_note(metrics: dict) -> str` (мост, при ошибке → `rule_based_note`)
  - `main() -> int` — читает `state/last_metrics.json`, пишет `output/Пояснительная_записка_<дата>.md`, `state/last_note.txt`, `state/last_alerts.json`.
  - Константы порогов: `MIN_BODY_UTIL=95.0`, `MAX_IDLE_H=1.0`, `MAX_VOLUME_DROP=30.0` (переопределяемы через env `HR_MIN_BODY_UTIL` и т.п.).

- [ ] **Step 1: Написать падающий тест (правиловая записка + аномалии)**

`tests/test_explain.py`:

```python
import unittest
import explain

METRICS = {
    "report_date": "2026-06-28",
    "period": ["2026-05-15", "2026-06-28"],
    "totals": {"volume": 30000.0, "trips": 1500.0, "m3_per_trip": 20.0,
               "drivers": 10, "trucks": 8, "units": 2},
    "by_unit": [{"unit": "Эрел", "trips": 1000.0, "volume": 20000.0,
                 "m3_per_trip": 20.0, "share_pct": 66.7},
                {"unit": "Обман", "trips": 500.0, "volume": 10000.0,
                 "m3_per_trip": 20.0, "share_pct": 33.3}],
    "by_date": [{"date": "2026-06-27", "volume": 50000.0, "trips": 2500.0},
                {"date": "2026-06-28", "volume": 30000.0, "trips": 1500.0}],
    "truck_util": {"overall_pct": 90.0,
                   "by_mark": [{"mark": "БелАЗ 7547", "util_pct": 90.0}]},
    "idle": [{"unit": "Обман", "hours": 3}],
    "abc": {"A": 4, "B": 3, "C": 3},
}

class TestExplain(unittest.TestCase):
    def test_rule_based_note_mentions_key_facts(self):
        note = explain.rule_based_note(METRICS)
        self.assertIn("2026-06-28", note)
        self.assertIn("Эрел", note)          # лидер по объёму
        self.assertTrue(len(note) > 100)

    def test_anomalies_detected(self):
        al = explain.detect_anomalies(METRICS)
        types = {a["type"] for a in al}
        self.assertIn("body_util", types)    # 90% < 95%
        self.assertIn("idle", types)         # 3ч > 1ч
        self.assertIn("volume_drop", types)  # 30000 vs 50000 = -40% > 30%

    def test_no_anomalies_when_healthy(self):
        healthy = {**METRICS,
                   "truck_util": {"overall_pct": 99.0, "by_mark": []},
                   "idle": [],
                   "by_date": [{"date": "2026-06-27", "volume": 30000.0, "trips": 1500.0},
                               {"date": "2026-06-28", "volume": 30000.0, "trips": 1500.0}]}
        self.assertEqual(explain.detect_anomalies(healthy), [])

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Запустить тест — убедиться, что падает**

Run: `$PY -m unittest tests.test_explain -v`
Expected: FAIL — `No module named 'explain'`.

- [ ] **Step 3: Реализовать explain.py**

`explain.py`:

```python
"""Пояснительная записка (через мост llm.py, фолбэк — по правилам) и аномалии KPI.
Запуск: python3 explain.py — читает state/last_metrics.json.
"""
from __future__ import annotations
import json, os, sys
from pathlib import Path
import llm

BASE = Path(__file__).resolve().parent
STATE = BASE / "state"
OUT = BASE / "output"

MIN_BODY_UTIL = float(os.environ.get("HR_MIN_BODY_UTIL", "95"))
MAX_IDLE_H = float(os.environ.get("HR_MAX_IDLE_H", "1.0"))
MAX_VOLUME_DROP = float(os.environ.get("HR_MAX_VOLUME_DROP", "30"))


def _fmt_int(x) -> str:
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except (ValueError, TypeError):
        return str(x)


def build_brief(m: dict) -> str:
    t = m["totals"]
    lines = [
        f"Отчётная дата: {m['report_date']}. Период данных: {m['period'][0]}…{m['period'][1]}.",
        f"Итоги за период: объём {_fmt_int(t['volume'])} м³, рейсов {_fmt_int(t['trips'])}, "
        f"м³/рейс {t.get('m3_per_trip')}, водителей {t['drivers']}, самосвалов {t['trucks']}, "
        f"подразделений {t['units']}.",
        "По подразделениям (объём, доля%):",
    ]
    lines += [f"  - {u['unit']}: {_fmt_int(u['volume'])} м³ ({u['share_pct']}%), "
              f"м³/рейс {u['m3_per_trip']}" for u in m["by_unit"]]
    bd = m["by_date"]
    if len(bd) >= 2:
        lines.append(f"Динамика: {bd[-2]['date']} {_fmt_int(bd[-2]['volume'])} м³ → "
                     f"{bd[-1]['date']} {_fmt_int(bd[-1]['volume'])} м³.")
    lines.append(f"Загрузка кузова (общая): {m['truck_util'].get('overall_pct')}%.")
    if m["idle"]:
        lines.append("Простои (часов): " +
                     ", ".join(f"{i['unit']} {i['hours']}" for i in m["idle"]))
    lines.append(f"ABC водителей: A={m['abc']['A']}, B={m['abc']['B']}, C={m['abc']['C']}.")
    return "\n".join(lines)


def detect_anomalies(m: dict, thresholds: dict | None = None) -> list[dict]:
    th = {"util": MIN_BODY_UTIL, "idle": MAX_IDLE_H, "drop": MAX_VOLUME_DROP}
    if thresholds:
        th.update(thresholds)
    out: list[dict] = []
    util = m.get("truck_util", {}).get("overall_pct")
    if util is not None and util < th["util"]:
        out.append({"type": "body_util",
                    "text": f"Загрузка кузова {util}% ниже нормы {th['util']:.0f}%."})
    for i in m.get("idle", []):
        if i["hours"] > th["idle"]:
            out.append({"type": "idle",
                        "text": f"Простои в «{i['unit']}»: {i['hours']} ч "
                                f"(порог {th['idle']:.0f} ч/смену)."})
    bd = m.get("by_date", [])
    if len(bd) >= 2 and bd[-2]["volume"] > 0:
        drop = (bd[-2]["volume"] - bd[-1]["volume"]) / bd[-2]["volume"] * 100
        if drop > th["drop"]:
            out.append({"type": "volume_drop",
                        "text": f"Объём упал на {drop:.0f}% к {bd[-2]['date']} "
                                f"({_fmt_int(bd[-2]['volume'])} → {_fmt_int(bd[-1]['volume'])} м³)."})
    return out


def rule_based_note(m: dict) -> str:
    t = m["totals"]
    units = sorted(m["by_unit"], key=lambda u: u["volume"], reverse=True)
    leader = units[0] if units else None
    al = detect_anomalies(m)
    parts = [
        f"# Пояснительная записка · {m['report_date']}",
        f"За период {m['period'][0]}…{m['period'][1]} перевезено "
        f"{_fmt_int(t['volume'])} м³ торфа ({_fmt_int(t['trips'])} рейсов, "
        f"в среднем {t.get('m3_per_trip')} м³/рейс). Задействовано {t['drivers']} водителей "
        f"и {t['trucks']} самосвалов в {t['units']} подразделениях.",
    ]
    if leader:
        parts.append(f"Наибольший вклад — «{leader['unit']}»: {_fmt_int(leader['volume'])} м³ "
                     f"({leader['share_pct']}% объёма).")
    bd = m["by_date"]
    if len(bd) >= 2:
        d = "снизился" if bd[-1]["volume"] < bd[-2]["volume"] else "вырос"
        parts.append(f"К предыдущему дню объём {d}: "
                     f"{_fmt_int(bd[-2]['volume'])} → {_fmt_int(bd[-1]['volume'])} м³.")
    if al:
        parts.append("Внимание: " + " ".join(a["text"] for a in al))
    else:
        parts.append("Существенных отклонений KPI не зафиксировано.")
    return "\n\n".join(parts)


def generate_note(m: dict) -> str:
    if not llm.available():
        return rule_based_note(m)
    prompt = (
        "Ты — аналитик горнодобывающего предприятия. По данным ниже напиши "
        "пояснительную записку к ежедневному отчёту по транспортировке торфов: "
        "деловой стиль, 3–5 абзацев, что произошло за день, ключевые цифры и "
        "динамика к прошлому дню, узкие места (низкая загрузка кузова, простои), "
        "краткие рекомендации. НИЧЕГО не выдумывай сверх приведённых чисел. "
        "Без markdown-таблиц, только текст и абзацы.\n\nДАННЫЕ:\n" + build_brief(m)
    )
    try:
        text = llm.ask(prompt).strip()
        return text or rule_based_note(m)
    except Exception:
        return rule_based_note(m)


def main() -> int:
    mp = STATE / "last_metrics.json"
    if not mp.exists():
        print(f"[ERR] нет {mp} (запусти build_xlsx/analytics)"); return 2
    m = json.loads(mp.read_text(encoding="utf-8"))
    note = generate_note(m)
    alerts = detect_anomalies(m)

    OUT.mkdir(parents=True, exist_ok=True); STATE.mkdir(parents=True, exist_ok=True)
    (OUT / f"Пояснительная_записка_{m['report_date']}.md").write_text(note, encoding="utf-8")
    (STATE / "last_note.txt").write_text(note, encoding="utf-8")
    (STATE / "last_alerts.json").write_text(
        json.dumps(alerts, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] записка ({len(note)} симв.), аномалий: {len(alerts)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Запустить тест — убедиться, что проходит**

Run: `$PY -m unittest tests.test_explain -v`
Expected: PASS (3 теста).

- [ ] **Step 5: Smoke — записка на реальных метриках (через мост, с фолбэком)**

Run:
```bash
$PY -c "import analytics; from pathlib import Path; a=analytics.compute('logs/consolidated.csv','2026-06-28'); analytics.dump_metrics(a, Path('state/last_metrics.json'))"
$PY explain.py
sed -n '1,40p' "output/Пояснительная_записка_2026-06-28.md"
cat state/last_alerts.json
```
Expected: осмысленный текст записки; `state/last_alerts.json` — список (возможно пустой).

- [ ] **Step 6: Commit**

```bash
git add explain.py tests/test_explain.py
git commit -m "feat: explain.py — пояснительная записка (мост+фолбэк) и аномалии KPI"
```

---

## Task 7: email_html.py — HTML-тело письма

**Files:**
- Create: `email_html.py`
- Test: `tests/test_email_html.py`

**Interfaces:**
- Consumes: `metrics` (формат `analytics.to_metrics`), `note: str`, `alerts: list[dict]`.
- Produces:
  - `build_html(metrics: dict, note: str, alerts: list[dict]) -> str`
  - `build_text(metrics: dict, note: str, alerts: list[dict]) -> str`

- [ ] **Step 1: Написать падающий тест**

`tests/test_email_html.py`:

```python
import unittest
import email_html
from tests.test_explain import METRICS

class TestEmailHtml(unittest.TestCase):
    def test_html_has_key_blocks(self):
        html = email_html.build_html(METRICS, "Записка прозой.", [{"type": "idle", "text": "Простои в «Обман»: 3 ч."}])
        self.assertIn("<table", html)
        self.assertIn("Эрел", html)               # таблица подразделений
        self.assertIn("Записка прозой.", html)     # записка
        self.assertIn("Простои в «Обман»", html)   # блок алертов
        self.assertIn("2026-06-28", html)

    def test_text_fallback_plain(self):
        txt = email_html.build_text(METRICS, "Записка.", [])
        self.assertNotIn("<table", txt)
        self.assertIn("Эрел", txt)

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Запустить тест — убедиться, что падает**

Run: `$PY -m unittest tests.test_email_html -v`
Expected: FAIL — `No module named 'email_html'`.

- [ ] **Step 3: Реализовать email_html.py**

`email_html.py`:

```python
"""Рендер HTML- и текстового тела письма из metrics + записки + аномалий.
Только инлайн-стили (почтовые клиенты режут <style>). Зависимости — stdlib.
"""
from __future__ import annotations
from html import escape


def _fmt_int(x) -> str:
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except (ValueError, TypeError):
        return str(x)


def _kpi_tiles(m: dict) -> str:
    t = m["totals"]
    tiles = [("Объём, м³", _fmt_int(t["volume"])), ("Рейсов", _fmt_int(t["trips"])),
             ("м³/рейс", t.get("m3_per_trip")), ("Водителей", t["drivers"]),
             ("Самосвалов", t["trucks"]), ("Подразделений", t["units"])]
    cells = "".join(
        f'<td style="padding:8px 14px;border:1px solid #d0d7de;">'
        f'<div style="color:#57606a;font-size:12px;">{escape(str(k))}</div>'
        f'<div style="font-size:18px;font-weight:bold;">{escape(str(v))}</div></td>'
        for k, v in tiles)
    return f'<table style="border-collapse:collapse;margin:8px 0;"><tr>{cells}</tr></table>'


def _unit_table(m: dict) -> str:
    head = ("<tr>" + "".join(
        f'<th style="padding:6px 10px;background:#305496;color:#fff;'
        f'border:1px solid #d0d7de;text-align:left;">{h}</th>'
        for h in ("Подразделение", "Рейсы", "Объём, м³", "м³/рейс", "Доля, %")) + "</tr>")
    rows = ""
    for i, u in enumerate(m["by_unit"]):
        bg = "#f6f8fa" if i % 2 else "#ffffff"
        rows += ("<tr>" + "".join(
            f'<td style="padding:6px 10px;border:1px solid #d0d7de;background:{bg};">{c}</td>'
            for c in (escape(u["unit"]), _fmt_int(u["trips"]), _fmt_int(u["volume"]),
                      u["m3_per_trip"], u["share_pct"])) + "</tr>")
    return f'<table style="border-collapse:collapse;margin:8px 0;">{head}{rows}</table>'


def _trend_table(m: dict, days: int = 7) -> str:
    bd = m["by_date"][-days:]
    head = ('<tr><th style="padding:6px 10px;background:#305496;color:#fff;'
            'border:1px solid #d0d7de;">Дата</th>'
            '<th style="padding:6px 10px;background:#305496;color:#fff;'
            'border:1px solid #d0d7de;">Объём, м³</th>'
            '<th style="padding:6px 10px;background:#305496;color:#fff;'
            'border:1px solid #d0d7de;">Рейсы</th></tr>')
    rows = "".join(
        f'<tr><td style="padding:6px 10px;border:1px solid #d0d7de;">{escape(str(d["date"]))}</td>'
        f'<td style="padding:6px 10px;border:1px solid #d0d7de;">{_fmt_int(d["volume"])}</td>'
        f'<td style="padding:6px 10px;border:1px solid #d0d7de;">{_fmt_int(d["trips"])}</td></tr>'
        for d in bd)
    return f'<table style="border-collapse:collapse;margin:8px 0;">{head}{rows}</table>'


def _alerts_block(alerts: list[dict]) -> str:
    if not alerts:
        return ('<p style="color:#1a7f37;">Существенных отклонений KPI не зафиксировано.</p>')
    items = "".join(f"<li>{escape(a['text'])}</li>" for a in alerts)
    return ('<div style="border-left:4px solid #d1242f;background:#fff8f8;padding:8px 14px;margin:8px 0;">'
            '<b style="color:#d1242f;">Внимание — отклонения KPI:</b>'
            f'<ul style="margin:6px 0;">{items}</ul></div>')


def build_html(m: dict, note: str, alerts: list[dict]) -> str:
    note_html = "".join(f"<p>{escape(p)}</p>" for p in note.split("\n\n") if p.strip())
    return (
        '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;color:#1f2328;max-width:720px;">'
        f'<h2 style="margin:0 0 4px;">Хронометраж транспортировки торфов</h2>'
        f'<div style="color:#57606a;">Отчётная дата: {escape(m["report_date"])} · '
        f'период {escape(str(m["period"][0]))}…{escape(str(m["period"][1]))}</div>'
        f'<h3>Ключевые показатели</h3>{_kpi_tiles(m)}'
        f'{_alerts_block(alerts)}'
        f'<h3>Пояснительная записка</h3>{note_html}'
        f'<h3>По подразделениям</h3>{_unit_table(m)}'
        f'<h3>Динамика (последние дни)</h3>{_trend_table(m)}'
        '<p style="color:#8c959f;font-size:12px;margin-top:16px;">'
        'Детализация — во вложении (Excel: реестр + листы аналитики). '
        'Автоматическая рассылка хронометража.</p>'
        '</div>'
    )


def build_text(m: dict, note: str, alerts: list[dict]) -> str:
    t = m["totals"]
    lines = [f"Хронометраж торфов — отчёт за {m['report_date']}",
             f"Период: {m['period'][0]}…{m['period'][1]}", "",
             f"Объём {_fmt_int(t['volume'])} м³ · рейсов {_fmt_int(t['trips'])} · "
             f"м³/рейс {t.get('m3_per_trip')} · водителей {t['drivers']} · "
             f"самосвалов {t['trucks']}", ""]
    if alerts:
        lines.append("ВНИМАНИЕ — отклонения KPI:")
        lines += [f"  - {a['text']}" for a in alerts]
    else:
        lines.append("Существенных отклонений KPI не зафиксировано.")
    lines += ["", "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА:", note, "",
              "По подразделениям:"]
    for u in m["by_unit"]:
        lines.append(f"  {u['unit']}: {_fmt_int(u['volume'])} м³ ({u['share_pct']}%), "
                     f"рейсов {_fmt_int(u['trips'])}")
    lines += ["", "Детализация — во вложении (Excel)."]
    return "\n".join(lines)
```

- [ ] **Step 4: Запустить тест — убедиться, что проходит**

Run: `$PY -m unittest tests.test_email_html -v`
Expected: PASS (2 теста).

- [ ] **Step 5: Визуальный smoke HTML**

Run:
```bash
$PY -c "import json,email_html; m=json.load(open('state/last_metrics.json')); note=open('state/last_note.txt').read(); al=json.load(open('state/last_alerts.json')); open('/tmp/preview.html','w').write(email_html.build_html(m,note,al))"
open /tmp/preview.html
```
Expected: в браузере читаемое письмо: KPI, блок аномалий/«без отклонений», записка, таблицы.

- [ ] **Step 6: Commit**

```bash
git add email_html.py tests/test_email_html.py
git commit -m "feat: email_html.py — HTML/текст тело письма из metrics"
```

---

## Task 8: send_email.py — HTML-письмо + Excel, без PDF

**Files:**
- Modify: `send_email.py:59-99` (`find_latest_outputs`, `build_message`), `:121-172` (`main`)
- Test: `tests/test_send_email.py`

**Interfaces:**
- Consumes: `email_html.build_html/build_text`, `analytics`/`state/last_*`.
- Produces:
  - `build_message(env, recipients, xlsx, html_body, text_body) -> EmailMessage` (новая сигнатура: без pdf, с html/text).
  - `find_latest_xlsx() -> Path | None`.
  - CLI `main()` собирает тело из `state/last_metrics.json` + `state/last_note.txt` + `state/last_alerts.json`.

- [ ] **Step 1: Написать падающий тест на multipart/alternative + вложение**

`tests/test_send_email.py`:

```python
import unittest
from pathlib import Path
import send_email
from tests.test_explain import METRICS
import email_html

class TestBuildMessage(unittest.TestCase):
    def test_html_and_attachment(self):
        env = {"YANDEX_LOGIN": "shirobokov@pskgold.ru"}
        xlsx = Path("/tmp/_se_test.xlsx"); xlsx.write_bytes(b"PK\x03\x04test")
        html = email_html.build_html(METRICS, "Записка.", [])
        text = email_html.build_text(METRICS, "Записка.", [])
        msg = send_email.build_message(env, ["alexeimvc@gmail.com"], xlsx, html, text)
        self.assertEqual(msg["To"], "alexeimvc@gmail.com")
        # есть html-часть
        body = msg.get_body(preferencelist=("html",))
        self.assertIsNotNone(body)
        self.assertIn("Эрел", body.get_content())
        # есть вложение xlsx
        names = [p.get_filename() for p in msg.iter_attachments()]
        self.assertIn("_se_test.xlsx", names)

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Запустить тест — убедиться, что падает**

Run: `$PY -m unittest tests.test_send_email -v`
Expected: FAIL — `build_message() ... positional arguments` (старая сигнатура).

- [ ] **Step 3: Обновить send_email.py**

Заменить `find_latest_outputs` на `find_latest_xlsx`, переписать `build_message` под html/text, обновить `main`.

`build_message` (заменить тело функции `:77-99`):

```python
def build_message(env: dict, recipients: list[str], xlsx: Path | None,
                  html_body: str, text_body: str) -> EmailMessage:
    sender = env["YANDEX_LOGIN"]
    msg = EmailMessage()
    today = datetime.now().strftime("%Y-%m-%d")
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"Хронометраж торфов — отчёт за {today}"
    msg.set_content(text_body)              # text/plain (фолбэк)
    msg.add_alternative(html_body, subtype="html")  # text/html
    if xlsx and xlsx.exists():
        attach_file(msg, xlsx)
    return msg
```

`find_latest_xlsx` (заменить `find_latest_outputs` `:59-65`):

```python
def find_latest_xlsx() -> Path | None:
    xlsx = sorted(OUT.glob("Хронометраж_транспортировки_торфов_*.xlsx"),
                  key=lambda p: p.stat().st_mtime, reverse=True)
    return xlsx[0] if xlsx else None
```

`main()` — собрать тело из state и отправить (заменить `:121-172`):

```python
def main() -> int:
    import json
    import email_html
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="только проверить SMTP-логин")
    ap.add_argument("--xlsx", type=Path, help="путь к xlsx (по умолчанию — последний)")
    args = ap.parse_args()

    env = load_env()
    if "@" not in env.get("YANDEX_LOGIN", ""):
        log(f"[ERR] YANDEX_LOGIN некорректный: {env.get('YANDEX_LOGIN')!r}"); return 2

    if args.check:
        host = env.get("SMTP_HOST", "smtp.yandex.ru"); port = int(env.get("SMTP_PORT", "465"))
        try:
            with smtplib.SMTP_SSL(host, port, timeout=20) as s:
                s.login(env["YANDEX_LOGIN"], env["YANDEX_APP_PASSWORD"])
            log("[OK] SMTP-логин принят"); return 0
        except Exception as e:
            log(f"[ERR] SMTP-проверка не удалась: {e!r}"); return 3

    rcpts = load_recipients()
    if not rcpts:
        log(f"[ERR] список получателей пуст: {RCPT_FILE}"); return 4

    xlsx = args.xlsx or find_latest_xlsx()
    if not xlsx:
        log(f"[ERR] не найден xlsx в {OUT}"); return 5

    state = BASE / "state"
    metrics = json.loads((state / "last_metrics.json").read_text(encoding="utf-8"))
    note = (state / "last_note.txt").read_text(encoding="utf-8") if (state / "last_note.txt").exists() else ""
    alerts_p = state / "last_alerts.json"
    alerts = json.loads(alerts_p.read_text(encoding="utf-8")) if alerts_p.exists() else []
    html = email_html.build_html(metrics, note, alerts)
    text = email_html.build_text(metrics, note, alerts)

    msg = build_message(env, rcpts, xlsx, html, text)
    try:
        send(env, msg)
    except Exception as e:
        log(f"[ERR] отправка не удалась: {e!r}"); return 6
    log(f"[OK] отправлено {len(rcpts)} получателям: {', '.join(rcpts)}  |  {xlsx.name}")
    return 0
```

Также добавить `BASE = Path(__file__).resolve().parent` (уже есть) и убедиться, что `import json` доступен (импортирован вверху или локально в main, как выше).

- [ ] **Step 4: Запустить тест — убедиться, что проходит**

Run: `$PY -m unittest tests.test_send_email -v`
Expected: PASS.

- [ ] **Step 5: Реальная отправка себе (контролируемо)**

Предусловие: выполнены smoke из Task 6 (есть `state/last_metrics.json`, `state/last_note.txt`, `state/last_alerts.json`) и есть свежий xlsx в `output/`.
Run: `$PY send_email.py`
Expected: `[OK] отправлено 1 получателям: alexeimvc@gmail.com`. Проверить почту `alexeimvc@gmail.com`: HTML-письмо с запиской/таблицами и вложением Excel; PDF нет.

- [ ] **Step 6: Commit**

```bash
git add send_email.py tests/test_send_email.py
git commit -m "feat: send_email.py — HTML-письмо + вложение Excel, без PDF"
```

---

## Task 9: run_daily.py — новый поток, email-алерты, без Telegram

**Files:**
- Modify: `run_daily.py` (полностью переписать оркестрацию; удалить `notify_bot`, `tg_request`, `_read_admin_chat`)
- Test: `tests/test_run_daily.py`

**Interfaces:**
- Consumes: `build_xlsx.main` (через subprocess), `explain.main`, `send_email`, `state/last_alerts.json`.
- Produces:
  - `alert(env, subject: str, body: str, key: str) -> None` — отправляет email-алерт на `alexeimvc@gmail.com` с анти-спамом по `key` (файл `state/alerts_sent.json`, ключ `"<дата>|<key>"`).
  - `main() -> int` — поток `fetch → consolidate → build_xlsx → explain → (алерты) → send_email`.

- [ ] **Step 1: Написать падающий тест на анти-спам alert()**

`tests/test_run_daily.py`:

```python
import json, unittest
from pathlib import Path
from unittest import mock
import run_daily

class TestAlertDedup(unittest.TestCase):
    def setUp(self):
        run_daily.STATE.mkdir(parents=True, exist_ok=True)
        self.sent = run_daily.STATE / "alerts_sent.json"
        if self.sent.exists():
            self.sent.unlink()

    def test_alert_sent_once_per_key_per_day(self):
        calls = []
        with mock.patch.object(run_daily, "_send_alert_email",
                               side_effect=lambda *a, **k: calls.append(a)):
            run_daily.alert({}, "subj", "body", key="idle:Обман")
            run_daily.alert({}, "subj", "body", key="idle:Обман")  # дубль
        self.assertEqual(len(calls), 1)
        data = json.loads(self.sent.read_text(encoding="utf-8"))
        self.assertTrue(any("idle:Обман" in k for k in data))

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Запустить тест — убедиться, что падает**

Run: `$PY -m unittest tests.test_run_daily -v`
Expected: FAIL — `module 'run_daily' has no attribute 'alert'` / `_send_alert_email`.

- [ ] **Step 3: Переписать run_daily.py**

`run_daily.py`:

```python
"""Оркестратор: fetch → consolidate → build_xlsx → explain → алерты → send_email (авто).
Запускается launchd-ом каждый час в xx:10 (6–22). Доставка — только email.
"""
from __future__ import annotations
import json, smtplib, ssl, subprocess, sys, time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

BASE = Path(__file__).resolve().parent
LOGS = BASE / "logs"
STATE = BASE / "state"
PY = sys.executable
ALERT_TO = "alexeimvc@gmail.com"


def load_env() -> dict:
    p = BASE / ".env"; out = {}
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


def _send_alert_email(env: dict, subject: str, body: str) -> None:
    msg = EmailMessage()
    msg["From"] = env.get("YANDEX_LOGIN", ALERT_TO)
    msg["To"] = ALERT_TO
    msg["Subject"] = f"[hronos_torfa] {subject}"
    msg.set_content(body)
    host = env.get("SMTP_HOST", "smtp.yandex.ru"); port = int(env.get("SMTP_PORT", "465"))
    login = env.get("SMTP_LOGIN") or env.get("YANDEX_LOGIN")
    password = env.get("SMTP_PASSWORD") or env.get("YANDEX_APP_PASSWORD")
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
        s.login(login, password)
        s.send_message(msg)


def alert(env: dict, subject: str, body: str, key: str) -> None:
    """Email-алерт с анти-спамом: один и тот же key за дату отправляется один раз."""
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


def main() -> int:
    STATE.mkdir(parents=True, exist_ok=True); LOGS.mkdir(parents=True, exist_ok=True)
    env = load_env()
    t0 = time.time()
    log("=" * 60); log("RUN_DAILY START")

    rc, _ = run("fetch_kronos_torf.py")
    if rc != 0:
        log("[WARN] fetch с ошибкой, продолжаю с локальными вложениями")

    rd = _report_date()
    if rd != datetime.now().strftime("%Y-%m-%d"):
        alert(env, "нет свежих данных",
              f"Дата отчёта {rd} не сегодняшняя — возможно, не пришли новые письма.",
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

    # авто-рассылка
    rc, _ = run("send_email.py")
    if rc != 0:
        alert(env, "сбой отправки письма", f"send_email.py exit={rc}", key=f"fail_send:{rd}")

    _write_state({"ok": True, "report_date": rd,
                  "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                  "duration_sec": round(time.time() - t0, 1)})
    log(f"RUN_DAILY DONE за {round(time.time()-t0,1)}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

Примечание: `state/last_metrics.json` пишет `build_xlsx.main()` (см. Task 5 Step 3) — поэтому `explain.py` запускается строго после `build_xlsx.py`.

- [ ] **Step 4: Запустить тест — убедиться, что проходит**

Run: `$PY -m unittest tests.test_run_daily -v`
Expected: PASS.

- [ ] **Step 5: Прогнать весь набор тестов**

Run: `$PY -m unittest discover -s tests -v`
Expected: все тесты PASS.

- [ ] **Step 6: Интеграционный прогон end-to-end**

Run: `$PY run_daily.py`
Expected: в `logs/run_daily.log` — шаги fetch/consolidate/build_xlsx/explain/send_email; на почте `alexeimvc@gmail.com` — HTML-письмо с Excel-вложением; PDF не создаётся.

- [ ] **Step 7: Commit**

```bash
git add run_daily.py build_xlsx.py tests/test_run_daily.py
git commit -m "feat: run_daily.py — HTML-рассылка, email-алерты с анти-спамом, без Telegram"
```

---

## Task 10: Вывод из эксплуатации PDF/бота + README

**Files:**
- Modify: `README.md`
- Modify: `scripts/com.alexei.hronos_torfa_bot.plist` (выгрузить из launchd)

**Interfaces:**
- Consumes: ничего.
- Produces: актуальная документация; бот не запускается.

- [ ] **Step 1: Выгрузить launchd-задачу бота**

Run:
```bash
launchctl unload ~/Library/LaunchAgents/com.alexei.hronos_torfa_bot.plist 2>/dev/null
launchctl list | grep hronos_torfa_bot || echo "бот выгружен"
```
Expected: `бот выгружен`.

- [ ] **Step 2: Обновить README.md**

Привести README к новому состоянию: формат доставки — HTML-письмо + Excel (реестр + листы аналитики) + пояснительная записка; PDF и Telegram-бот удалены из пайплайна; расписание 6–22; рассылка с Mac (Фаза 1); SMTP-зависимость и план переноса на сервер (Фаза 2). Команды бота из README убрать. Указать запуск `python3 run_daily.py` и набор модулей (`fetch_kronos_torf`, `consolidate`, `build_xlsx`, `analytics`, `explain`, `email_html`, `send_email`, `llm`).

- [ ] **Step 3: Пометить устаревшие модули**

В начало `build_pdf.py` и `bot.py` добавить строку-комментарий после докстринга:
```python
# DEPRECATED (Фаза 1): не используется в пайплайне. Оставлено для истории; удалить после Фазы 2.
```

- [ ] **Step 4: Commit**

```bash
git add README.md build_pdf.py bot.py
git commit -m "docs: README под новый формат; bot.py/build_pdf.py помечены deprecated"
```

---

## Self-Review

**Spec coverage:**
- Расписание 6–22 → Task 1. ✔
- Получатель только alexeimvc → Task 2. ✔
- Мост llm.py → Task 3. ✔
- metrics + дамп → Task 4 (+ запись в Task 9/build_xlsx). ✔
- Excel листы аналитики → Task 5. ✔
- Записка (мост+фолбэк) + аномалии → Task 6. ✔
- HTML-письмо записка+таблицы → Task 7. ✔
- Авто-отправка HTML+Excel, без PDF → Task 8. ✔
- Поток + email-алерты + анти-спам + без Telegram → Task 9. ✔
- PDF/бот из эксплуатации, README → Task 10. ✔
- Фаза 2 (сервер) — вне этого плана (отложена по спеке). ✔

**Открытые мелочи для исполнителя:**
- `tests/__init__.py` создать пустым в Task 4 (чтобы `from tests.test_analytics import ...` работал). Добавить в Step 1 Task 4.
- Запись `state/last_metrics.json` происходит в `build_xlsx.main()` (Task 9 Step 3) — explain зависит от этого порядка.
