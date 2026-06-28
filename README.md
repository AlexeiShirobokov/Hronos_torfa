# Hronos_torfa

Автоматизация ежедневного хронометража транспортировки торфов: получает письма
из Яндекс 360, собирает консолидированный Excel (реестр + листы аналитики),
формирует пояснительную записку и **HTML-письмо** с ключевыми показателями,
шлёт алерты об аномалиях KPI и рассылает отчёт по email.

> **Фаза 1 (сейчас):** всё работает на Mac, рассылка — с Mac (исходящий SMTP
> хостинга geol-deploy заблокирован, открывается по тикету в REG.RU).
> **Фаза 2 (позже):** перенос пайплайна на сервер geol-deploy (venv + cron),
> после открытия SMTP. См. `docs/superpowers/specs/` и `docs/superpowers/plans/`.

## Что внутри

| Файл | Назначение |
|---|---|
| `fetch_kronos_torf.py` | IMAP-загрузка Excel-вложений за последнюю дату из папки `Hronos_torfa` |
| `consolidate.py` | Сборка `Сводный_Реестр` из последних версий вложений |
| `analytics.py` | Агрегаты по реестру + плоский `metrics` (`state/last_metrics.json`) |
| `build_xlsx.py` | Дописывает в книгу реестра листы аналитики; пишет `metrics` |
| `explain.py` | Пояснительная записка (по правилам) + аномалии KPI |
| `email_html.py` | HTML/текст тело письма из `metrics` + записки + аномалий |
| `send_email.py` | SMTP-рассылка HTML-письма + вложение Excel по `recipients.txt` |
| `run_daily.py` | Оркестратор: fetch → consolidate → build_xlsx → explain → алерты → send_email |
| `recipients.txt` | Получатели email (по одному в строке; `#` — выключен) |
| `scripts/com.alexei.hronos_torfa.plist` | launchd-задача пайплайна (xx:10, 6–22) |

Листы Excel: `Сводный_Реестр`, `Свод_подразделения`, `Свод_даты`, `ABC_водители`,
`Парк_марки`, `Парк_инв`.

> `build_pdf.py` и `bot.py` помечены DEPRECATED и не используются в пайплайне
> (PDF и Telegram-бот убраны из Фазы 1).

## Поведение

* В будни и выходные **в 10 минут каждого часа с 6:00 до 22:00** (TZ macOS)
  `run_daily.py` собирает отчёт и **автоматически отправляет письмо** активным
  получателям из `recipients.txt` (сейчас — только `alexeimvc@gmail.com`).
* **Алерты по email** (на `alexeimvc@gmail.com`):
  * сбои шагов (`consolidate` / `build_xlsx` / отправка), нет свежих данных;
  * аномалии KPI: загрузка кузова < 95%, простои > 1 ч/смену, падение объёма
    к прошлому дню > 30%.
  Анти-спам: один и тот же алерт за дату уходит один раз (`state/alerts_sent.json`).
* Пороги аномалий переопределяются через env: `HR_MIN_BODY_UTIL`, `HR_MAX_IDLE_H`,
  `HR_MAX_VOLUME_DROP`.

## Первый запуск

```bash
# 1. учётные данные
cp .env.example .env
$EDITOR .env

# 2. установить расписание (xx:10 в 6–22)
bash scripts/install_launchd.sh
```

## Запуск вручную

```bash
PY=/Library/Frameworks/Python.framework/Versions/3.12/bin/python3
$PY run_daily.py            # весь пайплайн + рассылка
$PY send_email.py --check   # проверить SMTP-логин
```

## Тесты

```bash
/Library/Frameworks/Python.framework/Versions/3.12/bin/python3 -m unittest discover -s tests
```

## Безопасность

`.env` и пользовательские данные исключены из git (`.gitignore`).
Никогда не пушьте `.env` и содержимое `input/`, `output/`, `logs/`.

## Зависимости

`python3` (стандартная библиотека) + `pandas`, `openpyxl`.
```bash
python3 -m pip install --user pandas openpyxl
```
