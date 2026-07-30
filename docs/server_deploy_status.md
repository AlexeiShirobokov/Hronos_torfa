# Server Deploy Status

Дата: 2026-07-03

Проект размещен как staging на сервере:

```text
geol-deploy:/home/deploy/hronos_torfa
```

Что сделано:

- создан каталог `/home/deploy/hronos_torfa`;
- перенесен код проекта, `.env`, `recipients.txt`, шаблон сводной и C# binaries;
- созданы рабочие папки `logs/`, `state/`, `output/`, `input/mail_attachments/`;
- установлен Python venv: `/home/deploy/hronos_torfa/.venv`;
- установлены Python-зависимости `pandas`, `openpyxl`, `pyxlsb`;
- установлен .NET SDK 9 в пользовательскую папку `/home/deploy/.dotnet`;
- выполнена контрольная сборка без отправки.

Контрольная сборка:

```text
cd /home/deploy/hronos_torfa
.venv/bin/python run_hronos_peski_autosend.py --no-fetch --no-send
```

Результат проверки:

- Python-консолидация прошла: 44163 строки;
- C# PivotBuilder прошел;
- `Таблица1` создана на диапазоне `A1:W44164`;
- источник сводной: `Сводный_Реестр!Таблица1`;
- свежесть источников OK по Дражный, Обман, Сайлык, Талынья, Эрел;
- объемы OK;
- итоговый файл создан:

```text
/home/deploy/hronos_torfa/output/Хронометраж транспортировки торфов и песков.xlsx
```

## Переключение на сервер

Дата переключения: 2026-07-03

SMTP-блокер снят:

```text
.venv/bin/python send_email.py --check
[OK] SMTP-логин принят
```

Серверный cron включен. С 2026-07-05 отправка идет даже при неполной свежести исходников, но предупреждение по отсутствующим свежим подразделениям остается в логах. С 2026-07-10 минута отправки сдвинута на :30, окно 06:30-20:30 Asia/Ust-Nera:

```cron
MAILTO=""
SHELL=/bin/bash
# Hronos_torfa: 06:30-20:30 Asia/Ust-Nera; server timezone is MSK (+03), so hours are 23,0-13 MSK.
30 0-13,23 * * * cd /home/deploy/hronos_torfa && TZ=Asia/Ust-Nera .venv/bin/python run_hronos_peski_autosend.py >> logs/cron_hronos_peski.log 2>&1
```

Mac launchd отключен, чтобы не было двойных отправок:

- `com.alexei.hronos_torfa` — disabled и выгружен;
- `com.alexei.hronos_torfa_bot` — disabled.

Теперь рабочая авторассылка должна выполняться с сервера.

## Миграция на новый сервер 185.20.226.122

Дата: 2026-07-25

Боевой пайплайн перенесён с `geol-deploy` (37.140.195.93, REG.RU) на новый сервер:

```text
hronos-deploy:/home/deploy/hronos_torfa      # ssh-хост hronos-deploy = 185.20.226.122, root; ключ ~/.ssh/hronos_deploy
```

Причина/контекст: работа переводится на сервер `185.20.226.122`. Старый `geol-deploy`
остаётся жив ради **отдельного** проекта `promyvka_peskov` — его cron не трогали.

Что сделано:

- создан пользователь `deploy`; проект и `/home/deploy/.dotnet` (SDK 9.0.315) перенесены
  напрямую сервер→сервер через `rsync` (без `logs/`, `output/`, C# `obj/`); `.venv`
  склонирован как есть (одинаковая ОС Ubuntu 20.04 / Python 3.8.10 / тот же путь);
- перенесены `input/mail_attachments/` (~15М) — нужны для консолидации;
- **исходящий SMTP на новом сервере не заблокирован** (порты 465/587/25 открыты),
  `send_email.py --check` → `[OK]`;
- контрольная сборка `--no-fetch --no-send`: `DONE`, 66902 строки, `Таблица1 A1:U66903`,
  свежесть/объёмы OK; C# PivotBuilder собрал финальную книгу;
- контрольная реальная отправка `--to shirobokov@pskgold.ru` → `[OK] отправлено 1 получателям`.

Переключение cron (чтобы не было двойной рассылки):

- на **старом** сервере строки hronos и dispatcher закомментированы префиксом
  `#MIGRATED-2026-07-25` (бэкап `~/crontab_backup_pre_migration_hronos.txt`);
  `promyvka_peskov` оставлена активной;
- на **новом** сервере установлен crontab пользователя `deploy` с теми же задачами:

```cron
MAILTO=""
SHELL=/bin/bash
# Hronos_torfa: 06:30-20:30 Asia/Ust-Nera; server timezone is MSK (+03), so hours are 23,0-13 MSK.
30 0-13,23 * * * cd /home/deploy/hronos_torfa && TZ=Asia/Ust-Nera .venv/bin/python run_hronos_peski_autosend.py >> logs/cron_hronos_peski.log 2>&1
# Dispatcher feedback: 09:35-19:35 Asia/Ust-Nera (MSK hours 2-12)
35 2-12 * * * cd /home/deploy/hronos_torfa && TZ=Asia/Ust-Nera .venv/bin/python dispatcher_feedback.py >> logs/dispatcher_feedback.log 2>&1
```

