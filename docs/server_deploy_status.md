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
