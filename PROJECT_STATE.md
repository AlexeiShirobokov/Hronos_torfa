# PROJECT_STATE — Hronos_torfa

Снимок состояния проекта для передачи следующему агенту.
Дата фиксации: 2026-05-20.

---

## 1. Цель проекта

Полностью автоматизировать ежедневный/ежечасный цикл «хронометраж транспортировки торфов»:

1. Раз в час (xx:10 с 9:00 до 19:00 по локальному TZ macOS) скачивать Excel-вложения из IMAP-папки `Hronos_torfa` ящика Яндекс 360 за последнюю доступную дату.
2. Объединять «Реестры» 4 подразделений (Эрел, Дражный, Обман, Сайлык) в один консолидированный xlsx.
3. Строить PDF-аналитику (KPI, сравнение подразделений, динамика по датам и часам, ABC-анализ водителей общий и по подразделениям, разбор парка по грузоподъёмности самосвалов, выводы, рекомендации, чек-лист начальника участка).
4. Уведомлять администратора в Telegram (`@alexeids_bot`) о готовом отчёте с кнопками «Отправить»/«Отмена».
5. По подтверждению — рассылать xlsx+pdf на адреса из `recipients.txt` через SMTP Яндекса.
6. Управлять всем через Telegram-бот командами `/run_now`, `/status`, `/list_emails`, `/add_email`, `/remove_email`, `/pdf`, `/xlsx`, `/send_now`, `/confirm`, `/cancel`.

Репозиторий: `git@github.com:AlexeiShirobokov/Hronos_torfa.git` (создан, push сделает пользователь — у агента нет SSH-доступа к github из песочницы).

---

## 2. Что уже сделано

| Шаг | Статус | Артефакт |
|---|---|---|
| Конфиг `.env` с IMAP-настройками | DONE (пользователь заполнил) | `/Users/alexei/Claude_v1/.env` |
| Структура папок (`input/mail_attachments`, `output`, `logs`, `state`) | DONE | те же |
| IMAP-загрузка Excel-вложений | DONE | `fetch_kronos_torf.py` (запускался пользователем локально — у агента нет сети к yandex) |
| Скачано 25 вложений за 2026-05-18 | DONE | `input/mail_attachments/` |
| Консолидированный реестр (1 лист «Сводный_Реестр», 848 строк, 23 колонки) | DONE | `output/Хронометраж_транспортировки_торфов_2026-05-18.xlsx` |
| PDF-аналитика (14 страниц, KPI, подразделения, даты, ABC, грузоподъёмность, рекомендации, чек-лист) | DONE | `output/Аналитика_хронометраж_торфов_2026-05-18.pdf` |
| Скрипты пайплайна | DONE | `consolidate.py`, `build_pdf.py`, `run_daily.py` |
| SMTP-рассылка по `recipients.txt` | DONE | `send_email.py` |
| Telegram-бот polling без зависимостей | DONE | `bot.py` |
| launchd: пайплайн в xx:10 9–19 + бот KeepAlive | DONE | `scripts/com.alexei.hronos_torfa.plist`, `scripts/com.alexei.hronos_torfa_bot.plist`, `install_launchd.sh`, `uninstall_launchd.sh` |
| `.gitignore`, `.env.example`, `README.md`, `recipients.txt` | DONE | те же |
| Git init/коммит локально в песочнице | ОТЛОЖЕНО | mount только для чтения для `.git/index.lock` (Operation not permitted) — пользователь сделает локально |
| Push на GitHub | ОТЛОЖЕНО | у агента нет сети к github.com (прокси 403) — пользователь сделает локально |
| Установка launchd-задач | ОТЛОЖЕНО | пользователь запустит `bash scripts/install_launchd.sh` |
| Регистрация админ-чата (`/start` боту) | ОТЛОЖЕНО | пользователь сделает один раз |

---

## 3. Файлы и их назначение

Корень проекта — `/Users/alexei/Claude_v1/`.

### Конфигурация и данные

- `.env` — секреты IMAP/SMTP/Telegram. **В git не пушится.**
- `.env.example` — шаблон без секретов.
- `recipients.txt` — список email-получателей (по одному в строке). Сейчас: `alexeimvc@gmail.com`.
- `AGENTS.md` — старые проектные заметки (содержит пароль; в `.gitignore`).

### Папки

- `input/mail_attachments/` — скачанные вложения IMAP.
- `output/` — финальный xlsx и pdf (имя содержит дату).
- `logs/` — `fetch_*.log`, `last_fetch.meta`, `last_consolidate.json`, `last_pdf.json`, `mail.log`, `run_daily.log`, `bot.log`, `consolidated.csv`.
- `state/` — `admin_chat_id.txt` (записывается при `/start`), `last_run.json` (состояние последнего запуска).

### Шаги пайплайна (можно запускать по отдельности)

| Файл | Что делает | Вход | Выход |
|---|---|---|---|
| `fetch_kronos_torf.py` | IMAP → SSL 993, ищет папку `Hronos_torfa`, читает INTERNALDATE батчами по 200, скачивает Excel-вложения только из писем за самую позднюю дату | `.env`, IMAP | `input/mail_attachments/*.xlsx`, `logs/fetch_*.log`, `logs/last_fetch.meta` |
| `consolidate.py` | Группирует вложения по базовому имени (`__N` — версия), берёт последнюю версию каждой группы, читает лист «Реестр», объединяет в один лист | `input/mail_attachments/*` | `output/Хронометраж_транспортировки_торфов_<date>.xlsx` (1 лист «Сводный_Реестр», 23 колонки без служебных), `logs/consolidated.csv`, `logs/last_consolidate.json` |
| `build_pdf.py` | matplotlib PdfPages, шрифт DejaVu Sans (кириллица). 14 страниц A4. | `logs/consolidated.csv` | `output/Аналитика_хронометраж_торфов_<date>.pdf`, `logs/last_pdf.json` |
| `send_email.py` | smtplib SMTP_SSL 465. Логин/пароль из `.env` (`SMTP_LOGIN`/`SMTP_PASSWORD` — опционально, иначе `YANDEX_LOGIN`/`YANDEX_APP_PASSWORD`). `--check` — только проверка логина. | `recipients.txt`, последние файлы из `output/` | письма с двумя вложениями (xlsx, pdf), `logs/mail.log` |
| `run_daily.py` | Оркестратор: fetch → consolidate → build_pdf → уведомление боту через Telegram API (inline-кнопки «✅ Отправить» / «✖️ Отмена»). **Сам не отправляет письма.** | всё выше | `state/last_run.json`, `logs/run_daily.log` |
| `bot.py` | Polling-бот (urllib, без зависимостей). Команды разрешены только admin (см. §8). Обрабатывает callback кнопок. | `.env` (`BOT_TOKEN` или `api_token`), `recipients.txt`, `output/` | сообщения в Telegram, мутации `recipients.txt`, запуск дочерних процессов |

### launchd

- `scripts/com.alexei.hronos_torfa.plist` — пайплайн, `StartCalendarInterval` 11 точек (Hour=9..19, Minute=10), `RunAtLoad=false`.
- `scripts/com.alexei.hronos_torfa_bot.plist` — бот, `RunAtLoad=true`, `KeepAlive=true`.
- `scripts/install_launchd.sh` — копирует plist в `~/Library/LaunchAgents/`, делает `bootout`+`bootstrap`+`enable` через `launchctl`.
- `scripts/uninstall_launchd.sh` — отключает и удаляет.

### Документация

- `README.md` — установка, команды бота, описание.
- `PROJECT_STATE.md` — этот файл.

---

## 4. Команды запуска и проверки

### Локальный one-shot

```bash
cd /Users/alexei/Claude_v1

python3 fetch_kronos_torf.py        # IMAP → input/mail_attachments
python3 consolidate.py              # → output/<...>.xlsx + logs/consolidated.csv
python3 build_pdf.py                # → output/<...>.pdf
python3 send_email.py --check       # проверка SMTP-логина
python3 send_email.py               # отправка последних файлов (никогда не из расписания напрямую)
python3 run_daily.py                # полный цикл + уведомление боту
```

### Установка расписания и бота

```bash
bash scripts/install_launchd.sh

launchctl list | grep hronos        # увидеть оба job
tail -f logs/launchd_bot.out.log    # лог бота
tail -f logs/launchd_pipeline.out.log
```

### Управление через Telegram (`@alexeids_bot`)

```
/start            — регистрация чата как админ (один раз)
/run_now          — синхронный fetch → xlsx → pdf, ответ с хвостом лога
/status           — последний запуск + хвост logs/run_daily.log
/list_emails
/add_email user@domain
/remove_email user@domain
/pdf              — прислать последний pdf файлом
/xlsx             — прислать последний xlsx файлом
/send_now         — выполнить рассылку прямо сейчас
/confirm          — то же, что кнопка «✅ Отправить»
/cancel           — снять флаг pending_send
```

### Деплой в GitHub (делает пользователь — нет SSH из песочницы)

```bash
cd /Users/alexei/Claude_v1
git init
git branch -M main
git add .
git commit -m "Initial: hronos_torfa pipeline + bot + launchd"
git remote add origin git@github.com:AlexeiShirobokov/Hronos_torfa.git
git push -u origin main
```

---

## 5. Переменные `.env` (без секретов)

```env
# Яндекс 360 — IMAP/SMTP
YANDEX_LOGIN=name@example.ru
YANDEX_APP_PASSWORD=__пароль_приложения_Яндекс_для_IMAP_SMTP__
IMAP_HOST=imap.yandex.ru
IMAP_PORT=993
IMAP_FOLDER=Hronos_torfa
SMTP_HOST=smtp.yandex.ru
SMTP_PORT=465
# Если для SMTP отдельный пароль — указать здесь, иначе берётся YANDEX_*
SMTP_LOGIN=
SMTP_PASSWORD=

# Telegram-бот @alexeids_bot
BOT_TOKEN=__telegram_bot_token__
# Список разрешённых tg user_id через запятую. Если пусто — админом становится тот, кто первым написал /start
BOT_ADMINS=
# Заполняется автоматически из state/admin_chat_id.txt, можно прописать руками
BOT_ADMIN_CHAT_ID=
```

В текущем `.env` поле бота называется `api_token=` — код понимает оба имени (`BOT_TOKEN` или `api_token`).
Рекомендуется привести к `BOT_TOKEN=` для единообразия.

`SMTP_HOST`/`SMTP_PORT` в текущем `.env` пока отсутствуют — добавятся при необходимости (значения по умолчанию подставляются в коде).

---

## 6. Ошибки, которые уже встречались, и как исправлены

| Симптом | Причина | Исправление |
|---|---|---|
| `Temporary failure in name resolution` для `imap.yandex.ru` | Песочница агента блокирует исходящие к Яндексу (прокси 403) | IMAP-вызов выполняет пользователь локально |
| `Cannot convert <NA> to Excel` при `ws.append(...)` | `pd.NA`/`pd.NaT` openpyxl не принимает | Введён `_clean(v)` в `consolidate.py` (None для NaN/NaT) |
| `ValueError: invalid literal for int(): '1900-01-01 00'` в `to_hour` | После CSV-roundtrip время приходит как `1900-01-01 HH:MM:SS` | `to_hour` отрезает дату до пробела и парсит остаток |
| `TypeError: type str doesn't define __round__` в `fmt_int` | Иногда инв.№ — строка | `fmt_int` обёрнут в try/except, на ошибке возвращает `str(x)` |
| Логин Yandex IMAP не принят как `shirobokov2pskgold.ru` | Пропущен `@` в `.env` | Пользователь правил вручную → `shirobokov@pskgold.ru` |
| `Реестр_Обман` всего 21 строка | У файла Обман `max_row≈36000` из-за форматирования, реальных данных мало | Это норма; данные верны |
| ABC «100%» использования кузова у всех марок | В реестре объём = `кузов × рейсы` (не фактический замер) | Зафиксировано в выводах PDF, дана рекомендация ввести фактический замер |
| Git init упал с `Operation not permitted` | Mount-права в песочнице на `.git/index.lock` | Пользователь сделает `git init` локально |

---

## 7. Что осталось сделать

Конкретные действия пользователя (агент сам не может — нет сети/прав):

1. **Привести `.env`** к шаблону `.env.example` (переименовать `api_token` → `BOT_TOKEN`, добавить `SMTP_HOST=smtp.yandex.ru`, `SMTP_PORT=465`, при желании заполнить `BOT_ADMINS=<свой_tg_user_id>`).
2. **Сгенерировать новые секреты** (так как старые мелькали в переписке):
   - Новый пароль приложения Yandex (id.yandex.ru → Безопасность → Пароли приложений → Почта/IMAP+SMTP).
   - Новый Telegram bot token у @BotFather (/revoke текущего).
3. **Установить зависимости**:
   ```bash
   python3 -m pip install --user pandas openpyxl matplotlib
   ```
4. **Локальный smoke-тест** (по одному шагу):
   ```bash
   python3 fetch_kronos_torf.py
   python3 consolidate.py
   python3 build_pdf.py
   python3 send_email.py --check
   ```
5. **Git push** в `git@github.com:AlexeiShirobokov/Hronos_torfa.git` (см. §4).
6. **launchd**: `bash scripts/install_launchd.sh`.
7. **Бот**: написать `/start` в Telegram `@alexeids_bot` — создастся `state/admin_chat_id.txt`.
8. **Полный E2E**: `/run_now` через бота, дождаться сообщения с кнопками, нажать ✅, проверить, что письмо пришло на `alexeimvc@gmail.com`.

Опциональные улучшения, которых пока нет (можно сделать вторым агентом по запросу):

- Дедупликация писем по `Message-ID` (сейчас `fetch_kronos_torf.py` скачивает заново каждый запуск).
- Хранение состояния «какие письма уже обработаны», чтобы не качать одно и то же.
- Опция «динамика по нескольким дням» в PDF (сейчас часовая динамика — только за отчётную дату).
- Графики по сменам (1/2), по экскаваторщикам, по блокам.
- Email-шаблон с инлайн-сводкой в тело письма (сейчас простой plain text + два вложения).
- Healthcheck-эндпойнт бота (`/ping`) и алёрт админу при провале пайплайна.
- Кэш `output/` за прошлые даты с автоочисткой старше 30 дней.
- Поддержка нескольких IMAP-папок (несколько объектов).

---

## 8. Правила для следующего агента

### Жёсткие (из проектных инструкций пользователя)

1. **Отвечать на русском.**
2. **Работать только в `/Users/alexei/Claude_v1/` и подпапках.** Файлы вне проекта не трогать без явного разрешения.
3. **Входные данные — `input/`, итоги — `output/`, промежуточные — `logs/`.**
4. **Перед сложной задачей — короткий план** и список затрагиваемых файлов.
5. **Почта:**
   - Пароли не выводить в чат.
   - Учётные данные брать только из `.env`.
   - Не удалять письма из IMAP.
   - **Не отправлять письма без явного подтверждения** (через кнопку в боте, `/send_now` или `/confirm`).
   - Вложения класть только в `input/mail_attachments/`.
6. **Excel:**
   - Не перезаписывать исходные файлы.
   - Итоговый xlsx — с датой/временем в имени.
   - После сборки проверять: число строк, листы, имена колонок, пустые критичные поля.
   - В ответе — краткий отчёт (что взято, что обработано, куда сохранено).
7. **На ошибку — не идти вслепую**, а описать причину и предложить исправление.
8. **Финальный ответ** обязан содержать: что сделано, где лежит результат, какие были проблемы, что нужно проверить человеку.

### Технические соглашения проекта

- Имена секретов в `.env`: `YANDEX_LOGIN`, `YANDEX_APP_PASSWORD`, `BOT_TOKEN` (или legacy `api_token`), `BOT_ADMINS`, `BOT_ADMIN_CHAT_ID`, `SMTP_*`. Не вводить новые имена без причины.
- Шрифт PDF — DejaVu Sans (есть на macOS/Linux, поддерживает кириллицу). Не менять без необходимости.
- Имена выходных файлов сохранять как есть: `Хронометраж_транспортировки_торфов_YYYY-MM-DD.xlsx`, `Аналитика_хронометраж_торфов_YYYY-MM-DD.pdf`.
- Опечатки колонок в источниках («Количство машин, шт», «Обьем работ, м3», «Обьем кузова,м3») **сохранять**, чтобы сверка с исходниками не ломалась. В отображаемых таблицах PDF — корректные термины.
- Listing IMAP-папок: имя «Hronos_torfa» (латиницей, как в Яндексе). Не путать с «Kronos_torf»/«Hronos_torfa».
- Отчёт строится за дату из `logs/last_fetch.meta` (`date=`), не из `datetime.now()`.
- В git **не пушить**: `.env`, `AGENTS.md`, `input/`, `output/`, `logs/`, `state/`, `__pycache__/`. `.gitignore` уже это исключает — не ослаблять.
- При смене расписания править `StartCalendarInterval` в `scripts/com.alexei.hronos_torfa.plist` и перезагружать через `install_launchd.sh`.
- Не выкладывать чувствительные значения в чат (пароли приложения Yandex, Telegram-токены). Их вычитывать только из `.env`.

### Поведение бота — инвариант

- Бот **никогда** не отправляет email без явного действия пользователя (нажатия кнопки или команды `/send_now`/`/confirm`).
- `run_daily.py` **никогда** сам не вызывает `send_email.py` — он только готовит файлы и пишет уведомление.
- Эти два инварианта вытекают из проектных правил и должны соблюдаться при любых изменениях.

### Если что-то неясно

Уточнять у пользователя через `AskUserQuestion` перед изменениями, особенно по:
- расписанию,
- списку получателей,
- логике отправки писем,
- структуре PDF/Excel,
- любым изменениям в IMAP/SMTP/Telegram-настройках.

---

_Конец PROJECT_STATE.md_
