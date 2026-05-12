# Google Sheets Summary Sync

Безопасная синхронизация двух Google Sheets-источников ("Наша сетка" и
"Яндекс сетка") в итоговую таблицу "Сводная по месяцам".

Главная цель — **невозможно случайно испортить боевые таблицы**.

---

## 1. Что делает проект

Раз в N секунд (или один раз при `--once`):

1. Открывает SUMMARY-таблицу.
2. Читает лист `Settings`: список месяцев и имена вкладок в источниках.
3. Для каждого месяца:
   - читает соответствующую вкладку "Наша сетка";
   - читает соответствующую вкладку "Яндекс сетка" (опционально);
   - считает по менеджерам: офферты, ИП, ТОО, договор есть, акцепт/оплата,
     процент акцепта, метки `nib_sale`/`nib`/`0`/пусто/другое, красные;
   - находит в листе `Сводная - Месяц` два блока — `НАША СЕТКА (...)` и
     `ЯНДЕКС СЕТКА (...)`;
   - записывает результаты **только в диапазон A:M** соответствующего блока;
   - пишет timestamp в `N1`.

Никакие другие колонки, листы, диапазоны, формулы — не трогаются.

---

## 2. Какие таблицы нужны

Три Google Sheets:

| Назначение | Описание |
|---|---|
| OUR_GRID_ID | "Наша сетка": исходник по месяцам |
| YANDEX_GRID_ID | "Яндекс сетка": исходник по месяцам |
| SUMMARY_SPREADSHEET_ID | "Сводная по месяцам": куда пишутся итоги |

В SUMMARY должен быть лист `Settings` со столбцами:
```
A1: Листы: Наша Сетка   B1: Листы: Яндекс Сетка
A2: Май 2026            B2: Май 2026
A3: Апрель 2026         B3: Апрель 2026
...
```

Для каждого месяца из Settings должен существовать лист
`Сводная - <Месяц>` с двумя блоками `НАША СЕТКА (...)` и `ЯНДЕКС СЕТКА (...)`,
строка ниже title — заголовки колонок:

```
Менеджеры | Офферты всего | ИП | ТОО | Договор есть | Акцепт/Оплата
| Акцепт % | Метка nib_sale | Метка nib | Метка 0 | Пусто | Другое | Красные
```

---

## 3. Как создать копии таблиц для теста

Открой каждую боевую таблицу:

1. Файл → Создать копию.
2. Назови, например, `[ТЕСТ] Сводная по месяцам`.
3. Открой копию и из URL скопируй ID (часть между `/d/` и `/edit`).
4. Дай созданному service account доступ **Editor** ко всем трём копиям
   (по `client_email` из JSON-файла service account).

> ⚠️ Боевые ID жёстко зашиты в коде. В `ENVIRONMENT=test` запуск с любым из
> них немедленно падает с понятной ошибкой.

---

## 4. Локальный запуск на Mac

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cp .env.example .env
# открой .env и поставь ID копий и креды
```

### 4.1. Первый dry-run

В `.env` оставь:
```
ENVIRONMENT=test
RUN_MODE=once
DRY_RUN=true
SAFE_MODE=true
ALLOW_PRODUCTION_WRITE=false
```

Запусти:

```bash
python main.py --once --dry-run
```

В консоли увидишь:
- баннер с настройками;
- проверку названий таблиц;
- список месяцев из Settings;
- для каждого блока: что было бы записано (`[DRY] would WRITE ...`);
- никаких реальных записей.

Можно проверить отдельный месяц:

```bash
python main.py --once --dry-run --month "Май 2026"
```

### 4.2. Реальная запись на копиях

Если dry-run выглядит хорошо, в `.env` поставь:
```
DRY_RUN=false
```
Всё остальное — без изменений (всё ещё на копиях, ENVIRONMENT=test).

```bash
python main.py --once
```

Теперь скрипт реально пишет в SUMMARY-копию. Открой её и проверь:
- блоки А:М заполнены;
- N1 содержит `Обновлено: дата время`;
- никакие ручные пометки за пределами A:M не пострадали;
- в папке `backups/` появились JSON-снимки старого состояния.

### 4.3. Loop-режим

```bash
python main.py --loop
```

Цикл с паузой `LOOP_SLEEP_SEC` между итерациями. Прервать — Ctrl+C.

---

## 5. Переход на production

Когда всё проверено на копиях:

1. В `.env` подставь **боевые** ID:
   ```
   ENVIRONMENT=production
   OUR_GRID_ID=1OgF4xLUqwSHs2S2NPXCJfsVgLM4V9W5c3yQq5lDoS-o
   SUMMARY_SPREADSHEET_ID=18nfkpxiPG6xB7uLcCpwV8Qwmx7VJ3ii7uHUzfUh4L2Y
   YANDEX_GRID_ID=1Qf4vPXqfpa83NkCrsTa0OHqv5PTKUB89_lzXIok1NWQ
   ```

2. Сначала сделай dry-run на production:
   ```
   DRY_RUN=true
   ALLOW_PRODUCTION_WRITE=false
   ```
   ```bash
   python main.py --once --dry-run
   ```

3. Если всё ок:
   ```
   DRY_RUN=false
   ALLOW_PRODUCTION_WRITE=true
   SAFE_MODE=true
   ```
   ```bash
   python main.py --once
   ```

> ⚠️ Если `ENVIRONMENT=production` и `ALLOW_PRODUCTION_WRITE=false`, скрипт
> автоматически форсирует `DRY_RUN=true` — это последний барьер.

---

## 6. Все переменные окружения

| Переменная | Значения | Что делает |
|---|---|---|
| `ENVIRONMENT` | `test` / `production` | В `test` запрещены боевые ID |
| `RUN_MODE` | `once` / `loop` | Однократно или бесконечный цикл |
| `DRY_RUN` | `true` / `false` | Только лог, без записи |
| `SAFE_MODE` | `true` / `false` | Жёсткие отказы при подозрительных ситуациях |
| `ALLOW_PRODUCTION_WRITE` | `true` / `false` | В prod без `true` всё уходит в dry-run |
| `OUR_GRID_ID` | id | ID источника "Наша сетка" |
| `YANDEX_GRID_ID` | id | ID источника "Яндекс сетка" |
| `SUMMARY_SPREADSHEET_ID` | id | ID итоговой таблицы |
| `EXPECTED_OUR_TITLE` | строка | Если задано, фактическое название должно совпадать |
| `EXPECTED_YANDEX_TITLE` | строка | То же |
| `EXPECTED_SUMMARY_TITLE` | строка | То же |
| `SUMMARY_SETTINGS_SHEET_NAME` | строка | По умолчанию `Settings` |
| `REQUIRE_YANDEX` | `true` / `false` | Если `true` — падать при недоступном Яндексе |
| `ALLOW_CREATE_SHEETS` | `true` / `false` | Разрешить создание листов (по умолчанию нет) |
| `CREATE_LOG_SHEET` | `true` / `false` | Создавать `Sync_Log` если нет |
| `SUMMARY_WRITE_START_COL` | буква | По умолчанию `A` |
| `SUMMARY_WRITE_END_COL` | буква | По умолчанию `M` |
| `UPDATED_AT_CELL` | A1-ссылка | По умолчанию `N1` |
| `CLEAR_TAIL` | `true` / `false` | По умолчанию `false`: хвост блока не чистится |
| `MAX_DROP_RATIO` | `0..1` | Защита: при падении больше — отказ от записи |
| `RED_GAP_ROWS` | int | Сколько подряд пустых строк включают "красную секцию" |
| `MAX_DATA_ROWS` | int | Верхняя граница чтения из источника |
| `BACKUP_BEFORE_WRITE` | `true` / `false` | JSON-снимок перед записью |
| `BACKUP_DIR` | путь | По умолчанию `backups` |
| `LOOP_SLEEP_SEC` | int | Пауза между итерациями в `--loop` |
| `GCP_SA_JSON` | JSON | Service account JSON в одной строке |
| `GOOGLE_CREDS_B64` | base64 | base64 от service account JSON |
| `GOOGLE_APPLICATION_CREDENTIALS` | путь | Путь к JSON-файлу |
| `TZ` | строка | По умолчанию `Asia/Almaty` |

### Что НЕ показывать публично

`GCP_SA_JSON`, `GOOGLE_CREDS_B64`, `GOOGLE_APPLICATION_CREDENTIALS` (содержимое),
а также при желании — сами spreadsheet IDs.

---

## 7. Что делать при ошибках Google API

| Ошибка | Что значит | Что делать |
|---|---|---|
| **403** | Нет доступа к таблице | Поделись таблицей с `client_email` service account как Editor |
| **404** | Не та таблица или удалена | Проверь spreadsheet ID; возможно лист переименован |
| **401** | Креды невалидны | Перевыпусти JSON service account, проверь поля |
| **429** | Rate limit | Скрипт ретраит до 3 раз. Если повторяется — увеличь `LOOP_SLEEP_SEC` |
| **500/502/503/504** | Сбой Google | Скрипт ретраит до 3 раз. Если стабильно — подожди |
| **JSON not valid** | Битый `GCP_SA_JSON` | Перепроверь экранирование переносов строк (`\n`) |

На 403/404/401 скрипт НЕ ретраит и выходит сразу — лучше остановиться, чем
бомбить API.

---

## 8. Что делать, если YANDEX sheet удалён/недоступен

Установи:
```
REQUIRE_YANDEX=false
```

Тогда при недоступной Яндекс-таблице скрипт:
- продолжит обновлять блок `НАША СЕТКА`;
- НЕ будет трогать существующий блок `ЯНДЕКС СЕТКА`;
- в лог напишет `YANDEX source unavailable, skipped YANDEX block`.

Если `REQUIRE_YANDEX=true` — запуск завершится без записи.

---

## 9. Красные строки

В источнике могут быть "красные" строки — обычно ниже основного блока, после
нескольких пустых строк подряд.

Поведение:
- Пока не встретилось `RED_GAP_ROWS` подряд пустых строк по колонке "Менеджер" —
  строки идут в обычные показатели.
- После порога — все последующие строки с менеджером считаются только в "Красные".
- Дополнительно, если в строке встретился текст `красн` / `red` / `красный` /
  `красная`, она добавляется в "Красные". **Это поведение унаследовано:
  строка с таким текстом, не находящаяся в "красной секции", считается и
  в обычные показатели, и в "Красные".** Это совместимо со старым кодом.

> ⚠️ Скрипт **не читает цвет фона** через API (Sheets API `values.get` его
> не возвращает). Если в источнике после пустого разрыва идут жёлтые строки,
> они тоже будут учтены как "красные" с точки зрения этого скрипта. Это
> допустимо — так работала старая логика.

Управление: переменная `RED_GAP_ROWS` (по умолчанию `5`).

---

## 10. Что считается каждой колонкой (по менеджеру)

Логика 1-в-1 как в предыдущей версии:

- **Офферты всего** — количество обычных (не красных) строк менеджера.
- **ИП** — в тексте строки/в колонке ОПФ есть `ип ` / `ип"` / `жк `.
- **ТОО** — в тексте строки есть `тоо`.
- **Договор есть** — значение в колонке "Договор" не пустое и не одно из:
  `нет`, `0`, `-`, `—`.
- **Акцепт/Оплата** — в колонке "Акцепт" значение длиной > 1 и не содержит
  `нет`, `отказ`, `ошибка`.
- **Акцепт %** — `round(accept / total * 100)%`, строкой.
- **Метка nib_sale / nib / 0 / Пусто / Другое** — взаимоисключающие категории
  значения колонки "Наличие метки".
- **Красные** — см. раздел 9.

---

## 11. Что скрипт никогда не сделает

- Не удалит лист (`deleteSheet`).
- Не удалит строки/столбцы (`deleteDimension`).
- Не очистит весь лист.
- Не очистит весь spreadsheet.
- Не запишет в колонки за пределами `A:M`.
- Не запишет в `N1`, если `UPDATED_AT_CELL` не на `N1` (но N1 — дефолт).
- Не создаст лист в `SAFE_MODE=true`.
- Не запишет в production без `ALLOW_PRODUCTION_WRITE=true`.
- Не очистит хвост блока, если `CLEAR_TAIL=false`.
- Не очистит хвост, если новых данных меньше старых на `MAX_DROP_RATIO`.
- Не запишет, если структура блока не совпадает с ожидаемой.
- Не запишет, если источник вернул 0 строк.

---

## 12. Деплой на Railway

### 12.1. Этап 1: dry-run на копиях

Переменные:
```
ENVIRONMENT=test
RUN_MODE=loop
DRY_RUN=true
SAFE_MODE=true
ALLOW_PRODUCTION_WRITE=false
REQUIRE_YANDEX=false

OUR_GRID_ID=ID_КОПИИ_OUR
YANDEX_GRID_ID=ID_КОПИИ_YANDEX
SUMMARY_SPREADSHEET_ID=ID_КОПИИ_SUMMARY

GCP_SA_JSON=<полный JSON service account одной строкой>

TZ=Asia/Almaty
LOOP_SLEEP_SEC=15
```

**Start Command:**
```
python main.py --loop
```

Посмотри Deploy logs — должно быть много `[DRY] would WRITE ...`.

### 12.2. Этап 2: реальная запись на копиях

```
DRY_RUN=false
```

Перезапусти. Открой копии — проверь блоки и timestamp.

### 12.3. Этап 3: production

```
ENVIRONMENT=production
ALLOW_PRODUCTION_WRITE=true
DRY_RUN=false
SAFE_MODE=true

OUR_GRID_ID=1OgF4xLUqwSHs2S2NPXCJfsVgLM4V9W5c3yQq5lDoS-o
YANDEX_GRID_ID=1Qf4vPXqfpa83NkCrsTa0OHqv5PTKUB89_lzXIok1NWQ
SUMMARY_SPREADSHEET_ID=18nfkpxiPG6xB7uLcCpwV8Qwmx7VJ3ii7uHUzfUh4L2Y
```

> На Railway нет постоянного диска, поэтому `BACKUP_DIR=backups` будет
> работать только в пределах одного контейнера. Если нужен надёжный backup —
> подключи volume или храни снимки во внешнем хранилище.

---

## 13. Полезные команды

```bash
# виртуальное окружение
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env

# проверка конфига + dry-run
python main.py --once --dry-run
python main.py --once --dry-run --month "Май 2026"

# реальная запись (если в .env DRY_RUN=false)
python main.py --once
python main.py --once --month "Май 2026"

# бесконечный цикл
python main.py --loop
```

---

## 14. Архитектура

```
main.py            CLI: --once / --loop / --dry-run / --month
                   парсит env через python-dotenv, валидирует Config,
                   запускает run_summary_once().

summary_sync.py    Вся бизнес-логика:
                     Config              — все настройки
                     build_sheets_service — приоритет GCP_SA_JSON →
                                            GOOGLE_CREDS_B64 → файл → ADC
                     verify_spreadsheet_title — защита от записи не туда
                     locate_block        — поиск и проверка структуры блока
                     analyze_single_sheet — подсчёт по менеджерам
                                            (СОХРАНЕНА БЕЗ ИЗМЕНЕНИЙ)
                     read_block_old_values — для backup и drop-ratio
                     backup_block_to_disk  — JSON-снимок
                     decide_write        — решает, можно ли писать
                     apply_write_decision — реально пишет (если не dry-run)
                     run_month_update    — обработка одного месяца
                     run_summary_once    — обход всех месяцев из Settings

worker.py          Тонкая обёртка, эквивалент `python main.py --loop`.
                   Не имеет собственного бесконечного цикла.

requirements.txt   google-api-python-client, google-auth, python-dotenv
.env.example       шаблон конфига
backups/           локальные JSON-снимки блоков перед записью
```

---

## 15. Если что-то пошло не так

1. Снимки в `backups/` содержат старые значения блока. Можно вручную
   восстановить из JSON (поле `values`).
2. В Google Sheets есть встроенная история версий: Файл → История версий.
3. Sync_Log в SUMMARY (если включён) хранит лог последних запусков.
