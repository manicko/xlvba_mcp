---
id: ru-xlvba-troubleshooting
domain: ru-10-troubleshooting
tags:
  - troubleshooting
  - errors
  - diagnostics
  - excel
  - com
  - vba
related:
  - ru-xlvba-overview
  - ru-xlvba-installation-and-setup
  - ru-xlvba-configuration
  - ru-xlvba-architecture
  - ru-xlvba-mcp-tools-reference
  - ru-xlvba-excel-capabilities
  - ru-xlvba-workflows-and-examples
  - ru-xlvba-security
  - ru-xlvba-limitations-and-compatibility
  - ru-xlvba-development-and-maintenance
---

## Назначение

Документ сопоставляет проблемы, с которыми вы вероятнее всего столкнётесь с `xlvba`, с их причинами,
диагностикой и решениями. Он сфокусирован на ошибках, которые реально видит клиент: структурированных
ошибках инструментов со стабильным `code`, сбоях запуска и проблемах состояния Excel.

Контракт ошибок см. в [Архитектура](04-architecture.md) и
[Справочник инструментов MCP](05-mcp-tools-reference.md). Конфигурация — в
[Конфигурация](03-configuration.md).

## Справочник проблем / ошибок

| Проблема / Ошибка | Возможная причина | Как диагностировать | Решение |
|-------------------|-------------------|---------------------|---------|
| MCP-клиент не может подключиться | Сервер не запущен, неверный путь команды или `uv` не в `PATH` | Проверьте конфигурацию MCP клиента; запустите `uv run excel-vba-mcp` вручную в терминале; ищите ненулевой код выхода | Исправьте `command` / `args` (используйте абсолютный `--directory`); подтвердите, что `uv sync --locked` прошёл успешно; перезапустите клиент |
| Сервер завершается при старте | Неверный `EXCEL_VBA_MCP_ACCESS_MODE`, `EXCEL_VBA_MCP_LOG_LEVEL` или `EXCEL_VBA_MCP_MACRO_TIMEOUT_S`; поднимается `ValidationError` в `lifespan` | Прочитайте строку `stderr` `Invalid server configuration: ...`; проверьте три переменные | Задайте `EXCEL_VBA_MCP_ACCESS_MODE` равным `full` или `restricted`; задайте `EXCEL_VBA_MCP_LOG_LEVEL` равным `DEBUG`/`INFO`/`WARNING`/`ERROR`/`CRITICAL` (или сбросьте его); задайте `EXCEL_VBA_MCP_MACRO_TIMEOUT_S` положительным числом (или сбросьте его) |
| Отсутствуют зависимости Python / сбой `uv sync` | `.venv` устарел, lock неактуален или wheel недоступен для версии Python | Запустите `uv sync --locked`; если он падает из-за изменения `pyproject.toml`, запустите `uv lock`, затем `uv sync` | Пересоберите окружение из lock; используйте поддерживаемый Python (3.11–3.14 x64) |
| Excel не запускается / не отвечает | Настольный Excel не установлен, не активирован или используется UWP/Store-версия | `excel_health` показывает `unavailable` / `needs_recovery`; предупреждение о старте воркера — в `stderr`; `server_health` всё ещё возвращает `ok` | Установите и активируйте настольный Excel; запустите его вручную один раз; закройте другие экземпляры Excel, если они мешают; повторите `excel_open` |
| Утечка процесса `EXCEL.exe` | Принадлежащий скрытый экземпляр не умер при завершении, или макрос/модальное окно удерживали его | Проверьте `excel_health.owned_pid`; осмотрите Task Manager на осиротевший `EXCEL.EXE` | Сервер завершает только свой принадлежащий pid; если процесс выжил, завершите конкретный pid, затем перезапустите сервер. Не убивайте пользовательские процессы Excel |
| COM-ошибки (`COM_BUSY`, `RPC_E_CALL_REJECTED`) | Excel/VBE занят; вызов был отклонён, пока STA был занят | `code` ошибки — `COM_BUSY` или `RPC_E_CALL_REJECTED`; `context.class` показывает класс операции | Повторите **чтение** после короткой паузы (чтения повторяются до 3 раз). Не повторяйте мутацию слепо; подождите и проверьте `excel_health` |
| `PATH_NOT_ALLOWED` | Путь вне `EXCEL_VBA_MCP_ROOTS`, device/UNC/traversal/symlink-путь или запрещённое расширение | Прочитайте `context.reason`: `empty_path`, `device_path_denied`, `unc_denied`, `traversal`, `outside_roots`, `symlink_escape`, `root_mismatch`, `extension_not_allowed`; `context.allowed_roots` перечисляет корни, загруженные работающим сервером. Для `vba_export_project` `out_dir` вне корня сообщает **абсолютный** артефакт-корень в `context.allowed_root` / `context.expected_root` (F4) | Поместите книгу внутрь разрешённого корня или расширьте `EXCEL_VBA_MCP_ROOTS`; используйте обычный путь с разрешённым расширением. `EXCEL_VBA_MCP_ROOTS` читается один раз при старте — **перезапустите/перезагрузите MCP-сервер** после изменения (см. [Конфигурация](03-configuration.md)) |
| `APPROVAL_DENIED` | Сервер работает в режиме `restricted`, и ни одна запись allowlist или живой грант не совпадает | `code` ошибки — `APPROVAL_DENIED`; `excel_get_status.access_mode` сообщает `restricted` | Сбросьте `EXCEL_VBA_MCP_ACCESS_MODE` (или задайте `full`) для полного доступа; сервер в режиме restricted предоставляет через MCP только работу для чтения |
| `WORKBOOK_READ_ONLY` | Сессия была открыта с `read_only=true` | Любой мутирующий инструмент возвращает `WORKBOOK_READ_ONLY`; `excel_get_status.read_only` — `true` | Откройте заново для записи (`excel_open` без `read_only`) или выполняйте только работу для чтения |
| `ACCESS_VBOM_DISABLED` | «Trust access to the VBA project object model» выключено, или корпоративная политика это блокирует | Инструменты VBA/VBIDE и `vba_compile_diagnostics` возвращают `ACCESS_VBOM_DISABLED`; `workbook_get_model` сообщает `vba_access = "vbom_disabled"` | Включите VBOM в Excel Trust Center → Macro Settings, перезапустите Excel, повторите. Если политика это блокирует, возможности VBA остаются недоступны |
| `VBA_PROJECT_LOCKED` | Проект VBA защищён паролем | Инструменты VBA возвращают `VBA_PROJECT_LOCKED`; `vba_project_info` сообщает состояние защиты | Разблокируйте проект вручную в VBE (владелец должен сообщить пароль); защита никогда не обходится |
| `ACCESS_VBOM_DISABLED` против заблокированного проекта | Две разные причины: отсутствие доверия против защищённого проекта | `ACCESS_VBOM_DISABLED` означает, что VBIDE не заслуживает доверия; `VBA_PROJECT_LOCKED` означает, что доверие есть, но проект защищён | Включите VBOM для первого; разблокируйте проект для второго. Не путайте их |
| Сбои макросов (`MACRO_NOT_FOUND`, `MACRO_BLOCKED`, `MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`) | Неверное имя; политика макросов Office; ошибка выполнения в макросе; макрос превысил свой мягкий дедлайн (по умолчанию 90 с, `EXCEL_VBA_MCP_MACRO_TIMEOUT_S`) | `macro_get_last_result` возвращает статус и структурированную ошибку; `context` несёт HRESULT/секцию; `session_recovery` сообщает результат внутрипроцессного восстановления | `MACRO_NOT_FOUND`: исправьте `Module.Procedure`; `MACRO_BLOCKED`: проверьте настройки макросов Office / execute-режим; `MACRO_RUNTIME_ERROR`: осмотрите макрос (модальное окно было завершено, `dismissed=false`); `MACRO_TIMEOUT`: упростите макрос или не ожидайте завершения. После любой из форм таймаута принадлежащий Excel PID завершён, а сессия восстановлена автоматически (F2b) — **откройте книгу заново** (прежний `workbook_id` недействителен) и проверьте `session_recovery` |
| `BACKGROUND_SAFETY_NOT_VERIFIED` | Процедура не классифицирована `background_safe` для текущего хэша книги и ревизии исходников VBA | `macro_run` возвращает `blocked` с `BACKGROUND_SAFETY_NOT_VERIFIED`; макрос не был запущен | Классифицируйте её через `macro_classify_background` после ревью кода; переклассифицируйте после любого изменения исходников VBA; не повторяйте слепо |
| `INTERACTIVE_UI_NOT_SUPPORTED` | Процедура попадает под токен-скан R13, эвристична или была классифицирована безопасной, будучи небезопасной | `macro_run` / `macro_classify_background` возвращает `INTERACTIVE_UI_NOT_SUPPORTED`; `context.flagged_tokens` перечисляет попадания | Удалите интерактивный UI или оставьте процедуру неклассифицированной; не обходите preflight |
| `EXCEL_HUNG` / `needs_recovery` | Вызов превысил свой мягкий дедлайн, Excel завис, или модальное окно блокирует COM | `excel_health` показывает `status = "needs_recovery"` и `needs_recovery = true`; `code` ошибки — `EXCEL_HUNG` | Зависание/таймаут `macro_run` восстанавливает сессию **автоматически** (F2b) — откройте книгу заново. Для любого другого зависшего вызова инструмента восстановления нет: закройте и откройте книгу либо перезапустите процесс сервера. Завершается только принадлежащий pid Excel |
| `SNAPSHOT_FAILED` | Нет успешного снапшота, или запись снапшота не удалась | Любой мутирующий инструмент возвращает `SNAPSHOT_FAILED`; `snapshot_list` пуст для книги | Сначала вызовите `snapshot_create`; проверьте доступ на запись в каталог снапшотов |
| `INVALID_RANGE` / `LISTOBJECT_INTERSECTION` / `PROTECTED_TARGET` | Некорректный/сверхбольшой адрес или payload; операция диапазона пересекает таблицу Excel; цель защищена | Прочитайте `context` (форма payload, `tables`, причина); `INVALID_RANGE` также используется для сбоев схемы аргументов (`reason=invalid_arguments`) | Исправьте адрес/payload; используйте инструменты Tables вместо структурных операций диапазона; снимите защиту листа или выберите другую цель |
| `UNSUPPORTED_PIVOT_SOURCE` / `REFRESH_FAILED` | Источник PivotCache — внешний, OLAP или Data Model; или обновление не вернуло сигнал успеха | `pivot_table_get` / `pivot_cache_get` показывает тип источника; результат обновления сообщает `unsupported_source` / `failed` | Read-only-осмотр всё ещё работает; обновлять можно только кэши на базе базы данных/таблицы. Проверьте исходные данные |
| `WORKBOOK_NOT_FOUND` | Неизвестный или устаревший `workbook_id`, или принадлежащий экземпляр был потерян | `code` ошибки — `WORKBOOK_NOT_FOUND`; `excel_list_workbooks` больше не перечисляет id | Откройте заново через `excel_open` и используйте новый `workbook_id`; перезапустите сервер, если экземпляр был потерян |
| `NOT_FOUND` | Фигура или элемент с указанным `name` не существует (F6) | `shape_get` / `shape_set_on_action` / `control_get` / `control_set_on_action` возвращают `code="NOT_FOUND"` с `context.name` | Сначала перечислите объекты (`shapes_list` / `controls_list`) и используйте точное имя. Это ошибка вызывающей стороны, а не конкуренция — **не повторяйте**. Реальная конкуренция за COM на тех же инструментах по-прежнему сообщается как `COM_BUSY` |
| `REVISION_CONFLICT` / `STALE_TARGET` | Устаревший `expected_source_hash` (правка VBA) или устаревшая идентичность PivotTable | `vba_replace_component_source` / `vba_insert_code` возвращает `REVISION_CONFLICT`; `macro_classify_background` возвращает `STALE_TARGET` | Перечитайте компонент (`vba_read_component`) и передайте текущий `expected_source_hash`; перечитайте цель и повторите |
| `AMBIGUOUS_FILTER` | Фильтр `where` совпал с нулём или несколькими строками | `table_update_rows` / `table_delete_rows` возвращает `AMBIGUOUS_FILTER` с предпросмотром | Используйте явные `indices` или уточните фильтр, пока он не совпадёт ровно с одной строкой |
| Неожиданные или устаревшие значения | Excel не пересчитал; `range_read_values` возвращает кэшированные значения | Перечитайте ячейку; сравните с `range_read_formulas` | Сервер не форсирует `Application.Calculate`; запустите пересчёт в Excel или учитывайте кэшированные значения |
| Таймауты | Большая книга, медленный COM-вызов или зависшая операция | `excel_health` показывает `needs_recovery`; ошибка — `EXCEL_HUNG` | Читайте большие диапазоны порциями; закройте/откройте; перезапустите сервер при необходимости |
| Блокировка книги | Файл заблокирован другим процессом, или тот же путь уже открыт в сервере | `excel_open` сообщает, что путь уже открыт, или происходит ошибка копирования файла | Закройте другой процесс; `excel_close` существующего `workbook_id` перед повторным открытием |
| Неверные пути | Относительный путь, путь вне корней или не-книжный файл | `PATH_NOT_ALLOWED` с `reason` | Используйте абсолютный путь внутри разрешённого корня с разрешённым расширением; см. [Конфигурация](03-configuration.md) |

## Чтение структурированной ошибки

Каждый сбой инструмента достигает клиента как структурированный JSON-payload. Форма такова:

```json
{
  "code": "PATH_NOT_ALLOWED",
  "message": "The path is not allowed by policy.",
  "hresult": null,
  "hresult_hex": null,
  "context": {
    "reason": "outside_roots",
    "operation": "mutating",
    "name": "report.xlsm"
  }
}
```

- `code` — стабильная строка `ErrorCode` (например, `APPROVAL_DENIED`, `PATH_NOT_ALLOWED`,
  `WORKBOOK_READ_ONLY`).
- `message` — человекочитаемое сообщение.
- `hresult` / `hresult_hex` — сырой COM HRESULT и его форма `0x%08X`, или `null`, когда ошибка не
  пришла из COM.
- `context` — структурированный словарь (путь, `workbook_id`, `reason`, `section`, `token`,
  `class` и так далее).

`ExcelVbaMcpServer.call_tool` перехватывает обобщённую обёртку MCP SDK, находит базовый
`ExcelVbaMcpError` в цепочке `__cause__` и повторно поднимает `ToolError`, сообщение которого — JSON
`to_dict()`. Сбои валидации схемы аргументов (`pydantic.ValidationError`) преобразуются в
`INVALID_RANGE` с `reason=invalid_arguments`, поэтому **каждый** сбой инструмента достигает клиента
структурированным.

## Проверки здоровья

Используйте их до и после сбоя:

| Инструмент | Что сообщает | Обращается к Excel? |
|------------|--------------|---------------------|
| `server_health` | Процесс сервера жив: `{"status": "ok", "service": "excel-vba-mcp", "version": ...}` | Нет |
| `excel_health` | Состояние воркера/сессии, `accepting`, `owned_pid`, `active_operation`, `needs_recovery`, `notes`. Выполняет read-only liveness-пробу owned-pid. | Нет (без COM, без принудительного перезапуска) |
| `excel_list_workbooks` | Открытые книги из bookkeeping сессий | Нет |

`excel_health` — самый быстрый способ отличить «сервер в порядке, но вызов книги не удался» от «воркеру
нужно восстановление».

## Сбор диагностики

| Источник | Расположение / инструмент | Использование |
|----------|--------------------------|---------------|
| Логи сервера | `stderr` процесса сервера | Предупреждения запуска, сбои воркера, трассировки (фиксированный уровень `INFO`) |
| Аудит-след | `EXCEL_VBA_MCP_AUDIT_LOG` (по умолчанию `<temp>/excel_vba_mcp/audit.jsonl`) | Одна JSON-запись на операцию: инструмент, отредактированные params, approval, результат |
| Диагностика компиляции | `vba_get_compile_diagnostics` | Кэшированные доказательства компиляции (всегда `is_compile=false` на валидированной сборке) |
| Последний результат макроса | `macro_get_last_result` | Статус, длительность, возвращаемое значение и структурированная ошибка последнего запуска |
| Сохранённый тест-прогон | `test_suite` (`action="get-run"`, с `run_id`) | Per-case и per-step статусы без обращения к Excel |

Сообщая о проблеме, включите имя инструмента, структурированные `code` и `context` ошибки, payload
`excel_health` и соответствующую строку `audit.jsonl` (с отредактированными секретами).

## Когда перезапускать сервер / восстанавливать

- **Зависший или превысивший дедлайн макрос не требует перезапуска.** `macro_run` завершает
  принадлежащий Excel PID и пересоздаёт скрытый экземпляр в процессе (F2b);
  `MacroRunResult.session_recovery` сообщает `"recovered"` или `"needs_recovery"`. Открытые книги
  отбрасываются, поэтому откройте книгу заново через `excel_open` и продолжайте. Сам серверный
  процесс этим путём никогда не перезапускается.
- **Сначала закройте и откройте книгу заново** при устаревшей сессии или блокировке на путь:
  `excel_close(workbook_id, save_changes=false)`, затем снова `excel_open`. Это решает большинство
  проблем с устаревшей сессией и блокировкой на путь.
- **Перезапустите процесс сервера**, когда воркер в `NEEDS_RECOVERY` по любой причине, кроме
  зависания макроса (оно восстанавливается само), принадлежащий процесс Excel исчез или
  `excel_health` сообщает `unavailable`. Автовосстановление есть только на пути макроса;
  `SessionService.recover_session` существует внутри, но не вызывается через MCP.
- **Не убивайте пользовательские процессы Excel.** Сервер когда-либо завершает только свой
  `DispatchEx`-pid; используйте конкретный `owned_pid` из `excel_health`, если нужно очистить
  осиротевший.

## Сбои запуска против рантайм-сбоев

| Симптом | Вероятный слой | Первая проверка |
|---------|----------------|-----------------|
| Клиент сообщает, что сервер не запустился | Процесс / запуск | Запустите `uv run excel-vba-mcp` вручную; проверьте `command` / `args` клиента |
| Сервер запускается, но инструменты не отвечают | Транспорт / конфигурация | `server_health`; проверьте `stderr` на `Invalid server configuration` |
| Инструменты отвечают, но каждый вызов книги падает | Excel / COM | `excel_health`; подтвердите, что настольный Excel запущен и VBOM заслуживает доверия |
| Read-only работает, но мутации падают | Сессия / политика | `excel_get_status`; проверьте `read_only`, `access_mode` и снапшот |
| Одна книга работает, а другая падает | Путь / файл | `PATH_NOT_ALLOWED` reason; блокировка файла; расширение |

## Диагностические рецепты

### Путь отклонён

1. Прочитайте `context.reason` из ошибки `PATH_NOT_ALLOWED`.
2. Подтвердите, что путь абсолютный и находится внутри `EXCEL_VBA_MCP_ROOTS`.
3. Если это UNC- или device-путь, скопируйте файл в разрешённый корень.
4. Повторите `excel_open`.

### Мутация или макрос отклонён

1. Проверьте `excel_get_status.access_mode`.
2. Если это `restricted`, переключитесь на `full` (или примите работу для чтения).
3. Если это `full`, а `APPROVAL_DENIED` сохраняется, перепроверьте класс операции и убедитесь, что
   никакое другое предусловие не сообщается неверно.

### Макрос заблокирован

1. Вызовите `macro_list`, чтобы подтвердить `Module.Procedure` и его классификацию.
2. Прочитайте процедуру через `vba_read_component`.
3. Если `BACKGROUND_SAFETY_NOT_VERIFIED`, вызовите `macro_classify_background` после ревью;
   переклассифицируйте после любого изменения исходников VBA.
4. Если `INTERACTIVE_UI_NOT_SUPPORTED`, удалите отмеченный токен или оставьте процедуру
   неклассифицированной.

### Сессия зависла

1. Вызовите `excel_health`.
2. Если зависание было таймаутом/зависанием `macro_run`, сервер уже восстановился в процессе (F2b):
   прочитайте `session_recovery` из `MacroRunResult` и **откройте книгу заново** (прежний
   `workbook_id` недействителен). Если `session_recovery` — `"needs_recovery"`, перезапустите сервер.
3. Для любого другого состояния `needs_recovery` закройте и откройте книгу заново.
4. Если воркер `unavailable`, перезапустите процесс сервера.
5. Не убивайте пользовательские процессы Excel; завершать можно только принадлежащий pid.

## Перед сообщением о проблеме

Включите: имя инструмента, структурированные `code` и `context` ошибки, payload `excel_health`,
сборку Excel (`excel_version` в аудит-записи или результате `excel_open`) и соответствующую строку
`audit.jsonl` с отредактированными секретами. Обычно этого достаточно, чтобы отличить проблему
конфигурации от проблемы Excel или COM.

## Связанные документы

- [Обзор](01-overview.md) — область.
- [Установка и настройка](02-installation-and-setup.md) — проверки первого запуска.
- [Конфигурация](03-configuration.md) — переменные и пути.
- [Архитектура](04-architecture.md) — обработка ошибок и восстановление.
- [Справочник инструментов MCP](05-mcp-tools-reference.md) — ошибки по инструментам.
- [Возможности Excel](06-excel-capabilities.md) — пределы возможностей.
- [Рабочие процессы и примеры](07-workflows-and-examples.md) — рабочий процесс обработки ошибок.
- [Безопасность](08-security.md) — границы безопасности.
- [Ограничения и совместимость](09-limitations-and-compatibility.md) — известные пределы.
- [Разработка и поддержка](11-development-and-maintenance.md) — руководство для контрибьюторов.