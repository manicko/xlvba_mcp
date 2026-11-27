---
id: ru-xlvba-installation-and-setup
domain: ru-02-installation-and-setup
tags:
  - installation
  - setup
  - windows
  - excel
  - uv
related:
  - ru-xlvba-overview
  - ru-xlvba-configuration
  - ru-xlvba-troubleshooting
  - ru-xlvba-mcp-tools-reference
---

## Назначение

Документ позволяет новому пользователю установить `xlvba` с нуля и проверить, что сервер
запускается и базовые read-only операции работают. Здесь описаны требования, подготовка
Windows и Excel, установка через `uv`, подключение MCP-клиента, первый запуск и расположение
данных времени выполнения.

Все команды приведены для PowerShell (Windows). Ни одна из команд установки не изменяет
исходные книги пользователя.

## Требования к системе

| Компонент | Требование | Примечание |
|-----------|------------|------------|
| ОС | Windows 10/11 x64 | Валидировано на Windows 10 Enterprise LTSC 2019 (сборка 17763). Только Windows — нужны COM и VBOM. |
| Excel | Настольный Microsoft Excel с COM/VBE | Валидировано на 16.0.14332 x64. Универсальная платформа (Microsoft Store/UWP) в общем случае не подходит. |
| Python | `>=3.11,<3.15` x64 | Валидировано на 3.14 x64 (`.python-version`). |
| uv | Установленный [uv](https://docs.astral.sh/uv/) | Менеджер пакетов и окружений. |
| Права | Запись во временный каталог, доступ к файлам книг | Книги должны находиться внутри разрешённых корней (см. конфигурацию). |
| Диск | Несколько МБ под `.venv` и рантайм-данные | `.venv` и `<temp>/excel_vba_mcp/` растут по мере работы. |

Почему только Windows: сервер управляет Excel через COM (`pywin32`) и использует VBIDE
(VBA project object model). Оба механизма доступны только в настольной версии Excel на Windows.

## Подготовка Windows и Excel

1. Установите настольный Microsoft Excel (не UWP/Store-версию в общем случае) и убедитесь,
   что приложение хотя бы раз запускалось вручную и активировано.
2. Включите доверие к VBA project object model (VBOM):
   - Excel → `File` → `Options` → `Trust Center` → `Trust Center Settings` → `Macro Settings`;
   - установите флажок **«Trust access to the VBA project object model»**;
   - подтвердите и перезапустите Excel.
3. Учтите, что корпоративные политики (Group Policy, Intune, реестр) могут блокировать
   VBOM или запуск макросов. Если параметр недоступен или сбрасывается, это блокирует VBA-функции.
4. Закройте все открытые книги и сам Excel перед операциями, которые требуют монопольного
   доступа к файлам (например, перед восстановлением из снапшота).
5. Проверьте настройки безопасности макросов: Excel может блокировать макросы в файлах,
   полученных из интернета (Protected View, «Block macros from running in Office files from
   the Internet»). Это влияет на execute-режим и запуск макросов.

> [!NOTE]
> `excel_open` по умолчанию использует `mode="execute", read_only=False` — сессия доступна для
> записи и выполнения. Передайте `mode="inspect"` и/или `read_only=True`, чтобы выбрать сессию с
> отключёнными макросами или только для чтения. `inspect` отключает макросы
> (`AutomationSecurity = ForceDisable`), `execute` применяет `ByUI` непосредственно перед
> открытием. Ни один из режимов не меняет настройки Trust Center или реестра.

## Установка

1. Откройте терминал в корне репозитория (там, где лежит `pyproject.toml`).
2. Установите зависимости и создайте окружение:

   ```powershell
   uv sync --locked
   ```

   Команда создаёт каталог `.venv/` и консольный скрипт `excel-vba-mcp`. Группа `dev`
   (pytest, ruff, mypy) устанавливается по умолчанию. Глобальная установка и Node.js не требуются.

3. Если менялся `pyproject.toml`, намеренно обновите lock и окружение:

   ```powershell
   uv lock
   uv sync
   ```

## Проверка установки

1. Запустите сервер напрямую — он стартует как STDIO-сервер и ожидает JSON-RPC на `stdin`:

   ```powershell
   uv run excel-vba-mcp
   ```

   Ожидаемое поведение: процесс запускается и ждёт ввод; при отсутствии входных сообщений
   он завершается (или остаётся ждать). Это нормально для STDIO-сервера. Если Excel недоступен,
   сервер всё равно стартует (см. [Проверка работоспособности](#проверка-работоспособности)).

2. Альтернативный запуск через модуль:

   ```powershell
   uv run python -m excel_vba_mcp
   ```

3. Проверка наличия CLI-флагов. В текущей реализации (`src/excel_vba_mcp/__main__.py`)
   CLI-флаги отсутствуют: точка входа `main()` просто вызывает `mcp.run(transport="stdio")`
   без разбора аргументов. Поэтому команда `uv run excel-vba-mcp --help` не документируется —
   флагов `--help` нет.

## Подключение MCP-клиента

### Kilo

```json
{
  "$schema": "https://app.kilo.ai/config.json",
  "mcp": {
    "xlvba": {
      "type": "local",
      "command": [
        "uv",
        "run",
        "--directory",
        "<absolute-path-to-repo>",
        "excel-vba-mcp"
      ],
      "enabled": true,
      "timeout": 60000
    }
  },
  "permission": {
    "xlvba_*": "ask"
  }
}
```

- Замените `<absolute-path-to-repo>` на абсолютный путь к корню репозитория.
- Инструменты показываются с префиксом `xlvba_` (например, `xlvba_excel_open`).
- Блок `permission` задаёт политику доступа клиента к инструментам: `"ask"` спрашивает
  разрешение перед каждым вызовом. Read-only инструменты можно внести в allowlist позже,
  например `"xlvba_server_health": "allow"`.

### Claude Desktop / generic MCP client

```json
{
  "mcpServers": {
    "xlvba": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "<absolute-path-to-repo>",
        "excel-vba-mcp"
      ]
    }
  }
}
```

Kilo также принимает модульную форму (`uv run --directory <repo> python -m excel_vba_mcp`),
консольный скрипт после `uv sync` (`<repo>\.venv\Scripts\excel-vba-mcp.exe`) или глобальную
установку (`uv tool install --editable <repo>`).

## Первый запуск

Сервер по умолчанию работает в **режиме полного доступа (`full`)**: `excel_open` открывает
сессию, доступную для записи и выполнения, а мутирующие инструменты и `macro_run` авторизованы
без отдельного гранта на действие (базовый снапшот перед мутацией всё равно требуется). Чтобы
включить политику deny-by-default, задайте `EXCEL_VBA_MCP_ACCESS_MODE=restricted` перед запуском
сервера; любое другое значение для этой переменной приводит к быстрому сбою старта. См.
[Конфигурация](03-configuration.md).

Минимальная read-only последовательность (все аргументы — JSON). Передайте
`mode="inspect", read_only=true`, чтобы открыть книгу только для чтения, хотя по умолчанию
`excel_open` открывает сессию для записи:

1. Откройте книгу как read-only inspect-сессию. Вызов **`excel_open`** с аргументами:

   ```json
   { "path": "C:/data/report.xlsm", "mode": "inspect", "read_only": true }
   ```

   Ожидаемый ответ содержит: `workbook_id` (назначенный сервером идентификатор),
   `source_path` (исходный путь), `path` (путь рабочей копии) и `read_only: true`.
   Путь должен находиться внутри разрешённых корней (по умолчанию — CWD и системный temp;
   см. [Конфигурация](03-configuration.md)).

2. Получите список листов. Вызов **`sheets_list`** с аргументами:

   ```json
   { "workbook_id": "<workbook_id>" }
   ```

   Ожидаемый ответ — массив листов с полями `name`, `index`, `object_type`, `visibility`,
   а для рабочих листов также `code_name`, `protected`, `table_count`, `used_range`.

3. Прочитайте ограниченный диапазон. Вызов **`range_read_values`** с аргументами:

   ```json
   { "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1:C10", "offset": 0, "limit": 10 }
   ```

   Ожидаемый ответ — типизированные значения ячеек и флаг `truncated`.

4. Закройте книгу без сохранения. Вызов **`excel_close`** с аргументами:

   ```json
   { "workbook_id": "<workbook_id>", "save_changes": false }
   ```

   Ожидаемый ответ подтверждает закрытие сессии.

## Проверка работоспособности

- **`server_health`** — read-only liveness-проба, никогда не обращается к Excel/COM.
  Показывает, что сервер жив и отвечает.
- **`excel_health`** — состояние воркера/сессии (`state`, `accepting`, `owned Excel pid`,
  `active operation`, `needs_recovery`) без принудительного перезапуска и без обращения к COM.
- **`excel_list_workbooks`** — список открытых книг из bookkeeping сессий.

Что означает «Excel недоступен»: сервер всё равно запускается по STDIO (ошибка старта воркера
логируется в `stderr`, но процесс не падает), `server_health` и листинг инструментов продолжают
работать, а инструменты работы с книгами возвращают правдивую ошибку о недоступности сессии,
а не «ложный успех».

## Расположение данных времени выполнения

По умолчанию рантайм-данные пишутся вне репозитория, в системный temp-каталог:

```text
<temp>/excel_vba_mcp/
├─ work/          # рабочие копии открытых книг
├─ snapshots/     # хранилище снапшотов (манифест + SHA-256)
├─ artifacts/     # артефакты тестовых прогонов и доказательства
├─ runs/          # сохранённые декларативные тестовые прогоны
└─ audit.jsonl    # append-only журнал аудита (JSONL)
```

`<temp>` — системный temp-каталог (`%TEMP%` на Windows). Исходные файлы пользователя никогда
не изменяются: открывается только копия из `work/`.

## Обновление

1. Получите изменения и при необходимости обновите lock:

   ```powershell
   uv lock
   uv sync
   ```

2. Перезапустите MCP-клиент, чтобы он подхватил новую версию сервера.

## Миграция поверхности инструментов (single switchover)

Поверхность MCP-инструментов была консолидирована с **92 инструментов до 41** за **один переход**
(single switchover, без dual-run). Продакшн `tools/list` идёт напрямую 92 → 41: **нет**
`meta.deprecated`-алиасов и **нет** окна совместимости. Клиенты, которые сопоставляют старые
per-verb имена инструментов, должны мигрировать на консолидированный инструмент + форму `action`.

| Старый per-verb инструмент | Новая форма |
|---------------------------|-------------|
| `excel_open` | `workbook` `action="open"` |
| `excel_close` / `excel_save` | `workbook` `action="close"` / `action="save"` |
| `range_read_values` | `range_read` `action="read-values"` |
| `table_update_rows` | `table` `action="update-rows"` |
| `macro_classify_background` | `macro` `action="classify-background"` |
| `test_suite_run` | `test` `action="run-suite"` |

**Approval globs.** Approval в Kilo остаётся name-glob'ом `xlvba_*`; аннотации инструментов
(`readOnly` / `destructive` / `idempotent`) — это **подсказки, а не энфорсмент**. Высокорисковые
standalone-имена не изменились и сохраняют own permission: `macro_run`, `excel_save_as`,
`table_delete_column`. `workbook` `close` — защищённое действие: при несохранённых изменениях оно
**по умолчанию отказывает**; передайте `discard_changes=true`, чтобы отбросить (или
`save_changes=true`, чтобы сохранить при закрытии). См. [Конфигурация](03-configuration.md) и
[Справочник инструментов MCP](05-mcp-tools-reference.md).

**Финальные группы инструментов (41):**

- **Lifecycle / сессия:** `workbook`, `workbook_read`, `excel_save_as`, `excel_health`, `excel_recover`
- **Снапшоты:** `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff`
- **Jobs:** `job_get`, `job_request_cancel`
- **VBA:** `vba_read`, `vba_export_project`, `vba_replace_component_source`, `vba_insert_code`, `vba_compile_diagnostics`
- **Macro / test:** `macro`, `macro_run`, `test`
- **Листы / таблицы / имена:** `sheet`, `sheet_read`, `sheet_delete`, `table`, `table_read`, `table_delete_rows`, `table_delete_column`, `name`, `name_read`, `name_delete`
- **Диапазоны:** `range_read`, `range_write`, `range_structural`, `range_delete`, `range_merge`
- **Pivot:** `pivot_read`, `pivot_refresh`, `pivot_field`
- **Объекты:** `shape`, `control`
- **Возможности / здоровье:** `xlvba_environment`, `server_health`

## Удаление

1. Удалите виртуальное окружение:

   ```powershell
   Remove-Item -Recurse -Force .venv
   ```

2. При необходимости удалите рантайм-каталог:

   ```powershell
   Remove-Item -Recurse -Force "$env:TEMP\excel_vba_mcp"
   ```

Исходные книги и файлы пользователя при удалении сервера не затрагиваются.

## Следующие шаги

- [Конфигурация](03-configuration.md) — переменные окружения, пути, политика путей, сессии.
- [Справочник инструментов MCP](05-mcp-tools-reference.md) — полный список инструментов.
- [Архитектура](04-architecture.md) — если Excel недоступен или VBOM заблокирован.