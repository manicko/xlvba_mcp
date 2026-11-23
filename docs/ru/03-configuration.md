---
id: ru-xlvba-configuration
domain: ru-03-configuration
tags:
  - configuration
  - environment
  - security
  - paths
  - audit
related:
  - ru-xlvba-overview
  - ru-xlvba-installation-and-setup
  - ru-xlvba-security
  - ru-xlvba-troubleshooting
  - ru-xlvba-architecture
---

## Назначение

Документ — справочник по всем параметрам конфигурации, которые реально существуют в `xlvba`:
переменные окружения, каталоги времени выполнения, политика путей, параметры сессии, значения,
задаваемые в коде, логирование и аудит.

## Способы конфигурации

Конфигурации в виде файла нет. Все настройки задаются **исключительно переменными окружения**,
которые читаются при старте сервера. Модель конфигурации типизирована через Pydantic
(`src/excel_vba_mcp/config.py`) и теперь подключена к серверу: `config.load_config()` вызывается
в `server.lifespan` при старте, поэтому `EXCEL_VBA_MCP_ACCESS_MODE` и `EXCEL_VBA_MCP_LOG_LEVEL`
валидируются вместе, и неверное значение любой из них приводит к быстрому сбою старта
(`ValidationError`).

## Переменные окружения

| Переменная | Назначение | По умолчанию | Допустимые значения | Эффект |
|------------|------------|--------------|---------------------|--------|
| `EXCEL_VBA_MCP_ACCESS_MODE` | Политика доступа сервера (поведение approval) | `full` | `full`, `restricted` | `full` авторизует `MUTATING`/`EXECUTING` без гранта на действие; `restricted` сохраняет deny-by-default approval. Любое другое значение вызывает быстрый сбой старта (`ValidationError`). |
| `EXCEL_VBA_MCP_ROOTS` | Allowlist корневых каталогов для путей книг | CWD + системный temp | Строка путей, разделённых `os.pathsep` (на Windows `;`) | Пути вне корней отклоняются с `PATH_NOT_ALLOWED`. Пустое/отсутствующее значение → значение по умолчанию. |
| `EXCEL_VBA_MCP_WORK_DIR` | Каталог рабочих копий | `<temp>/excel_vba_mcp/work` | Абсолютный путь | Где создаются рабочие копии открываемых книг. |
| `EXCEL_VBA_MCP_SNAPSHOT_DIR` | Хранилище снапшотов | `<temp>/excel_vba_mcp/snapshots` | Абсолютный путь | Каталог снапшотов с манифестом и SHA-256. |
| `EXCEL_VBA_MCP_ARTIFACT_DIR` | Каталог артефактов | `<temp>/excel_vba_mcp/artifacts` | Абсолютный путь | Артефакты и доказательства; также регистрируется как логический корень `artifacts`. |
| `EXCEL_VBA_MCP_RUN_DIR` | Каталог сохранённых прогонов | `<temp>/excel_vba_mcp/runs` | Абсолютный путь | Хранилище декларативных тестовых прогонов. |
| `EXCEL_VBA_MCP_AUDIT_LOG` | Путь к журналу аудита | `<temp>/excel_vba_mcp/audit.jsonl` | Абсолютный путь | Append-only JSONL-журнал. |
| `EXCEL_VBA_MCP_LOG_LEVEL` | Уровень логирования | `INFO` | `DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL` | Разбирается в `config.load_config()`, но **не применяется**. См. раздел «Логирование». |

`EXCEL_VBA_MCP_ROOTS` — список каталогов, разделённых `os.pathsep` (на Windows — `;`).
Каждый элемент становится логическим корнем `root_N`; артефактный каталог дополнительно
регистрируется под именем `artifacts`. Если переменная пуста или отсутствует, используются
текущий рабочий каталог (`Path.cwd()`) и системный temp-каталог.

> [!IMPORTANT]
> `EXCEL_VBA_MCP_ROOTS` читается **один раз при старте сервера** в `server._configured_roots()`.
> После изменения переменной **необходимо перезапустить MCP-сервер** (перезагрузить сервер в
> Kilo / MCP-клиенте), иначе новый корень не вступит в силу. Запущенная сессия сохраняет исходный
> allowlist.

> [!TIP]
> `EXCEL_VBA_MCP_ROOTS` — поддерживаемый способ разрешить дополнительные деревья проектов; правка
> исходного кода не требуется. Пути вне всех настроенных корней отклоняются с `PATH_NOT_ALLOWED`
> (reason `outside_roots`); в `context` ошибки также перечислены текущие разрешённые корни
> (`allowed_roots`), чтобы было проще диагностировать пропущенную запись.

`EXCEL_VBA_MCP_LOG_LEVEL` разбирается функцией `config.load_config()` и проходит валидацию через
Pydantic, но фактически **не влияет** на логирование: логирование настраивается один раз при импорте
модуля через `_configure_logging()` с фиксированным уровнем `INFO`, до вызова `load_config()` в
lifespan. Переменная принимается (и валидируется), но в этой редакции не имеет эффекта.

## Пути и каталоги

| Каталог | Переменная | Назначение |
|---------|------------|------------|
| `work/` | `EXCEL_VBA_MCP_WORK_DIR` | Рабочие копии открытых книг; открывается только копия, оригинал не изменяется. |
| `snapshots/` | `EXCEL_VBA_MCP_SNAPSHOT_DIR` | Снапшоты: манифест `manifest.json` и SHA-256; снапшоты иммутабельны. |
| `artifacts/` | `EXCEL_VBA_MCP_ARTIFACT_DIR` | Артефакты тестовых прогонов и доказательства; регистрируется как логический корень `artifacts`. |
| `runs/` | `EXCEL_VBA_MCP_RUN_DIR` | Сохранённые декларативные тестовые прогоны (по `run_id`). |
| `audit.jsonl` | `EXCEL_VBA_MCP_AUDIT_LOG` | Append-only журнал аудита в формате JSONL. |

Каталог `artifacts` помимо пользовательского назначения регистрируется в path policy под
логическим именем `artifacts`. Каталоги снапшотов и артефактов дополнительно допускают
расширения `.json` и `.sha256` (для манифестов и контрольных сумм).

## Политика путей

Все пути, попадающие в Excel COM, проходят через `PathPolicy.validate_path`
(`src/excel_vba_mcp/security/path_policy.py`). Политика канонизирует путь через
`Path.resolve` и перекрёстно проверяет его через `os.path.realpath`.

Запрещённые категории (поле `reason` в контексте ошибки, код `PATH_NOT_ALLOWED`):

| `reason` | Значение |
|----------|----------|
| `empty_path` | Пустой путь. |
| `device_path_denied` | Device-путь (`\\?\`, `\\.\` и их варианты). |
| `unc_denied` | UNC-путь (`\\server\share`), если `allow_unc` отключён. |
| `traversal` | `..`-traversal, выходящий за пределы корня. |
| `outside_roots` | Путь вне всех разрешённых корней. |
| `symlink_escape` | Лексически внутри корня, но разрешается наружу (symlink/junction/reparse point). |
| `root_mismatch` | Путь не в ожидаемом логическом корне. |
| `extension_not_allowed` | Расширение файла не входит в разрешённый набор. |

Разрешённые расширения книг: `.xlsm`, `.xlsb`, `.xlsa`, `.xlam`, `.xltm`, `.xltx`, `.xlsx`, `.xls`.
Для корней `snapshots` и `artifacts` дополнительно разрешены `.json` и `.sha256`.

Параметр `allow_unc` всегда `False` (настройка задаётся в коде и недоступна через переменные
окружения), поэтому UNC- и device-пути по умолчанию запрещены.

## Параметры сессии

`excel_open(path, mode, read_only)` принимает:

| Параметр | Значения | Смысл |
|----------|----------|-------|
| `mode` | `execute` (по умолчанию), `inspect` | `execute` применяет `ByUI` непосредственно перед открытием (макросы включены политикой Office UI); `inspect` открывает книгу с отключёнными макросами (`AutomationSecurity = ForceDisable`). |
| `read_only` | `false` (по умолчанию), `true` | Read-only-сессия открывает только чтение; мутации в такой сессии недоступны. |

Влияние на approval (зависит от `EXCEL_VBA_MCP_ACCESS_MODE`):

- В режиме `full` (по умолчанию) `excel_open` с `read_only: false` или `mode: "execute"`
  авторизован политикой; отдельный грант не нужен.
- В режиме `restricted` `excel_open` с `read_only: false` или `mode: "execute"` требует записи в
  allowlist или живого гранта и иначе **отклоняется** с `APPROVAL_DENIED`.
- `mode: "inspect"` + `read_only: true` проходит как read-only и не требует grant ни в одном режиме.
- Ни `inspect`, ни `execute` не изменяют настройки Trust Center или реестра.

## Параметры, задаваемые в коде

Следующие значения существуют в коде, но **не настраиваются через переменные окружения или MCP**.
Они приведены как справка, а не как настройки:

- **Мягкие таймауты воркера.** Для большинства операций `_DEFAULT_TIMEOUT_S = 60.0` секунд
  (session, snapshot, sheets, tables, ranges, names, shapes, controls, pivots, vba_project,
  compile). Для `macro_service` `_DEFAULT_TIMEOUT_S = 120.0`; в `test_runner` используются
  `_CALL_TIMEOUT_S = 60.0` и `_MACRO_TIMEOUT_S = 120.0`.
- **`orphan_timeout_s`.** Таймаут "осиротевшего" воркера/потока задаётся при построении
  `ExcelComWorker` (значение по умолчанию — `10.0` секунды) и не вынесен в конфигурацию.
- **Path policy `allow_unc`.** Всегда `False`.
- **Approval allowlist.** Пустой (`ApprovalPolicy()` без записей allowlist), и grant-инструмент
  через MCP не предоставляется. В режиме `full` (по умолчанию,
  `EXCEL_VBA_MCP_ACCESS_MODE=full`) этого достаточно, поскольку `MUTATING`/`EXECUTING`
  авторизованы политикой; режим `restricted` снова включает deny-by-default.

## Логирование

- Логи пишутся **только в `stderr`**; `stdout` зарезервирован под протокол MCP.
- Уровень фиксирован на `INFO` (`_configure_logging()` вызывает `logging.basicConfig(level=INFO, stream=sys.stderr, ...)` при импорте).
- `EXCEL_VBA_MCP_LOG_LEVEL` разбирается и валидируется в `config.load_config()` (вызывается при
  старте сервера), но не применяется: уровень задаётся один раз при импорте. В этой редакции
  переменная не имеет эффекта.
- Чувствительные значения не должны попадать в логи и артефакты.

## Аудит

Аудит ведётся в append-only JSONL-журнале (по умолчанию `<temp>/excel_vba_mcp/audit.jsonl`,
переопределяется `EXCEL_VBA_MCP_AUDIT_LOG`). Каждая запись содержит (по возможности):

- `schema_version`, `timestamp` (UTC), `event_id`, `correlation_id`, `session_id`;
- `workbook_id` и `workbook` (канонический путь + SHA-256, имя файла);
- `tool`, `operation_class`;
- `params` — параметры с обязательной редакцией (redaction);
- `approval` — решение approval (источник, id гранта, причина);
- `snapshot_id`, `result`, `user`, `process`, `excel_version`.

Редакция (redaction) обязательна и намеренно «сжимает» данные: значения по «секретным» ключам
(например, `password`, `token`, `secret`) заменяются на `"<redacted>"`, пути сводятся к basename,
аргументы макросов логируются как типизированные и ограниченные по длине сводки без COM-ссылок.

## Примеры конфигурации

### Локальная разработка (явные корни)

```powershell
$env:EXCEL_VBA_MCP_ROOTS = "C:\data\workbooks;D:\projects\excel"
uv run excel-vba-mcp
```

### Локальное дерево проекта (`C:\py_exp\mko_treport`)

Разрешите всё дерево `mko_treport`, чтобы `xlvba_excel_open` принимал книги внутри него. Задайте
переменную **до** запуска сервера, затем запустите/перезагрузите MCP-сервер:

```powershell
# Разрешить только дерево mko_treport (плюс любые другие корни через `;`).
$env:EXCEL_VBA_MCP_ROOTS = "C:\py_exp\mko_treport"
uv run excel-vba-mcp
```

Чтобы сохранить значения по умолчанию (CWD + temp) **и** добавить проект, соедините их явно через
`;` (Windows) — явное значение не объединяется со значениями по умолчанию:

```powershell
$env:EXCEL_VBA_MCP_ROOTS = "$env:TEMP;C:\py_exp\mko_treport"
```

> [!NOTE]
> Перезапустите / перезагрузите MCP-сервер после установки переменной. Значение читается только
> при старте; живая сессия продолжает использовать исходный allowlist. Если `xlvba_excel_open`
> всё ещё возвращает `PATH_NOT_ALLOWED / outside_roots`, убедитесь, что сервер перезапущен и путь
> находится внутри одного из настроенных корней (см. `context.allowed_roots` в ошибке).

### Production-ish (выделенные каталоги + аудит)

```powershell
$env:EXCEL_VBA_MCP_ROOTS = "D:\xlvba\roots"
$env:EXCEL_VBA_MCP_WORK_DIR = "D:\xlvba\work"
$env:EXCEL_VBA_MCP_SNAPSHOT_DIR = "D:\xlvba\snapshots"
$env:EXCEL_VBA_MCP_ARTIFACT_DIR = "D:\xlvba\artifacts"
$env:EXCEL_VBA_MCP_RUN_DIR = "D:\xlvba\runs"
$env:EXCEL_VBA_MCP_AUDIT_LOG = "D:\xlvba\audit.jsonl"
uv run excel-vba-mcp
```

### Ограниченная конфигурация (один корень)

```powershell
$env:EXCEL_VBA_MCP_ROOTS = "C:\data\readonly"
uv run excel-vba-mcp
```

### Режим restricted (deny-by-default)

```powershell
$env:EXCEL_VBA_MCP_ACCESS_MODE = "restricted"
uv run excel-vba-mcp
```

С `restricted` мутирующие и исполняющие инструменты отклоняются с `APPROVAL_DENIED`, если нет
записи в allowlist или живого гранта; через MCP доступна только работа для чтения. Если переменную
не задавать (или задать `full`), сохраняется поведение полного доступа по умолчанию.

## Диагностика конфигурации

Если инструменты возвращают `PATH_NOT_ALLOWED`, проверьте, что путь к книге находится внутри
`EXCEL_VBA_MCP_ROOTS` (по умолчанию — CWD и temp). `context.allowed_roots` ошибки перечисляет
корни, которые реально загрузил работающий сервер; если вашего нового корня там нет — сервер не
был перезапущен после изменения (см. заметку о перезапуске выше). Если мутирующие или исполняющие
инструменты возвращают `APPROVAL_DENIED`, сервер работает в режиме `restricted`
(`EXCEL_VBA_MCP_ACCESS_MODE=restricted`); сбросьте переменную (или задайте `full`) для поведения
полного доступа по умолчанию. Если VBA-функции недоступны, проверьте настройку VBOM в Trust
Center. Подробнее — [Архитектура](04-architecture.md).