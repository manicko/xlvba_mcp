---
id: ru-xlvba-mcp-tools-reference
domain: ru-05-mcp-tools-reference
tags:
  - mcp
  - tools
  - reference
  - excel
  - vba
related:
  - ru-xlvba-overview
  - ru-xlvba-architecture
  - ru-xlvba-excel-capabilities
  - ru-xlvba-security
  - ru-xlvba-limitations-and-compatibility
  - ru-xlvba-troubleshooting
---

## Назначение

Справочник описывает **все 88 инструментов** MCP-сервера `xlvba`: назначение, параметры,
возвращаемое значение, побочные эффекты, типовые ошибки и класс доступа. Документ
основан на фактических схемах инструментов, полученных с живого `MCPServer`, и на исходных
модулях `api/*`, `application/*` и `adapters/*`.

Документ не пересказывает архитектуру — она описана в [Архитектура](04-architecture.md).
Настройки — в [Конфигурация](03-configuration.md), модель безопасности — в
[Безопасность](08-security.md), пределы — в
[Ограничения и совместимость](09-limitations-and-compatibility.md).

## Как читать справочник

| Элемент | Значение |
|---------|----------|
| `Инструмент` | Точное MCP-имя. Клиенты показывают его с префиксом (например, Kilo — `xlvba_excel_open`). В таблицах — без префикса. |
| `Параметры` | `имя:тип`; `?` означает необязательный параметр; `=value` — значение по умолчанию. `—` означает отсутствие параметров. |
| `Возврат` | Тип DTO или структура ответа. Все DTO сериализуются в JSON и не содержат COM-ссылок. |
| `Побочные эффекты` | Что меняется за пределами ответа: файлы, снапшоты, состояние книги, артефакты, аудит. |
| `Ошибки` | Стабильные коды `ErrorCode` или иные коды, которые может вернуть инструмент. |
| `Доступ` | `read-only` (разрешён автоматически), `mutating` (guarded), `executing` (gated), `mixed` (зависит от параметров). |

Общие соглашения:

1. **Режим доступа.** В режиме `full` (по умолчанию, `EXCEL_VBA_MCP_ACCESS_MODE=full`)
   `ApprovalPolicy` авторизует `mutating` и `executing` операции без гранта на действие. В режиме
   `restricted` действует deny-by-default: такие операции отклоняются с `APPROVAL_DENIED`, если нет
   in-process записи allowlist или живого гранта. Grant-инструмента для approval или редактирования
   allowlist через MCP нет, поэтому сервер в режиме `restricted` через MCP предоставляет только
   работу для чтения. По таксономии ошибок B4: отказ approval — это `APPROVAL_DENIED`;
   `WORKBOOK_READ_ONLY` поднимается только для read-only-сессии; `COM_BUSY`/`RPC_E_CALL_REJECTED` —
   только при реальной конкуренции за COM (и конфликте ревизий VBA); `MACRO_BLOCKED` — только для
   Office macro-security / не-execute контекстов. В таблицах инструментов `COM_BUSY` у mutating-
   строки означает конкуренцию или конфликт ревизий, а не отказ approval.
2. **Сначала `excel_open`.** Любой инструмент, которому нужен `workbook_id`, требует
   предварительного открытия книги. `workbook_id` — это идентификатор рабочей копии,
   выданный сервером, а не путь к файлу.
3. **Снапшот перед мутацией.** Мутирующие инструменты требуют успешного снапшота
   (`snapshot_create`); иначе — `SNAPSHOT_FAILED`.
4. **Границы.** Чтение диапазонов и таблиц ограничено (`offset`/`limit`, флаг
   `truncated`); запросы сверх лимита отклоняются, а не «молча» урезаются.
5. **Никакого GUI.** Инструменты не активируют окна, не показывают диалоги и не
   отправляют клавиши.

## Классы доступа

| Класс | Кол-во | Поведение |
|-------|--------|-----------|
| `read-only` | 43 | Разрешён автоматически; не меняет книгу. Может писать файлы снапшотов, экспорта, артефактов и in-memory состояние классификации. |
| `mutating` | 40 | Guarded: writable-сессия → снапшот → approval → проверки → postcondition + аудит. |
| `executing` | 3 | Выполняет код (`macro_run`, `test_run_suite`, `test_suite_run`); дополнительно гейтится approval (в режиме `restricted`) и `background_safe`. |
| `mixed` | 2 | `excel_open` и `excel_close`: класс зависит от параметров (`read_only`, `mode`, `save_changes`). |

Важно различать «инструмент реализован» и «инструмент доступен через MCP». Все 40
`mutating` и 3 `executing` инструмента существуют и описаны в справочнике. В режиме `full`
(по умолчанию) approval-гейт их авторизует (снапшот перед мутацией всё равно требуется); только в
режиме `restricted` они отклоняются с `APPROVAL_DENIED`. Требование `background_safe` для
`macro_run` действует в обоих режимах.

## Сокращения

| Сокращение | Значение |
|------------|----------|
| `STA` | Single-Threaded Apartment — модель COM-апартамента, в которой живёт Excel. |
| `DTO` | Data Transfer Object — типизированная структура без COM-ссылок. |
| `VBOM` | VBA project object model — объектная модель проекта VBA (требует доверия в Trust Center). |
| `ListObject` | Excel Table — структурированная таблица листа. |
| `BG06` | Правило классификации `background_safe` для фонового выполнения макросов. |
| `R13` | Preflight интерактивного UI (без учёта регистра): `MsgBox`, `InputBox`, `.Show`, `FileDialog`, `Shell`, `CreateObject`, `FollowHyperlink`, `SendKeys`, `DoEvents`, `GetOpenFilename`, `GetSaveAsFilename`, `CallByName`, `ExecuteGlobal`, `Execute`, `Eval`, `OnTime`, `DisplayAlerts`, `.Run`, `CommandBars`. |
| `VBE` | VBA Editor — среда разработки VBA. |
| `UIA` | UI Automation — доступ к дереву элементов Windows (только чтение, без активации). |

## Сводка по группам

| Группа | Кол-во | Инструменты | Класс доступа |
|--------|--------|-------------|----------------|
| Workbook/session | 6 | `excel_open`, `excel_close`, `excel_save`, `excel_get_status`, `excel_list_workbooks`, `excel_health` | mixed / mutating / read-only |
| Snapshots | 4 | `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff` | read-only + 1 mutating |
| Test runner | 3 | `test_run_suite`, `test_suite_run`, `test_get_run` | executing + read-only |
| VBA | 9 | `vba_project_info`, `vba_list_components`, `vba_list_procedures`, `vba_read_component`, `vba_export_project`, `vba_get_references`, `vba_search_code`, `vba_replace_component_source`, `vba_insert_code` | read-only + 2 mutating |
| Compile | 2 | `vba_compile`, `vba_get_compile_diagnostics` | read-only |
| Macros | 4 | `macro_list`, `macro_run`, `macro_get_last_result`, `macro_classify_background` | read-only + 1 executing |
| Sheets | 7 | `sheets_list`, `sheet_get`, `sheet_add`, `sheet_rename`, `sheet_copy`, `sheet_delete`, `sheet_set_visibility` | read-only + 5 mutating |
| Tables | 11 | `tables_list`, `table_get`, `table_read_rows`, `table_append_rows`, `table_insert_row`, `table_update_rows`, `table_delete_rows`, `table_add_column`, `table_delete_column`, `table_create`, `table_resize` | read-only + 8 mutating |
| PivotTables | 13 | `pivot_tables_list`, `pivot_table_get`, `pivot_cache_get`, `pivot_fields_list`, `pivot_table_refresh`, `pivot_cache_refresh`, `pivot_field_set_layout`, `pivot_field_set_position`, `pivot_field_set_aggregation`, `pivot_field_set_properties`, `pivot_field_get_filters`, `pivot_field_apply_filter`, `pivot_field_clear_filters` | read-only + 8 mutating |
| Ranges | 16 | `range_get`, `range_read_values`, `range_read_formulas`, `range_write_values`, `range_write_formulas`, `range_get_format`, `range_set_format`, `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge`, `range_get_validation`, `range_set_validation`, `range_get_conditional_formatting`, `range_set_conditional_formatting` | read-only + 10 mutating |
| Names | 5 | `names_list`, `name_get`, `name_create`, `name_update`, `name_delete` | read-only + 3 mutating |
| Shapes | 3 | `shapes_list`, `shape_get`, `shape_set_on_action` | read-only + 1 mutating |
| Controls | 3 | `controls_list`, `control_get`, `control_set_on_action` | read-only + 1 mutating |
| Workbook model | 1 | `workbook_get_model` | read-only |
| Health | 1 | `server_health` | read-only |
| **Итого** | **88** | — | 43 read-only, 40 mutating, 3 executing, 2 mixed |

Сумма по группам: 6 + 4 + 3 + 9 + 2 + 4 + 7 + 11 + 13 + 16 + 5 + 3 + 3 + 1 + 1 = **88**.

## Workbook/session

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `excel_open` | Открывает разрешённую книгу как скрытую рабочую копию. | `path:string, mode:SessionMode=execute?, read_only:boolean=false?` | `WorkbookInfo` | Копия файла в `work/`; запись в аудит | `PATH_NOT_ALLOWED`, `APPROVAL_DENIED` (restricted), `EXCEL_HUNG` | mixed |
| `excel_close` | Закрывает открытую книгу; по умолчанию без сохранения. | `workbook_id:string, save_changes:boolean=false?` | `SessionStatus` | Закрытие рабочей копии; удаление копии; запись в аудит | `APPROVAL_DENIED` (restricted, при `save_changes=true`), `EXCEL_HUNG` | mixed |
| `excel_save` | Явно сохраняет рабочую копию. | `workbook_id:string` | `WorkbookInfo` | Запись файла рабочей копии; запись в аудит | `APPROVAL_DENIED` (restricted), `EXCEL_HUNG`, `COM_BUSY` | mutating |
| `excel_get_status` | Возвращает состояние сессии/книги без обращения к COM. | `workbook_id:string?` | `SessionStatus` | — | — | read-only |
| `excel_list_workbooks` | Список открытых книг из bookkeeping. | — | `list[WorkbookInfo]` | — | — | read-only |
| `excel_health` | Состояние воркера/сессии без принудительного перезапуска и без COM. | — | `ExcelHealth` | — | — (всегда payload) | read-only |

`excel_open` — `mixed`: `mode="inspect"` + `read_only=true` проходит как `READ_ONLY`;
`mode="inspect"` + `read_only=false` — `MUTATING`; `mode="execute"` — `EXECUTING`.
`excel_close` — `read-only` при `save_changes=false` и `mutating` при `save_changes=true`.

## Snapshots

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `snapshot_create` | Создаёт иммутабельный снапшот открытой или известной книги. | `workbook_id:string` | `SnapshotManifest` | Файлы снапшота + `manifest.json`; запись в аудит | `SNAPSHOT_FAILED`, `EXCEL_HUNG` | read-only |
| `snapshot_list` | Список снапшотов и их метаданных. | — | `list[SnapshotManifest]` | — | `SNAPSHOT_FAILED` | read-only |
| `snapshot_restore` | Восстанавливает рабочую сессию из снапшота. | `snapshot_id:string, workbook_id:string?` | `WorkbookInfo` | Закрытие текущей копии без сохранения; новая рабочая копия; аудит | `SNAPSHOT_FAILED`, `APPROVAL_DENIED` (restricted), `EXCEL_HUNG` | mutating |
| `snapshot_diff` | Сравнивает два снапшота по метаданным и хэшу. | `snapshot_id_a:string, snapshot_id_b:string` | `SnapshotDiff` | — | `SNAPSHOT_FAILED` | read-only |

## Test runner

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `test_run_suite` | Запускает декларативный тестовый набор на одноразowych копиях книги. | `suite:TestSuiteSpec` | `TestRunResult` | Базовый снапшот, рабочие копии на каждый кейс, артефакты, `runs/`, аудит | `INVALID_RANGE`, `APPROVAL_DENIED` (restricted), `SNAPSHOT_FAILED`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `EXCEL_HUNG` | executing |
| `test_suite_run` | Явная точка входа: принимает набор объектом или JSON-подобным mapping. | `suite:TestSuiteSpec` или `object` | `TestRunResult` | То же, что `test_run_suite` | `INVALID_RANGE`, `APPROVAL_DENIED` (restricted), `SNAPSHOT_FAILED`, `BACKGROUND_SAFETY_NOT_VERIFIED` | executing |
| `test_get_run` | Возвращает сохранённый прогон по `run_id`. | `run_id:string` | `TestRunResult` | — | `ExcelVbaMcpError` (не найден) | read-only |

Оба запускающих инструмента открывают рабочую копию в режиме `execute`. В режиме `full`
(по умолчанию) это авторизовано; в режиме `restricted` каждый кейс получает `blocked`
(`APPROVAL_DENIED`), а набор — `blocked`. Шаг `macro_run` по неклассифицированной процедуре
блокируется с `BACKGROUND_SAFETY_NOT_VERIFIED` в любом режиме.

## VBA

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `vba_project_info` | Сводка проекта VBA: имя, защита, компоненты, ссылки. | `workbook_id:string` | `VbaProjectInfo` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `EXCEL_HUNG` | read-only |
| `vba_list_components` | Список компонентов с типом, числом строк и хэшем. | `workbook_id:string` | `list[VbaComponentInfo]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_list_procedures` | Процедуры одного компонента через VBIDE CodeModule. | `workbook_id:string, component:string` | `list[VbaProcedureInfo]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_read_component` | Читает модуль компонента (код + нормализованный хэш). | `workbook_id:string, component:string` | `VbaComponentSource` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_export_project` | Экспортирует все компоненты и возвращает манифест. | `workbook_id:string, out_dir:string?` | `VbaExportManifest` | Файлы `.bas`/`.cls`/`.frm`/`.frx`; артефакт-корень; аудит | `ACCESS_VBOM_DISABLED`, `PATH_NOT_ALLOWED`, `SNAPSHOT_FAILED` | read-only |
| `vba_get_references` | Ссылки проекта с Name/GUID/version/path. | `workbook_id:string` | `list[VbaReferenceInfo]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_search_code` | Поиск по коду всех компонентов (подстрока или regex). | `workbook_id:string, pattern:string, regex:boolean=false?` | `list[VbaSearchMatch]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_replace_component_source` | Полностью заменяет модуль Standard/Class. | `workbook_id:string, component:string, new_source:string, expected_source_hash:string` | `VbaMutationResult` | Изменение проекта VBA; экспорт до/после; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `APPROVAL_DENIED` (restricted), `COM_BUSY` (конфликт ревизии), `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | mutating |
| `vba_insert_code` | Вставляет код в семантическую позицию Standard/Class модуля. | `workbook_id:string, component:string, code:string, position:string, expected_source_hash:string` | `VbaMutationResult` | Изменение проекта VBA; аудит | те же; неверная `position` — `COM_BUSY` | mutating |

`vba_replace_component_source` и `vba_insert_code` работают только со Standard/Class
модулями; Document-модули и UserForms отклоняются с `COM_BUSY`
(`reason=non_replaceable_component`).

## Compile

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `vba_compile` | Fail-closed результат компиляции; на валидированной сборке всегда `indeterminate`. | `workbook_id:string` | `CompileResult` | — (не запускает UI) | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_get_compile_diagnostics` | Последние диагностические данные без повторного запуска. | `workbook_id:string` | `CompileDiagnostics` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |

## Macros

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `macro_list` | Список вызываемых макросов с классификацией `background_safe`. | `workbook_id:string` | `list[MacroInfo]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `EXCEL_HUNG` | read-only |
| `macro_run` | Запускает один одобренный `background_safe` макрос. | `workbook_id:string, macro_name:string, args:array?` | `MacroRunResult` | Исполнение VBA; возможные внешние эффекты; аудит | `APPROVAL_DENIED` (restricted), `MACRO_BLOCKED` (Office macro-security), `MACRO_NOT_FOUND`, `MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED` | executing |
| `macro_classify_background` | Записывает `background_safe` процедуры для текущего хэша книги и ревизии исходников VBA (ungated bookkeeping). | `workbook_id:string, procedure:string, background_safe:boolean, expected_project_source_hash:string?` | `BackgroundSafetyClassification` | Запись в in-memory реестре; аудит | `MACRO_NOT_FOUND`, `STALE_TARGET` (несовпадение хэша исходников), `INTERACTIVE_UI_NOT_SUPPORTED` (интерактивная/эвристическая + `background_safe=true`), `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `EXCEL_HUNG` | read-only |
| `macro_get_last_result` | Возвращает кэш последнего запуска без повторного выполнения. | `workbook_id:string` | `MacroRunResult` | — | — | read-only |

`macro_classify_background` **не гейтится**: он только записывает состояние классификации и не
является границей безопасности — `macro_run` сохраняет approval-гейт и детерминированный вето R13.
Классификация привязана к `(workbook_file_sha256, project_source_hash, procedure)` и
инвалидируется при изменении исходников VBA (поздний `macro_run`/`macro_list` пересчитывает живой
`project_source_hash`; несовпадение означает, что процедура снова не классифицирована и
возвращает `BACKGROUND_SAFETY_NOT_VERIFIED`). Устаревший `expected_project_source_hash`
отклоняет запись с `STALE_TARGET`.

## Sheets

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `sheets_list` | Все листы книги с типом, видимостью, `code_name`, защитой. | `workbook_id:string` | `list[SheetInfo]` | — | — | read-only |
| `sheet_get` | Один лист по имени (учитывает `Sheets`/`Worksheets`). | `workbook_id:string, name:string` | `SheetInfo` | — | `COM_BUSY` (лист не найден) | read-only |
| `sheet_add` | Добавляет рабочий лист (MVP: только `worksheet`). | `workbook_id:string, name:string, sheet_type:string=worksheet?` | `SheetInfo` | Новая страница в книге; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `sheet_rename` | Переименовывает лист; `CodeName` не меняется. | `workbook_id:string, name:string, new_name:string` | `SheetInfo` | Имя листа; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `sheet_copy` | Копирует лист и возвращает DTO нового листа. | `workbook_id:string, name:string` | `SheetInfo` | Новый лист; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `sheet_delete` | Удаляет лист (разрушительно). | `workbook_id:string, name:string` | `SheetInfo` | Удаление листа; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `sheet_set_visibility` | Задаёт видимость листа. | `workbook_id:string, name:string, visible:boolean` | `SheetInfo` | Видимость листа; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

`sheet_delete` отклоняется, если после удаления не останется ни одного видимого листа;
`sheet_set_visibility` отклоняет скрытие последнего видимого листа.

## Tables

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `tables_list` | Все таблицы (ListObject) книги с диапазонами и колонками. | `workbook_id:string` | `list[TableInfo]` | — | — | read-only |
| `table_get` | Одна таблица по имени/`DisplayName`. | `workbook_id:string, name:string` | `TableInfo` | — | `COM_BUSY` (не найдена) | read-only |
| `table_read_rows` | Типизированный ограниченный срез строк таблицы. | `workbook_id:string, name:string, offset:integer=0?, limit:integer?` | `TableRows` | — | `COM_BUSY` | read-only |
| `table_append_rows` | Добавляет строки через `ListRows.Add` (структура сохраняется). | `workbook_id:string, name:string, rows:array<object>` | `TableMutationResult` | Строки таблицы; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_insert_row` | Вставляет одну строку в 1-based позицию `at`. | `workbook_id:string, name:string, at:integer, values:object?` | `TableMutationResult` | Строка таблицы; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_update_rows` | Обновляет строки по индексу или однозначному `where`. | `workbook_id:string, name:string, updates:array<object>` | `TableMutationResult` | Значения строк; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_delete_rows` | Удаляет строки по индексам или однозначному `where`. | `workbook_id:string, name:string, indices:array?, where:object?` | `TableMutationResult` | Строки таблицы; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_add_column` | Добавляет колонку через `ListColumns.Add(Position)`. | `workbook_id:string, name:string, header:string, position:integer?` | `TableMutationResult` | Колонка таблицы; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_delete_column` | Удаляет колонку по уникальному заголовку. | `workbook_id:string, name:string, header:string` | `TableMutationResult` | Колонка таблицы; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_create` | Создаёт ListObject на листе над A1-диапазоном или именем. | `workbook_id:string, sheet:string, range_or_name:string, style:string?` | `TableMutationResult` | Новая таблица; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_resize` | Изменяет размер таблицы через `ListObject.Resize`. | `workbook_id:string, name:string, new_range:string` | `TableMutationResult` | Диапазон таблицы; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

## PivotTables

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `pivot_tables_list` | Все сводные таблицы и их кэши по всем листам. | `workbook_id:string` | `PivotInventory` | — | `COM_BUSY` | read-only |
| `pivot_table_get` | Метаданные одной сводной таблицы по `Name`. | `workbook_id:string, pivot_table:string` | `PivotTableInfo` | — | `COM_BUSY` | read-only |
| `pivot_cache_get` | Метаданные PivotCache по 1-based `CacheIndex`. | `workbook_id:string, cache_index:integer` | `PivotCacheInfo` | — | `COM_BUSY` | read-only |
| `pivot_fields_list` | Поля сводной таблицы с ориентацией и агрегацией. | `workbook_id:string, pivot_table:string` | `list[PivotFieldInfo]` | — | `COM_BUSY` | read-only |
| `pivot_table_refresh` | Обновляет одну сводную таблицу (`RefreshTable`). | `workbook_id:string, pivot_table:string` | `PivotRefreshResult` | Данные сводной таблицы; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `REFRESH_FAILED`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_cache_refresh` | Обновляет один PivotCache (`PivotCache.Refresh`). | `workbook_id:string, cache_index:integer` | `PivotRefreshResult` | Данные всех таблиц кэша; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `REFRESH_FAILED`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_set_layout` | Задаёт ориентацию поля или удаляет его. | `workbook_id:string, pivot_table:string, field:string, orientation:PivotFieldOrientation, delete:boolean=false?` | `PivotFieldInfo` | Структура сводной таблицы; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_set_position` | Задаёт 1-based позицию поля в его ориентации. | `workbook_id:string, pivot_table:string, field:string, position:integer` | `PivotFieldInfo` | Порядок полей; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_set_aggregation` | Добавляет или обновляет агрегацию data-поля. | `workbook_id:string, pivot_table:string, field:string, function:PivotAggregation, caption:string?` | `PivotFieldInfo` | Data-поле; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_set_properties` | Задаёт `caption` / `number_format` / `show_all_items`. | `workbook_id:string, pivot_table:string, field:string, properties:PivotFieldProperties` | `PivotFieldInfo` | Свойства поля; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_get_filters` | Текущее состояние фильтров поля. | `workbook_id:string, pivot_table:string, field:string` | `PivotFieldFilters` | — | `AMBIGUOUS_FILTER`, `COM_BUSY` | read-only |
| `pivot_field_apply_filter` | Применяет ручной фильтр по видимым элементам. | `workbook_id:string, pivot_table:string, field:string, spec:PivotFilterSpec` | `PivotFieldFilters` | Видимость элементов; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_clear_filters` | Очищает фильтры поля или всей таблицы. | `workbook_id:string, pivot_table:string, field:string?, scope:string=field?` | `PivotFieldFilters` | Фильтры; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

Создание и удаление PivotTables/PivotCaches не поддерживается. `PivotFilters.Add`
(label/value/date/top-bottom) на валидированной сборке не работает: возвращается явный
`unsupported` результат с `filter.supported=false` и причиной, а не «тихая» конвертация в
ручной фильтр.

## Ranges

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `range_get` | Метаданные A1-диапазона (адрес, размеры, форма). | `workbook_id:string, sheet:string, address:string` | `RangeAddress` | — | `INVALID_RANGE` | read-only |
| `range_read_values` | Типизированные значения ограниченного окна. | `workbook_id:string, sheet:string, address:string, offset:integer=0?, limit:integer?` | `RangeValues` | — | `INVALID_RANGE` | read-only |
| `range_read_formulas` | Типизированные формулы ограниченного окна. | `workbook_id:string, sheet:string, address:string, offset:integer=0?, limit:integer?` | `RangeFormulas` | — | `INVALID_RANGE` | read-only |
| `range_write_values` | Запись типизированных значений в ограниченный диапазон. | `workbook_id:string, sheet:string, address:string, payload:RangeWriteRequest` | `RangeMutationResult` | Значения ячеек; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `LISTOBJECT_INTERSECTION` | mutating |
| `range_write_formulas` | Запись формул в ограниченный диапазон. | `workbook_id:string, sheet:string, address:string, payload:FormulaWriteRequest` | `RangeMutationResult` | Формулы ячеек; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `LISTOBJECT_INTERSECTION` | mutating |
| `range_get_format` | Три-state форматирование диапазона. | `workbook_id:string, sheet:string, address:string` | `FormatInfo` | — | `INVALID_RANGE` | read-only |
| `range_set_format` | Частичное обновление форматирования. | `workbook_id:string, sheet:string, address:string, patch:RangeFormatPatch` | `FormatMutationResult` | Форматирование; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `PROTECTED_TARGET` | mutating |
| `range_copy` | Копирование диапазона в другой лист/адрес (без буфера). | `workbook_id:string, sheet:string, source:string, destination_sheet:string, destination:string` | `StructuralMutationResult` | Данные получателя; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `LISTOBJECT_INTERSECTION` | mutating |
| `range_insert` | Вставка ячеек со сдвигом `right`/`down`. | `workbook_id:string, sheet:string, address:string, shift:string, allow_merged:boolean=false?` | `StructuralMutationResult` | Сдвиг ячеек; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | mutating |
| `range_delete` | Удаление ячеек со сдвигом `left`/`up`. | `workbook_id:string, sheet:string, address:string, shift:string, allow_merged:boolean=false?` | `StructuralMutationResult` | Удаление ячеек; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | mutating |
| `range_merge` | Объединение диапазона в одну ячейку. | `workbook_id:string, sheet:string, address:string, across:boolean=false?, allow_data_loss:boolean=false?` | `StructuralMutationResult` | Объединение; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `LISTOBJECT_INTERSECTION` | mutating |
| `range_unmerge` | Снимает объединение всех пересечённых областей. | `workbook_id:string, sheet:string, address:string` | `StructuralMutationResult` | Разъединение; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `LISTOBJECT_INTERSECTION` | mutating |
| `range_get_validation` | Правило проверки данных диапазона. | `workbook_id:string, sheet:string, address:string` | `ValidationRule` | — | `INVALID_RANGE` | read-only |
| `range_set_validation` | Добавляет или изменяет правило проверки данных. | `workbook_id:string, sheet:string, address:string, spec:ValidationSpec` | `ValidationRule` | Правило; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `PROTECTED_TARGET` | mutating |
| `range_get_conditional_formatting` | Правила условного форматирования по приоритету. | `workbook_id:string, sheet:string, address:string` | `list[ConditionalFormatRule]` | — | `INVALID_RANGE` | read-only |
| `range_set_conditional_formatting` | Добавляет, изменяет или удаляет правило. | `workbook_id:string, sheet:string, address:string, spec:ConditionalFormatSpec, mode:string=add?` | `ConditionalFormatMutationResult` | Правила; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `PROTECTED_TARGET` | mutating |

Структурные операции (`range_insert`, `range_delete`, `range_copy`, `range_merge`),
пересекающие Excel Table (ListObject), отклоняются с `LISTOBJECT_INTERSECTION` и должны
выполняться через инструменты Tables.

## Names

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `names_list` | Все именованные диапазоны книги и областей листов. | `workbook_id:string` | `list[DefinedNameInfo]` | — | — | read-only |
| `name_get` | Одно имя по имени и явной области. | `workbook_id:string, name:string, scope:string, sheet:string?` | `DefinedNameInfo` | — | `COM_BUSY` | read-only |
| `name_create` | Создаёт имя в явной области. | `workbook_id:string, name:string, refers_to:string, scope:string, sheet:string?` | `DefinedNameInfo` | Новое имя; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `name_update` | Обновляет `RefersTo` (область сохраняется). | `workbook_id:string, name:string, refers_to:string, scope:string, sheet:string?` | `DefinedNameInfo` | Имя; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `name_delete` | Удаляет имя и возвращает его предыдущее значение. | `workbook_id:string, name:string, scope:string, sheet:string?` | `DefinedNameInfo` | Имя; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

## Shapes

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `shapes_list` | Все фигуры книги с типом, `on_action`, позицией и текстом. | `workbook_id:string` | `list[ShapeInfo]` | — | — | read-only |
| `shape_get` | Одна фигура по имени (регистронезависимо). | `workbook_id:string, name:string` | `ShapeInfo` | — | `COM_BUSY` | read-only |
| `shape_set_on_action` | Назначает макрос фигуре (макрос не выполняется). | `workbook_id:string, name:string, macro_name:string` | `ShapeInfo` | `OnAction`; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

Пустой `macro_name` снимает назначение.

## Controls

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `controls_list` | Все Form Control и ActiveX-элементы книги. | `workbook_id:string` | `list[ControlEntry]` | — | — | read-only |
| `control_get` | Один элемент по имени (регистронезависимо). | `workbook_id:string, name:string` | `ControlEntry` | — | `COM_BUSY` | read-only |
| `control_set_on_action` | Назначает макрос Form Control (не ActiveX). | `workbook_id:string, name:string, macro_name:string` | `ControlEntry` | `OnAction`; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

ActiveX-элементы отклоняются: у них нет `OnAction`-семантики Form Control.

## Workbook model

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `workbook_get_model` | Агрегированная модель книги: листы, таблицы, имена, фигуры, элементы, VBA. | `workbook_id:string, include_vba:boolean=true?` | `WorkbookModel` | — | — (состояние VBA — в `vba_access`) | read-only |

Модель не читает «все ячейки» и не выполняет VBA; при недоступном VBOM остальные секции
всё равно возвращаются, а `vba_access` сообщает состояние.

## Health

| Инструмент | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | Доступ |
|------------|------------|-----------|---------|------------------|--------|--------|
| `server_health` | Read-only liveness-проба; никогда не обращается к Excel/COM. | — | `{"status", "service", "version"}` | — | — | read-only |

`excel_health` сгруппирован с Workbook/session, поскольку описывает состояние сессии и
воркера; группа Health содержит только liveness-пробу сервера.

## Подробно о сложных инструментах

### excel_open

`excel_open` по умолчанию использует `mode="execute", read_only=False` — сессия для записи и
выполнения. Передайте `mode="inspect"` и/или `read_only=True` для сессии с отключёнными макросами
или только для чтения. Read-only inspect-режим:

```json
{ "path": "C:/data/report.xlsm", "mode": "inspect", "read_only": true }
```

Ожидаемая форма ответа:

```json
{
  "workbook_id": "1f0c…",
  "name": "report.xlsm",
  "format": "xlsm",
  "path": "<temp>/excel_vba_mcp/work/1f0c….xlsm",
  "source_path": "C:/data/report.xlsm",
  "read_only": true,
  "has_vba": true,
  "excel_version": "16.0.14332",
  "session_id": "9b2a…",
  "mode": "inspect"
}
```

writable/execute-открытие (по умолчанию, если `mode`/`read_only` не заданы):

```json
{ "path": "C:/data/report.xlsm", "mode": "execute", "read_only": false }
```

В режиме `full` (по умолчанию) это открытие авторизовано и рабочая копия создаётся. В режиме
`restricted` оно классифицируется как `MUTATING`/`EXECUTING` и **отклоняется** с
`code="APPROVAL_DENIED"`; книга не открывается, рабочая копия не создаётся, поскольку
grant-инструмент через MCP отсутствует.

### excel_close / excel_save

```json
{ "workbook_id": "1f0c…", "save_changes": false }
```

`excel_close` без сохранения — read-only и разрешён; ответ — `SessionStatus` с
`state="closed"`. Сохранение при закрытии — мутация:

```json
{ "workbook_id": "1f0c…", "save_changes": true }
```

Авторизовано в режиме `full` (по умолчанию); в режиме `restricted` отклоняется с
`APPROVAL_DENIED`. Явное сохранение:

```json
{ "workbook_id": "1f0c…" }
```

Требует `MUTATING`-approval (авторизовано в `full`; `APPROVAL_DENIED` в `restricted`) и
возвращает обновлённый `WorkbookInfo` с `dirty=false`. Изменения не сохраняются неявно никогда.

### range_read_values / range_write_values

Чтение ограниченного окна:

```json
{
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "A1:C10",
  "offset": 0,
  "limit": 10
}
```

Ответ — `RangeValues`:

```json
{
  "range": { "sheet": "Sheet1", "address": "A1:C10", "rows": 10, "cols": 3 },
  "shape": "block",
  "rows": 10,
  "cols": 3,
  "values": [
    [ { "kind": "string", "value": "Header" }, { "kind": "empty" }, { "kind": "empty" } ]
  ],
  "truncated": false,
  "notes": []
}
```

Значения типизированы: `string`, `number`, `boolean`, `date`, `error`, `empty`; потеря
точности при преобразовании в строку исключена. Запрос сверх лимита ячеек отклоняется,
а не урезается; `truncated=true` появляется, когда окно `offset`/`limit` обрезано.

Запись требует точного прямоугольника `payload`:

```json
{
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "B2:C3",
  "payload": {
    "rows": 2,
    "cols": 2,
    "values": [
      [ { "kind": "number", "value": 1 }, { "kind": "string", "value": "=1+1" } ],
      [ { "kind": "boolean", "value": true }, { "kind": "empty" } ]
    ]
  }
}
```

Строка, начинающаяся с `=`, сохраняется как **текст**, а не как формула (для формул есть
отдельный инструмент). Неверный размер payload отклоняется с `INVALID_RANGE` без изменений.
Ответ — `RangeMutationResult` с `range_before`, `range_after`, `verified` и
`postcondition_values` (перечитанные значения). Мутация требует writable-сессии и снапшота;
approval-гейт авторизует её в режиме `full` (по умолчанию) и отклоняет с `APPROVAL_DENIED` в
`restricted`.

### range_set_format

Разреженная (sparse) запись: применяются только явно заданные поля `patch`.

```json
{
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "A1:C1",
  "patch": {
    "font_bold": true,
    "font_size": 11,
    "fill_color": 15773696,
    "number_format": "0.00",
    "alignment_horizontal": "center"
  }
}
```

Отсутствующие поля остаются без изменений; `null` сбрасывает значение в автоматическое
только там, где это допускает Excel. Ответ:

```json
{
  "action": "set_format",
  "range": { "sheet": "Sheet1", "address": "A1:C1", "rows": 1, "cols": 3 },
  "applied": ["font.bold", "number_format"],
  "ignored": [],
  "verified": true,
  "result_format": { "font": { "bold": true }, "number_format": "0.00" },
  "notes": []
}
```

### table_update_rows / table_delete_rows

Обновление допускает **явный индекс** или **однозначный `where`**:

```json
{
  "workbook_id": "1f0c…",
  "name": "Sales",
  "updates": [
    { "index": 2, "values": { "Amount": 120 } },
    { "where": { "Region": "North" }, "values": { "Flag": "Y" } }
  ]
}
```

Если `where` не разрешается в ровно одну строку, инструмент возвращает `COM_BUSY` с
сообщением о неоднозначности и предпросмотром найденных строк — автоматического
«угадывания» нет.
Ответ — `TableMutationResult` с `before`/`after` (`TableInfo`) и `affected_rows`.

Удаление строк:

```json
{ "workbook_id": "1f0c…", "name": "Sales", "indices": [3, 5] }
```

или через `where`; отсутствие обоих аргументов — ошибка валидации. Удаление
и обновление требуют writable-сессии и снапшота; approval-гейт авторизует их в режиме `full`
(по умолчанию) и отклоняет с `APPROVAL_DENIED` в `restricted`.

### snapshot_create / snapshot_restore

```json
{ "workbook_id": "1f0c…" }
```

Ответ — `SnapshotManifest`:

```json
{
  "snapshot_id": "snap-…",
  "created_at": "2026-11-21T18:41:56Z",
  "source_path": "C:/data/report.xlsm",
  "file_name": "report.xlsm",
  "stored_path": "<snapshots>/snap-…/report.xlsm",
  "sha256": "…",
  "size": 48213,
  "snapshot_method": "SaveCopyAs",
  "external_side_effects_not_undone": true
}
```

Восстановление:

```json
{ "snapshot_id": "snap-…", "workbook_id": "1f0c…" }
```

Текущая рабочая копия закрывается **без сохранения**, затем открывается новая копия из
baseline; возвращается `WorkbookInfo`. Внешние эффекты макросов не откатываются.

Обратите внимание, что `snapshot_create` и `vba_export_project` записывают файлы: снапшоты — в
temp-каталог снапшотов, экспортируемые компоненты — в артефакт-/temp-каталог. Эти записи
выполнены намеренно; они не требуют approval в режиме `full` и не гейтятся в режиме
`restricted` (R-19).

### vba_replace_component_source

Перед заменой нужны два вызова: `vba_list_components` (убедиться, что компонент
Standard/Class) и `vba_read_component` (получить `source_hash`). Затем:

```json
{
  "workbook_id": "1f0c…",
  "component": "Module1",
  "new_source": "Option Explicit\n\nPublic Sub Hello()\nEnd Sub",
  "expected_source_hash": "<source_hash из vba_read_component>"
}
```

`expected_source_hash` — защита от устаревшей ревизии: сервер сравнивает его с текущим
нормализованным хэшем до и **внутри** STA-воркера. Несовпадение — `COM_BUSY` с
`context.reason="revision_conflict"`; мутация не выполняется. Ответ — `VbaMutationResult`
с `before_hash`, `after_hash` и путями экспорта. Document-модули и UserForms
отклоняются.

### macro_run

```json
{ "workbook_id": "1f0c…", "macro_name": "Module1.Hello", "args": [] }
```

Порядок гейтов:

1. execute-режим + writable-сессия (`_require_execute_context`);
2. разрешение процедуры из инвентаря (`Module.Procedure`);
3. approval (ключ: canonical path + SHA-256 + действие + процедура) — авторизован в режиме
   `full`, `APPROVAL_DENIED` в `restricted`;
4. `background_safe` для текущего хэша книги **и** `project_source_hash` (BG06);
5. preflight интерактивного UI (R13): `MsgBox`, `InputBox`, `.Show`, `FileDialog`, `Shell`,
   `CreateObject`, `FollowHyperlink`, `SendKeys`, `DoEvents`, `GetOpenFilename`,
   `GetSaveAsFilename`, `CallByName`, `ExecuteGlobal`, `Execute`, `Eval`, `OnTime`,
   `DisplayAlerts`, `.Run`, `CommandBars` (без учёта регистра);
6. один `Application.Run` на STA-воркере.

Возможные статусы `MacroRunResult`: `success`, `not_found`, `runtime_error`, `blocked`,
`timeout`, `needs_recovery`. При блокировке макрос **не запускается**: `blocked` с
`APPROVAL_DENIED` (approval в restricted), `BACKGROUND_SAFETY_NOT_VERIFIED` (не
классифицирована для текущего хэша книги и ревизии исходников) или
`INTERACTIVE_UI_NOT_SUPPORTED` (опасный токен, классификация «небезопасен» или эвристическая
процедура). Таймаут с модальным окном, принадлежащим owned PID, даёт `MACRO_RUNTIME_ERROR` и
`dismissed=false`; без модального окна — `MACRO_TIMEOUT`/`needs_recovery`. Повторного
«слепого» запуска нет.

**Остаточная граница ложно-безопасного.** Токен-скан R13 — эвристический запрет, а не
статическое доказательство. Динамическая/косвенная UI (`CallByName`/`ExecuteGlobal`/`Eval`) и
UI, достижимый транзитивно через другие процедуры, полностью не покрываются. Список запрещённых
токенов теперь включает также токены окна/фокуса и видимости (`AppActivate`, `Application.Visible`,
`.Activate`, `WindowState`), но процедура, достигающая UI только косвенно, всё ещё может обойти
подстроковый скан. Классификация `background_safe` — результат человеческой проверки, а
детерминированный слой может только запретить, но не удостоверить.

### macro_classify_background

Классификация процедуры `background_safe` для текущего хэша книги и ревизии исходников VBA:

```json
{
  "workbook_id": "1f0c…",
  "procedure": "Module1.Hello",
  "background_safe": true
}
```

Необязательный `expected_project_source_hash` — защита от устаревшей ревизии: если он больше
не совпадает с живым хэшем исходников, запись отклоняется с `STALE_TARGET`. Классификация
эвристической или интерактивной процедуры как `background_safe` отклоняется с
`INTERACTIVE_UI_NOT_SUPPORTED` (вето R13). Результат — `BackgroundSafetyClassification` с
`workbook_sha256`, `project_source_hash` и `classified_at`. Инструмент **не гейтится** (без
approval) и не является границей безопасности — `macro_run` по-прежнему применяет approval-гейт
и вето R13. Реестр in-memory, живёт в рамках процесса сервера; позднее изменение исходников VBA
инвалидирует классификацию.

### test_suite_run

```json
{
  "suite": {
    "suite": "smoke",
    "workbook_id": "1f0c…",
    "baseline": "create",
    "cases": [
      {
        "id": "reads-range",
        "steps": [
          {
            "action": "range_read_values",
            "sheet": "Sheet1",
            "range_address": "A1:B2",
            "limit": 2
          }
        ],
        "assertions": [
          {
            "type": "range_structure",
            "sheet": "Sheet1",
            "address": "A1:B2",
            "rows": 2,
            "cols": 2
          }
        ]
      }
    ]
  }
}
```

`action` и `type` берутся из **закрытых** словарей: `TestAction` (37 действий) и
`AssertionType` (29 проверок); схема запрещает лишние поля (`extra="forbid"`). Неизвестное
действие, произвольный код, shell-команда или неограниченный payload отклоняются до
любого обращения к Excel (`INVALID_RANGE`).

Изоляция: один иммутабельный baseline-снапшот, затем **своя одноразовая рабочая копия на
каждый кейс**; кейс не может «запачкать» следующий. Поддерживаются последовательности
edit→macro→edit→macro: шаг `macro_run` может нести необязательное поле только с данными
`background_safe`, и раннер классифицирует процедуру через защищённый путь
`MacroService.classify_background` перед запуском (напрямую реестр не трогает). Компиляция
перед макросом явная — нужен шаг `vba_compile`; неудачный шаг `vba_compile` блокирует
последующие `macro_run` шага кейса, а шаг `macro_run` по неклассифицированной процедуре
блокируется с `BACKGROUND_SAFETY_NOT_VERIFIED`. `indeterminate` компиляции не считается
ошибкой компиляции. Возврат — `TestRunResult` с per-case/per-step статусами; прогон
сохраняется (`runs/`), читается через `test_get_run`.

### vba_compile

```json
{ "workbook_id": "1f0c…" }
```

Fail-closed ответ:

```json
{
  "workbook_id": "1f0c…",
  "status": "indeterminate",
  "reason": "NON_ACTIVATING_COMPILE_UNAVAILABLE",
  "diagnostics": {
    "is_compile": false,
    "vbom_available": true,
    "protection": 0,
    "locked": false,
    "component_count": 3,
    "broken_references": []
  },
  "command_discovery": { "route_usable": false },
  "notes": [
    "NOT a compile: non-activating VBA compilation is unavailable on this build.",
    "status=indeterminate reason=NON_ACTIVATING_COMPILE_UNAVAILABLE"
  ]
}
```

Инструмент никогда не возвращает `compiled` или `compile_error`, не показывает и не
активирует UI. Причины: `NON_ACTIVATING_COMPILE_UNAVAILABLE`, `ACCESS_VBOM_DISABLED`,
`VBA_PROJECT_LOCKED`. `vba_get_compile_diagnostics` возвращает те же данные из кэша без
повторного сбора.

## Зависимости между инструментами

1. **`excel_open` перед любым workbook-инструментом.** Все инструменты с `workbook_id`
   требуют открытой рабочей копии; иначе — `ExcelVbaMcpError` «unknown workbook_id».
2. **`snapshot_create` до любой guarded-мутации.** Без успешного снапшота мутация
   отклоняется с `SNAPSHOT_FAILED`.
3. **`vba_list_components` / `vba_read_component` до `vba_replace_component_source` и
   `vba_insert_code`.** Нужны имя компонента и актуальный `source_hash` для
   `expected_source_hash`.
4. **`vba_compile` до `vba_get_compile_diagnostics` — только чтобы «засеять» кэш.**
   Без первого вызова диагностика собирается заново; это не меняет результат
   (`is_compile=false`).
5. **`macro_list` до `macro_run`.** Позволяет узнать точное имя `Module.Procedure` и
   статус `background_safe` до запуска.
6. **`test_get_run` после `test_run_suite` / `test_suite_run`.** `run_id` берётся из
   `TestRunResult.test_run_id`; инструмент только читает сохранённый прогон.

Дополнительно: `excel_close` (без сохранения) и `excel_save` несовместимы с последующими
вызовами по тому же `workbook_id` — после закрытия идентификатор больше не действителен.

## Мутирующие и потенциально опасные операции

| Инструмент | Что меняется | Почему гейтится | Как использовать безопасно |
|------------|--------------|-----------------|---------------------------|
| `excel_open` (writable/execute) | Создаёт рабочую копию и открывает её с макросами (`execute`) | Открытие по умолчанию; авторизовано в `full`, `APPROVAL_DENIED` в `restricted` | Использовать `mode="inspect", read_only=true` для read-only работы; `execute` применять только к проверенным копиям |
| `excel_close(save_changes=true)` | Сохраняет рабочую копию на диск | `MUTATING` | Сохранять только после `snapshot_create` и осознанной проверки |
| `excel_save` | Записывает файл рабочей копии | `MUTATING` | Комбинировать со снапшотом; не сохранять «на всякий случай» |
| `snapshot_restore` | Отбрасывает текущую копию без сохранения и открывает baseline | `mutating` (включая writable-open) | Делать свежий `snapshot_create` перед восстановлением |
| `sheet_delete` | Удаляет лист с данными | Разрушительно, гейтится | Проверять `sheets_list`; полагаться на снапшот |
| `table_delete_rows` | Удаляет строки таблицы | Разрушительно, гейтится | Использовать явные `indices`; `where` только при однозначности |
| `table_delete_column` | Удаляет колонку таблицы | Разрушительно, гейтится | Проверять `table_get`; не передавать неоднозначный заголовок |
| `range_write_values` | Перезаписывает значения ячеек | `MUTATING`, гейтится | Проверять размер payload; делать снапшот |
| `range_write_formulas` | Перезаписывает формулы | `MUTATING`, гейтится | Не использовать для констант (для них — значения) |
| `range_delete` | Удаляет ячейки со сдвигом | Разрушительно, гейтится | Всегда задавать `shift` явно; проверять пересечение с таблицами |
| `range_merge` | Объединяет ячейки, отбрасывая данные вне якоря | Может потерять данные | Просматривать `preconditions` и передавать `allow_data_loss=true` только после ревью |
| `name_delete` | Удаляет определённое имя | Может сломать формулы | Проверять `names_list` и зависимости; полагаться на снапшот |
| `vba_replace_component_source` | Полностью заменяет код модуля | `MUTATING`, гейтится | Передавать точный `expected_source_hash`; проверять `vba_compile` после |
| `vba_insert_code` | Вставляет код в модуль | `MUTATING`, гейтится | Указывать точную `position`; не вставлять код «наугад» |
| `shape_set_on_action` | Назначает макрос фигуре | `MUTATING`, гейтится | Проверять `shapes_list`; пустой `macro_name` снимает назначение |
| `control_set_on_action` | Назначает макрос Form Control | `MUTATING`, гейтится | Работать только с Form Control; ActiveX не поддерживается |
| `macro_run` | Исполняет VBA-код | `EXECUTING`, гейтится approval (restricted) + `background_safe` + R13 | Только после `macro_list`/`macro_classify_background` и preflight; не запускать макросы с внешними эффектами |

## Read-only ресурсы

| URI template | Name | MIME | Содержимое |
|--------------|------|------|------------|
| `excel://workbooks/{workbook_id}/model` | `workbook_model` | `application/json` | Агрегированная модель книги Stage 1 (листы, таблицы, имена, фигуры, элементы, опционально VBA) |
| `excel://workbooks/{workbook_id}/vba/components` | `workbook_vba_components` | `application/json` | Инвентарь VBA-компонентов + `vba_access` |
| `excel://workbooks/{workbook_id}/vba/components/{component_name}` | `workbook_vba_component_source` | `text/plain` | Исходный текст одного VBA-компонента |
| `excel://test-runs/{run_id}` | `test_run_summary` | `application/json` | Сохранённый декларативный тестовый прогон |

Ресурсы доступны только для чтения: они не мутируют книгу и не выполняют VBA.

## Известные ограничения инструментов

1. **Компиляция всегда `indeterminate`.** `vba_compile` не может подтвердить компиляцию на
   валидированной сборке; `is_compile=false`.
2. **Макросы требуют классификации `background_safe`.** `BackgroundSafetyRegistry` стартует
   пустым (fail-closed); классифицируйте процедуру через `macro_classify_background`.
   Неклассифицированная процедура возвращает `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`);
   интерактивная блокируется вето R13 (`INTERACTIVE_UI_NOT_SUPPORTED`). Токен-скан R13 —
   эвристический запрет, а не доказательство.
3. **Нет grant-инструмента для approval.** Управление approval/allowlist через MCP невозможно.
   В режиме `full` (по умолчанию) мутации и execute-режим авторизованы без гранта; в
   `restricted` они отклоняются с `APPROVAL_DENIED`. Классификация `background_safe` доступна
   через `macro_classify_background`.
4. **Нет создания/удаления PivotTables и PivotCaches.** Только управление существующими
   объектами; внешние/OLAP/Data-Model источники отклоняются.
5. **Пересечение диапазона с ListObject.** Структурные операции диапазонов над таблицами
   отклоняются с `LISTOBJECT_INTERSECTION`; нужно использовать инструменты Tables.
6. **Ограниченные чтения.** `range_read_values`, `range_read_formulas`,
   `table_read_rows` возвращают ограниченные окна; большие диапазоны нужно читать
   порциями через `offset`/`limit`.
7. **Нет диаграмм, сводных диаграмм и слайсеров.** Инструменты для них не
   предоставляются.
8. **Нет создания книг.** `excel_open` открывает только существующие файлы; создать новую
   книгу через MCP нельзя.
9. **Нет полного пересчёта и cross-workbook копирования.** `Application.Calculate` не
   предоставляется; `range_copy` работает только внутри одной книги и без буфера обмена.
10. **Нет пагинации списков.** `tables_list` и `pivot_tables_list` возвращают полный
    ограниченный список без постраничного вывода.
11. **Устаревший воркер / потеря owned-Excel не детектируются автоматически через `excel_health` (R-17).**
    `excel_health` сообщает состояние воркера только из внутрипроцессного воркера; если
    принадлежащий скрытый процесс `EXCEL.EXE` умирает (например, убит извне), health может
    по-прежнему показывать `ok`/`running` с устаревшим `owned_pid`, пока вызов не завершится
    ошибкой. Нет in-band проверки живости PID и нет MCP-восстановления. Решение: `excel_close`
    для затронутого `workbook_id` и повторное открытие, либо перезапуск серверного процесса.
    Это известное ограничение, а не тихая порча данных.
12. **У `macro_run` нет предусловия снапшота (задокументированное отклонение от спеки).**
    В отличие от контент-мутирующих инструментов, `macro_run` не требует предварительного
    `snapshot_create`; он запускает классифицированный, одобренный макрос без него.
    Рекомендованный процесс всё равно делает снапшот первым, чтобы внешние эффекты макроса имели
    baseline для отката. Эта асимметрия — известное, задокументированное ограничение, а не
    гарантия того, что макросы защищены снапшотом.
13. **Восстановление сессии — только close/reopen; инструмента `excel_recover_session` нет.**
    `SessionService.recover_session`/`close_session` существуют внутри, но не выставлены как
    MCP-инструменты. `excel_health` показывает `needs_recovery`, но клиент не может запустить
    восстановление in-band; поддерживаемое решение — закрыть и повторно открыть книгу (или
    перезапустить сервер).
14. **`macro_policy.classify_operation` неавторитетен (мёртвая/иллюстративная таблица).**
    Функция `classify_operation` в `security/macro_policy.py` сохранена для иллюстрации и **не**
    подключена к пути рантайм-энфорсмента; живой класс операции каждого инструмента определён
    по-сервисному. Её таблица устарела (предшествует текущему каталогу из 88 инструментов) и не
    должна читаться как авторитетная классификация инструментов — используйте этот справочник
    или код сервисов.